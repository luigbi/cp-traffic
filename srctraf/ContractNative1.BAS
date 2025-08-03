Attribute VB_Name = "ContractNative1"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ContractNative1.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the extra Contract routines so that code can be compiled as Native code
Option Explicit
Option Compare Text
'10525
Public bgModelNTR
'Selection
Public tmStepCtrls(0 To 3)  As FIELDAREA
Public imStepBoxNo As Integer    'Current Step Box
Public imStep1Map(0 To 8) As Integer  'List index to actual index for lbcStep(0)
Public imStep1Index As Integer 'lbcStep(0) index
        'Order imStep1Index: 0=New Scratch; 1=New Model; 2=Revise Order; 3= Chg Hold to Order
        'Proposal          : 0=New Scratch; 1=New Model; 2=Chg Proposal; 3=New Version; 4=Revise Order; 5=Chg Hold to Order
Public imStep2Index As Integer 'lbcStep(1) index
Public imStep3Index As Integer 'lbcStep(2) index
Public imFirstTimeStep As Integer
Public imInSelect As Integer 'True=Processing a selection
Public imListField(0 To 3) As Integer
Public imResetSelection As Integer 'Used to indicate that Rate Bar dates should be set
Public imRBDatesSpotsRow As Integer
Public imRBDetailAnchorRow As Integer
Public imSetAll As Integer
Public imLoadingForm As Integer
Public imLoadingContract As Integer
Public smSpfUsingFeatures6 As String * 1  'Clear INVEXPORTPARAMETERS as this feature was for Channel One. It has been reused for the General Invoice Export.
                                        'By clearing that field, the Channel One questions "Project Code and SMS # will not show with the Order # area
'10256
Public imACT1CodesDefined As Integer
Public bmAct1PackageEnable As Boolean
Public imToggleAvail As Integer      '-1=None; 0=line length; 1=30/60; 2= 10s breaks
Public tmAvToggleCtrls(0 To 1) As FIELDAREA

Public lmAirTimeArrowBackColor As Long

'Contract header
Public tmChfSrchKey As LONGKEY0  'CHF key record image
Public tmChfSrchKey1 As CHFKEY1  'CHF key record image (contract #)
Public hmCHF As Integer        'CHF Handle
Public imCHFRecLen As Integer      'CHF record length
Public lmCPriDemoGRP As Long       'Primary Demo contract total GRP, set in mSetPriDemoResearch
Public lmCPriDemoGrImp As Long     'Primary Demo contract total GrImp, set in mSetPriDemoResearch
Public lmCPriDemoCPP As Long       'Primary Demo contract total CPP, set in mSetPriDemoResearch
Public lmCPriDemoCPM As Long       'Primary Demo contract total CPM, set in mSetPriDemoResearch
Public imCPriDemoRating As Integer 'Primary Demo contract total Avg rating, set in mSetPriDemoResearch
Public imMktFound As Integer   'Indicate if vehicle within market found with contract
Public imAllAnsw As Integer    'Used to indicate if all header questions anwswered
Public lmRevNoCode() As Long    'Contract Code for contract stored in RevNo list box
Public lmDiffCntrCode() As Long    'Contract Code for contract stored in difference list box
Public smOrigType As String
Public smChfStartDate As String
Public smChfEndDate As String
Public imRecLocked As Integer  'True=Recorded locked; False=Record not locked
Public lmLockRecCode As Long   'Lock Rec code
Public imCntrCluster As Integer 'True=Contract has vehicle with market cluster = "Y"- don't schedule contract
Public imCntrRep As Integer 'True=Contract has rep vehicle with market cluster = "N"- don't schedule contract
Public imCntrAir As Integer
Public imLnTotalType As Integer  '0=Show average price; 1=Show total price; 2=By Quarter and indicates which quarter (2=first quarter) or Show Research in price field (Avail tab or Protection Tab or Research Tab)
Public smTotalSpots As String  'Total number of airing spots defined for the contract
Public smTotalDollars As String 'Total dollars for the contract
Public smTotalNA As String * 40 'Line to be shown when total is N/A
Public imCppCpm As Integer   '0=N/a; 1= CPM; 2=CPP
Public imSvCppCpm As Integer   'Used to test if value chaged
Public imHideDemo As Integer   '0=No; 1= Yes
Public imSvHideDemo As Integer   'Used to test if value chaged
Public imAgyComm As Integer    'Agency commissionable on trades (0=Yes; 1=No:Default)
Public smCurRC As String
Public imCreditRestr As Integer    '-1=No restriction; 1= No new Orders
'Fixed Price and item billing
Public hmSbf As Integer        'Special billing
Public imAnyFixedBilled As Integer   'True = fixed price billed => disallow changing of billing cycle
'Contract Lines
Public tmClfSrchKey As CLFKEY0  'CLF key record image
Public tmClfSrchKey2 As LONGKEY0  'CLF key record image
Public hmClf As Integer        'CLF Handle
Public tmClf As CLF
Public tmChf As CHF    'Used when updating only to eliminate conflict
Public imClfRecLen As Integer      'CLF record length
Public imNoWkDays(0 To 6) As Integer

'5/19/12: Retain Daily line definition
Public imOvDailySpots(0 To 6) As Integer
Public smAlterHidden As String


Public imLnChg As Integer
Public imSvLnChg As Integer
Public imShowOV2 As Integer
Public imSv1stPos As Integer
Public imSvSolo As Integer
Public fmDragX As Single       'Start x location of drag
Public fmDragY As Single       'Start y location
Public imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Public imDragBoxNo As Integer
Public imAnyLnSchd As Integer   'True=line sch (can't change rate card); False=No lines scheduled
Public imLine As Integer   '0=8 lines, 1=13 lines
Public imLnDaily As Integer 'Air Week for Daily Buy:0=Yes; 1=No
Public smLineType As String    'S=Standard; O or A or E=Package (Parent); H=Hidden (Child)
Public smLineTypeOv As String  'S=Standard or P (Package)
Public imBypassLnPaint As Integer  'Bypass painting when forming packages
Public bmInAvailTab As Boolean
Public lmToolTipNext As Long
Public imToolTipWeekIndex As Integer
Public imExpandingPkg As Integer
Public bmInCancel As Integer
Public imInFlightScreen As Integer 'True= In flight screen- ignore tabs focus
Public bmSettingPriceInTab As Boolean
Public imInBuildWkInfo As Integer
Public imInResearch As Integer 'True=In Research-ignore changes
Public imInSummary As Integer
Public imPaintLnRowNo As Integer   'Line row number to be painted only (-1=all)
Public imCurrHightLightLnRowNo As Integer
Public imPrevHightLightLnRowNo As Integer
Public imBypassResetDemoBar As Integer
Public imInFlightReturn As Integer 'Aviod calling mFlightReturn multi-times, added on 11/03/2003.  fixed
                                'error when in cycle cmcAirCancel->mFlightReturn->mSetHdDate->lacPrice_click->aAllSetSDhow->mFlightReturn
                                'This only happened if input via Quarters
Public smSvAllowDailyWeekly As String
Public tmVehGp5Code() As SORTCODE
'Public smVehGp5CodeTag As String

Public tmRBBookname() As SORTCODE
Public smRBBooknameTag As String

Public tmRegionCode() As SORTCODE
Public smRegionCodeTag As String

'Contract Flights
Public tmDGCtrls(0 To 16) As FIELDAREA
Public imDGBoxNo As Integer
Public smDGSave(0 To 15) As String   'Values saved for Gen, Index:1=# Weeks, 2= # spots/week, 3-9=Daily spots or weekly flag, 10=Start Date, 11=End Date;
                                    '12=Price; 13=Override Start Time; 14=Override End Time, 15=ACT1 Lineup
Public imDGSave(0 To 11) As Integer   'Values saved for Gen, Index:1=Run every; 2=Daily(0)/Weekly(1)/Daily %(2)/Weekly%(3); 3=Preempt (0=Yes; 1=No) Not Used; 4 = Rank (priority); 5=BB Open Length; 6=BB Close Length; 7=Live Copy (0=Live Cmml; 1=Rec Cmml; 2=Live Promo; 3=Rec Promo)
                                    '8=Stored Times (0=Yes; 1=No); 9=Stored Spots (0=Yes; 1=No); 10=Stored Clear % (0=Yes; 1=No); 11=Daypart Filter (0=Yes; 1=No)
Public imInChkForStdPkg As Integer     'Set when mChkForStdPkg is called to aviod repainting each line in OV1INDEX as each hidden line is created
Public imAiredWk As Integer    'Flag that indicats if originally defined as aired week
Public imGenReq As Integer 'Indicates if tab on gen button allowed
Public hmCff As Integer        'CFF Handle
Public tmCff As CFF
Public imCffRecLen As Integer     'CFF record length
Public tmCffSrchKey1 As LONGKEY0
Public imStartDayIndex As Integer
Public imEndDayIndex As Integer
Public imAirCancel As Integer
Public imMatchPattern As Integer   'Indicates if all air weeks have same pattern
Public imDollarMatch As Integer   'Indicates if all air weeks have same dollar match
Public imAllWeeks As Integer       'True=Running all weeks; False=Not running all weeks
Public smOrigTotalDollars As String
Public smOrigPctTrade As String
Public smOrigStatusSaved As String
Public smOrigSchdStatusSaved As String
Public smOrigStartDate As String   'Original start date
Public smOrigEndDate As String     'Original end date
Public smOrigPattern As String     'Original pattern- every week, 2nd, 3rd or 4th week
Public smOrigNoSpots As String     'Original # spots per week
Public smOrigDaily(0 To 7) As String       'Original # spots per day or day flag if spots per week defined
Public imCalYear As Integer    'Month of displayed calendar
Public imCalMonth As Integer   'Year of displayed calendar
Public lmCalStartDate As Long  'Start date of displayed calendar
Public lmCalEndDate As Long    'End date of displayed calendar
Public imCalType As Integer
Public smLnLLD As String   'Last log date or Now (tgSpf.sSMove is checked)- set in mLnDate
Public lmAllowedChgWkStartDate As Long   'Monday of smLnLLD unless smLnLLD day is sunday then monday of next week
Public imPastMsgAsked As Integer   'Changed spots in Past message Asked
Public imWkDays(0 To 6) As Integer 'Days from flight
Public imDailySpots(0 To 6) As Integer
Public lmWkSpots As Long        'Number of spots per week
Public imWkCff As Integer
Public imPBLenClicked As Integer

'Game
Public hmGhf As Integer
Public hmGsf As Integer
Public hmCgf As Integer        'CGF Handle
Public tmCgf As CGF
Public imCgfRecLen As Integer     'CGF record length
Public tmCgfSrchKey As LONGKEY0
'Extra Fields
Public hmEff As Integer        'CFF Handle
Public tmEff As EFF
Public imEffRecLen As Integer     'CFF record length
Public tmEffSrchKey As LONGKEY0
'Media Sold
Public hmMsf As Integer        'CGF Handle
Public tmMsf As MSF
Public imMsfRecLen As Integer     'CGF record length
Public tmMsfSrchKey2 As MSFKEY2
'Media Game Numbers
Public hmMgf As Integer        'CGF Handle
Public tmMgf As MGF
Public imMgfRecLen As Integer     'CGF record length
Public tmMgfSrchKey1 As MGFKEY1
'Comment record-Header/Line
Public tmCxf As CXF            'CXF record image
Public tmCxfSrchKey As LONGKEY0  'CXF key record image
Public hmCxf As Integer        'CXF Handle
Public imCxfRecLen As Integer      'CXF record length
Public lmCxfRecPos(0 To 5) As Long
Public imInLnComment As Integer    'True=Adding line comment- in pbcHdSTab and pbcHdTab send focus to
                                'pbcLnSTab or pbcLnTab
'Comment record- Boilerplate
Public tmCmf As CMF            'CMF record image
Public tmCmfSrchKey As LONGKEY0  'CMF key record image
Public hmCmf As Integer        'CMF Handle
Public imCmfRecLen As Integer      'CMF record length
Public lmCmfRecPos As Long
Public smContrNoForSave        'Contract number used to set edcStepDropDown at end of save
'Delete
Public hmDsf As Integer        'DSF Handle
Public hmVLF As Integer

Public hmCef As Integer        'CEF Handle
Public imCefRecLen As Integer     'CEF record length
Public tmCef As CEF            'CEF record image

'Dimc imAllowedDays(1 To 7) As Integer    'Composite of allowed days from program
Public imAllowedDays(0 To 6) As Integer    'Composite of allowed days from program
'Advertiser
Public tmAdf As ADF            'ADF record image
Public tmAdfSrchKey As INTKEY0 'ADF key record image
Public hmAdf As Integer        'ADF Handle
Public imAdfRecLen As Integer     'ADF record length
Public smAdfStamp As String
'Agency
Public tmAgf As AGF            'AGF record image
Public tmAgfSrchKey As INTKEY0 'AGF key record image
Public hmAgf As Integer        'AGF Handle
Public imAgfRecLen As Integer     'AGF record length
Public smAgfStamp As String
'Spot file
Public tmSdf As SDF            'SDF record image
Public tmSdfSrchKey0 As SDFKEY0 'SDF key record image (agency)
Public tmSdfSrchKey3 As LONGKEY0 'SDF key record image (code)
Public hmSdf As Integer        'SDF Handle
Public imSdfRecLen As Integer     'SDF record length
Public hmPsf As Integer        'PSF Handle
'MG Spot file
Public tmSmf As SMF            'SMF record image
Public tmSmfSrchKey0 As SMFKEY0 'SMF key record image (agency)
Public hmSmf As Integer        'SMF Handle
Public imSmfRecLen As Integer     'SMF record length
'Public imSmfRecLen As Integer     'SMF record length
'MultiName
Public tmMnf As MNF            'MNF record image
Public tmMnfSrchKey As INTKEY0 'MNF key record image
Public hmMnf As Integer        'MNF Handle
Public imMnfRecLen As Integer     'MNF record length
Public hmVsf As Integer        'Combo
Public imVsfRecLen As Integer     'Vsf record length
Public tmVsf As VSF
Public tmVsfSrchKey As LONGKEY0
'Demo Research Data
Public tmDrf As DRF            'DRF record image
Public hmDrf As Integer        'DRF Handle
Public imDrfRecLen As Integer     'DRF record length
Public tmDrfSrchKey As DRFKEY0
'Demo Plus Data
Public tmDpf As DPF            'DPF record image
Public hmDpf As Integer        'DPF Handle
Public imDpfRecLen As Integer     'DPF record length
Public hmDef As Integer        'DEF handle
'Region
Public tmRaf As RAF            'RAF record image
Public tmRafSrchKey As LONGKEY0  'RAF key record image
Public hmRaf As Integer        'RAF Handle
Public imRafRecLen As Integer      'RAF record length
'Product
Public hmPrf As Integer 'Product file handle
Public tmPrf As PRF
Public tmPrfSrchKey0 As LONGKEY0
Public imPrfRecLen As Integer
'Short Title
Public hmSif As Integer 'Short Title file handle
Public tmSif As SIF
Public imSifRecLen As Integer
'Personnel
Public tmPnf As PNF            'PNF record image
Public tmPnfSrchKey As INTKEY0 'PNF key record image
Public hmPnf As Integer        'PNF Handle
Public imPnfRecLen As Integer     'PNF record length
'Library
Public hmLcf As Integer 'Library file handle
Public tmLcf As LCF
Public imLcfRecLen As Integer
Public tmLcfSrchKey As LCFKEY0     'LCF key record image
'Package vehicle file
Public hmPvf As Integer 'Package file handle
Public tmPvf As PVF
Public tmPvfSrchKey As LONGKEY0
Public imPvfRecLen As Integer
Public lmPvfCode() As Long
Public imPvfSpotAdj() As Integer
'Dim smRBPackageValues(1 To 5) As String
Public smRBPackageValues(0 To 5) As String
'Barter by Vehicle
Public hmVbf As Integer
Public tmVbfSrchKey1 As VBFKEY1    'Vef key record image
Public imVBfRecLen As Integer
Public tmVbf As VBF

'Gather Info file (Used to communicate to stand alone program that contract should be sent)
Public tmGif As GIF            'GIF record image
Public hmGif As Integer        'GIF Handle
Public imGifRecLen As Integer     'GIF record length

'Record Locks
Public hmRlf As Integer 'Record Lock file handle

'Recievable
Public hmRvf As Integer 'Recievable file handle
Public hmPhf As Integer
Public hmAsf As Integer
Public hmSsf As Integer
Public imExtraRcfAdded As Integer
Public tmICtrls(0 To 1) As FIELDAREA
'Calendar
Public tmCDCtrls(0 To 7) As FIELDAREA
'Rate Bar
Public tmRBCtrls(0 To 9) As FIELDAREA
Public imRBVehType() As Integer    '0=Non package; 1=package
Public imAllClicked As Integer
Public imRBBox As Integer
Public imRBRow As Integer
Public imRBInLnGridOrSum As Integer    '0=Line Grid; 1= Summary
Public smEqualValueCPP As String
Public smEqualValueCPM As String
Public smEqualValueSpots As String
Public smEqualValuePrice As String
Public imInEqualValueChgMode As Integer
Public tmLnPBDPGen() As PBDPGEN
Public imPBSpots As Integer
Public imPBFactor As Integer
Public imPBMnfDemo As Integer
Public imRBOnWks As Integer
Public imRBOffWks As Integer
Public imINRBPBCPaint As Integer
Public imINPBCPaint As Integer
'Help
Public tmChfHelp() As HLF
Public tmClfHelp() As HLF
'Fixed Price and item prices
Public smFBITotal As String    'Fixed Price amount
Public smIBCTotal As String    'Item contracted amount
Public imComputeReq As Integer 'True=Compute Required; False=Values clean
Public imSvComputeReq As Integer
Public imUpdateAllowed As Integer
Public imStartMode As Integer
Public imChgMode As Integer
Public imDGChgMode As Integer
Public imBSMode As Integer
Public imPBChgMode As Integer
Public imPBDateIndex As Integer
Public imTerminate As Integer  'True = terminating task, False= OK
Public imInSave As Integer 'Flag indicating if call to mTestField is from mSave
Public smPhoneImage As String  'Blank phone image- obtained from mkcPhone.text before input
Public smPot As String    'Potential name, saved to determine if changed
Public smSeg As String    'Segment name, saved to determine if changed
Public smPropNo As String
Public lmOrigCntrNo As Long    'Save original number incase it is changed- then change all Previous Revisions
Public smEstNo As String
Public smResvAdvt As String
Public smEDIClientCode As String
Public smEDIProductCode As String
'Public smInvProjCode As String
'Public smInvSMSNo As String
Public smDtNeeded As String
Public smInitResvNo As String
Public smInitDtNeed As String
Public smAdvertiser As String  'Advertiser name, saved to determine if changed
Public smProduct As String     'Product name, saved to determine if changed
Public smShtTitle As String     'Short Title, saved to determine if changed
Public smInitShtTitle As String     'Short Title, saved to determine if changed
Public smAgency As String      'Agency name, saved to determine if changed
Public imInCommentBranch As Integer    'True=In comment branch, exit if come in again
Public imComboBoxIndex As Integer
Public imSPersonFirstTime As Integer   'First time at field-set default if required
Public imProdFirstTime As Integer  'First time at field-set default if required
Public imAgyFirstTime As Integer   'First time at field-set default if required
Public imComp0FirstTime As Integer  'First time at field-set default if required
Public imComp1FirstTime As Integer  'First time at field-set default if required
Public imExcl0FirstTime As Integer  'First time at field-set default if required
Public imExcl1FirstTime As Integer  'First time at field-set default if required
Public imBusFirstTime As Integer
Public imGuarFirstTime As Integer
Public imBuyerFirstTime As Integer 'First time at field-set default if required
Public imPhoneFirstTime As Integer 'First time at field-set default if required
Public imDemoFirstTime As Integer  'First time at field-set default if required
Public smBaseLen As String         'Base length as definied in the rate card
Public imInQtPrice As Integer  'True=Don't call lacPrice as in lacPrice; False=call lacPrice
Public imSettingValue As Integer
Public imSettingDay As Integer     'True- allow days to be set which are illegal
Public imSettingCD As Integer     'True- don't display
Public smDefSave(0 To 20) As String '1=Product(name);2=Agency(code); 3=Salesperson(code); 4-5=Competitive(code); 6=Buyer(name); 7=Phone #;
                                 '8-16=Demo; 17-18=Exclusions; 19=Bus Cat; 20=Guarantee
Public imDefSource(0 To 20) As Integer '-1=undefined; 0 from advertiser; 1= from agency; 2=Advt user defined; 3=Agy user defined; 4=user defined (original source unknow)
                                    'Source set in LostFocus instead of mSetShow since mSetShow is called without the user entering info.
Public smRateCardPrice As String
Public smHighestPrice As String
Public smAvgPrice As String
Public smLowestPrice As String
Public imLbcArrowSetting As Integer
Public imDoubleClickName As Integer    'Name from a list was selected by double clicking
Public imPBCDoubleClick As Integer    'Name from a list was selected by double clicking
Public imDirProcess As Integer
Public imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Public imTotColor As Integer   'Color of line total box (0=Red, 1=Yellow, 2=Green)
Public imCreditColor As Integer 'Color of credit box (0=Red, 1=Yellow, 2=Green)
Public imStatsColor As Integer 'Color of statistic box (0=Red, 1=Yellow, 2=Green)
Public imLbcMouseDown As Integer  'True=List box mouse down
Public imFirstActivate As Integer
Public imDeactivated As Integer
Public imBypassFocus As Integer    'Used by number pad for mouse up event
Public imVpfIndex As Integer   'Vehicle option index
Public imCntrVefCode As Integer
Public imVefDormantAdd As Integer  'True=Dormant vehicle added
Public bmCheckSlfDormant As Boolean    'True = Check for Dormant; False = Already checked, not required
Public imPopReqd As Integer         'Flag indicating if lbcStep(2) was populated
Public imSelectDelay As Integer    'True=lbcStep(2) change mode
Public imSelectUpdate As Integer    'True=lbcStep(2) change mode
Public imHoldShift As Integer  'Shift value for cmcHold
Public lmNowDate As Long   'Todays date
Public lmMonNowDate As Long
Public lmFirstSchMonDate As Long
Public imAskBookQ As Integer   'Ask Book Question
Public imDetailSortIndex(0 To 2) As Integer  'Retain sort history
Public imSumSortIndex(0 To 2) As Integer  'Retain sort history
Public imDelaySrce As Integer  'Delay source: 0=lbcAllVehicle; 1=edcPBDate or edcRBWks(2)
Public imDiffMode As Integer   'True=Looking at Difference; False=Not looking at difference
Public smEnabledLnPrice As String  'Price set in mLnEnableBox, used in mLnSetShow
Public smLnAvgPrice As String    'Avg price set in mLnEnabledBox
Public imSalesperson As Integer    'True=User is a True Salesperson
Public imTestNum As Integer    'Salesperson area, test number
Public imSaveStay As Integer
Public smShareCaption As String
Public imUndoSelection As Integer
Public imReturnToAlerts As Integer
Public imReturnToDashboard As Integer 'TTP 10937 - Contract Dashboard: new screen to show contract header information
Public imInTseLine As Integer
Public imGameInvChg As Integer
Public imMultiMediaVefCode As Integer
Public smMultiMediaTypeItem As String
Public imNTRChg As Integer
Public imCPMChg As Integer
Public imButtonIndex As Integer
'Lines trashed
Public imTrashLnNo() As Integer
Public bmAnyVehicleOnInsertion As Boolean

Public lmTabColor(0 To 10) As Long
Public imTabMap(0 To 11) As Integer '0 is ignored
Public Const TABMULTIMEDIA = 1
Public Const TABNTR = 2
Public Const TABAIRTIME = 3
Public Const TABPODCASTCPM = 4
Public Const TABMERCH = 5
Public Const TABPROMO = 6
Public Const TABINSTALL = 7
Public Const TABAVAIL = 8
Public Const TABPROTECTION = 9
Public Const TABRESEARCH = 10

'5/18/12: Only include dormant salesperson if referenced in contract
Public tmSvCntSalesperson() As SORTCODE

'''Dan M 4/27/2011 proposal xml
'7673
Public smStationXMLWOPath As String
Public smAgencyXMLPath As String
'8032
Public smStationXMLBasePath As String

Public imSPNoCol As Integer  'Number of Columns

Public bgCntrFirstCallToVpfFind  As Boolean

Public dnf_rst As ADODB.Recordset

Public Const OVWDINDEX = 1
Public Const OVSPOTTOTALINDEX = 2
Public Const OVDAYSINDEX = 3
Public Const OVSTARTTIMEINDEX = 10
Public Const OVENDTIMEINDEX = 11
'1/2/18: Moved to BB row
'Const OVCOPYINDEX = 12
Public Const OVHIGHINDEX = 12  '13
'Const OVPREEMPTINDEX = 10
Public Const OVRANKINDEX = 13  '14
Public Const OVBBOPENINDEX = 14    '15
Public Const OVBBCLOSEINDEX = 15   '16
Public Const OVCOPYINDEX = 16
Public Const ACT1LINEUPINDEX = 17
Public Const ACT1STOREDTIMESINDEX = 18
Public Const ACT1STOREDSPOTSINDEX = 19
Public Const ACT1STOREDCLEARPCTINDEX = 20
Public Const ACT1DAYPARTFILTERINDEX = 21


Public Const OV2DAYSINDEX = 1
Public Const OV2STARTTIMEINDEX = 8
Public Const OV2ENDTIMEINDEX = 9
Public Const OV2SPLITINDEX = 10
Public Const OV2ALLOCATIONINDEX = 11
Public Const OV2ACQUISITIONINDEX = 12
Public Const OV2SOLOINDEX = 13
Public Const OV21STINDEX = 14

Public Sub cbcDifferenceSub(LINECOMMINDEX As Integer, CNTRNOINDEX As Integer, MERCHINDEX As Integer, PROMOINDEX As Integer, INTCOMMINDEX As Integer, CANCLAUSEINDEX As Integer, CHGREASONINDEX As Integer, OTHERCOMMINDEX As Integer, STATUSINDEX As Integer, REVNOINDEX As Integer, PRICEINDEX As Integer, DATEINDEX As Integer)
    Dim ilRet As Integer
    Dim ilSbfRet As Integer
    Dim ilFound As Integer
    Dim ilClf As Integer
    Dim ilDiffClf As Integer
    Dim ilCff As Integer
    Dim ilDiffCff As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLnStartDate As Long
    Dim llLnEndDate As Long
    Dim ilFirstCff As Integer
    Dim ilSpots As Integer
    Dim ilDiffSpots As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    Dim ilCffFound As Integer
    Dim ilDiffCffFound As Integer
    Dim ilPrevCff As Integer
    Dim ilCffIndex As Integer
    Dim ilAddWk As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilClfIndex As Integer
    Dim ilLineNo As Integer
    Dim ilLoop As Integer
    Dim llPrice As Long
    Dim ilSvLnRowNo As Integer
    Dim slStr As String
    Dim ilPaintLnRowNo As Integer
    Dim llMoStartDate As Long
    Dim llMoLnStartDate As Long
    Dim ilSvFreezeState As Integer
    Dim blPcfRet As Boolean

    'Get contract and show difference
    'Read base contract
    If imChgMode Then
        Exit Sub
    End If
    imChgMode = True
    Screen.MousePointer = vbHourglass
    ilSvFreezeState = igFreezeState
    igFreezeState = 0
    Contract.pbcHeader.Cls
    Contract.pbcLine.Cls
    Contract.pbcLnTotal.Cls
    Contract.pbcLnDates.Cls
    Contract.mPaintHdTitle
    Contract.mPaintLnTitle
    Contract.mPaintLnDates
    Contract.mPaintLnTotal
    Contract.pbcDemoBar(0).Cls
    Contract.pbcDemoBar(1).Cls
    mPaintDemoBarGrid 0
    mPaintDemoBarGrid 1
    If Not Contract.mReadRec(imStep3Index) Then
        igFreezeState = ilSvFreezeState
        imChgMode = False
        Screen.MousePointer = vbDefault
        'cmcCancel_Click
        imStepBoxNo = -1
        Contract.pbcStepSTab.SetFocus
        Exit Sub
    End If
    If Not Contract.mReadClfCffRec() Then
        igFreezeState = ilSvFreezeState
        imChgMode = False
        Screen.MousePointer = vbDefault
        'cmcCancel_Click
        imStepBoxNo = -1
        Contract.pbcStepSTab.SetFocus
        Exit Sub
    End If
    ilSbfRet = mReadSbfRec(hmSbf, hmRvf, hmPhf, imStep1Index)
    ilSbfRet = mReadMultiMediaRec(hmMsf, hmMgf, hmCxf, imStep1Index)
    mMoveHdRecToCtrl DATEINDEX, tmChf, tmEff, tmAdf, tmAgf, tmVsf, tmRcf, imTerminate
    Contract.mSetRevNo       'required as mMoveRecToCtrl clears lbcRevNo
    imDiffMode = False
    'mMovelnRecToCtrl
    If Contract.cbcDifference.ListIndex <= 0 Then
        Contract.pbcLnDates.Visible = True
    Else
        Contract.pbcLnDates.Visible = False
        imDiffMode = True
        ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, lmDiffCntrCode(Contract.cbcDifference.ListIndex), False, tmDiffChf, tmDiffClf(), tmDiffCff(), tmDiffCgf())
        If Not ilRet Then
            igFreezeState = ilSvFreezeState
            imChgMode = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Compute Total spots and cost
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            tgClfCntr(ilClf).lCost = 0
            tgClfCntr(ilClf).lTSpots = 0
            ilCff = tgClfCntr(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate    'Week Start date
                    'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llEndDate    'Week Start date
                    llStartDate = tgCffCntr(ilCff).lStartDate
                    llEndDate = tgCffCntr(ilCff).lEndDate
                    llMoStartDate = llStartDate
                    Do While gWeekDayLong(llMoStartDate) <> 0
                        llMoStartDate = llMoStartDate - 1
                    Loop
                    For llDate = llMoStartDate To llEndDate Step 7
                        If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                            ilSpots = 0
                            For ilDay = 0 To 6 Step 1
                                If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                                    ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                End If
                            Next ilDay
                        Else
                            ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                        End If
                        Select Case tgCffCntr(ilCff).CffRec.sPriceType
                            Case "T"
                                llPrice = tgCffCntr(ilCff).CffRec.lActPrice
                            Case Else
                                llPrice = 0
                        End Select
                        If tgClfCntr(ilClf).ClfRec.sType <> "E" Then
                            tgClfCntr(ilClf).lCost = tgClfCntr(ilClf).lCost + llPrice * ilSpots
                        Else
                            tgClfCntr(ilClf).lCost = tgClfCntr(ilClf).lCost + llPrice
                        End If
                        tgClfCntr(ilClf).lTSpots = tgClfCntr(ilClf).lTSpots + ilSpots
                    Next llDate
                End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
        Next ilClf
        For ilDiffClf = LBound(tmDiffClf) To UBound(tmDiffClf) - 1 Step 1
            tmDiffClf(ilDiffClf).lCost = 0
            tmDiffClf(ilDiffClf).lTSpots = 0
            ilDiffCff = tmDiffClf(ilDiffClf).iFirstCff
            Do While ilDiffCff <> -1
                If (tmDiffCff(ilDiffCff).iStatus = 0) Or (tmDiffCff(ilDiffCff).iStatus = 1) Then
                    'gUnpackDateLong tmDiffCff(ilDiffCff).CffRec.iStartDate(0), tmDiffCff(ilDiffCff).CffRec.iStartDate(1), llStartDate    'Week Start date
                    'gUnpackDateLong tmDiffCff(ilDiffCff).CffRec.iEndDate(0), tmDiffCff(ilDiffCff).CffRec.iEndDate(1), llEndDate    'Week Start date
                    llStartDate = tmDiffCff(ilDiffCff).lStartDate
                    llEndDate = tmDiffCff(ilDiffCff).lEndDate
                    llMoStartDate = llStartDate
                    Do While gWeekDayLong(llMoStartDate) <> 0
                        llMoStartDate = llMoStartDate - 1
                    Loop
                    For llDate = llMoStartDate To llEndDate Step 7
                        If tmDiffCff(ilDiffCff).CffRec.sDyWk = "D" Then
                            ilSpots = 0
                            For ilDay = 0 To 6 Step 1
                                If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                                    ilSpots = ilSpots + tmDiffCff(ilDiffCff).CffRec.iDay(ilDay)
                                End If
                            Next ilDay
                        Else
                            ilSpots = tmDiffCff(ilDiffCff).CffRec.iSpotsWk + tmDiffCff(ilDiffCff).CffRec.iXSpotsWk
                        End If
                        Select Case tmDiffCff(ilDiffCff).CffRec.sPriceType
                            Case "T"
                                llPrice = tmDiffCff(ilDiffCff).CffRec.lActPrice
                            Case Else
                                llPrice = 0
                        End Select
                        If tmDiffClf(ilDiffClf).ClfRec.sType <> "E" Then
                            tmDiffClf(ilDiffClf).lCost = tmDiffClf(ilDiffClf).lCost + llPrice * ilSpots
                        Else
                            tmDiffClf(ilDiffClf).lCost = tmDiffClf(ilDiffClf).lCost + llPrice
                        End If
                        tmDiffClf(ilDiffClf).lTSpots = tmDiffClf(ilDiffClf).lTSpots + ilSpots
                    Next llDate
                End If
                ilDiffCff = tmDiffCff(ilDiffCff).iNextCff
            Loop
        Next ilDiffClf
        'Match up lines and determine differences
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            'Find matching line
            ilFirstCff = -1
            For ilDiffClf = LBound(tmDiffClf) To UBound(tmDiffClf) - 1 Step 1
                If tgClfCntr(ilClf).ClfRec.iLine = tmDiffClf(ilDiffClf).ClfRec.iLine Then
                    gUnpackDateLong tgClfCntr(ilClf).ClfRec.iStartDate(0), tgClfCntr(ilClf).ClfRec.iStartDate(1), llStartDate
                    gUnpackDateLong tgClfCntr(ilClf).ClfRec.iEndDate(0), tgClfCntr(ilClf).ClfRec.iEndDate(1), llEndDate
                    gUnpackDateLong tmDiffClf(ilDiffClf).ClfRec.iStartDate(0), tmDiffClf(ilDiffClf).ClfRec.iStartDate(1), llLnStartDate
                    gUnpackDateLong tmDiffClf(ilDiffClf).ClfRec.iEndDate(0), tmDiffClf(ilDiffClf).ClfRec.iEndDate(1), llLnEndDate
                    If llEndDate >= llStartDate Then
                        If llLnEndDate >= llLnStartDate Then
                            If llLnStartDate < llStartDate Then
                                llStartDate = llLnStartDate
                            End If
                            If llLnEndDate > llEndDate Then
                                llEndDate = llLnEndDate
                            End If
                        End If
                    Else
                        llStartDate = llLnStartDate
                        llEndDate = llLnEndDate
                    End If
                    If llEndDate >= llStartDate Then
                        'Build airing week
                        llMoStartDate = llStartDate
                        Do While gWeekDayLong(llMoStartDate) <> 0
                            llMoStartDate = llMoStartDate - 1
                        Loop
                        For llDate = llMoStartDate To llEndDate Step 7
                            ilCffFound = -1
                            ilCff = tgClfCntr(ilClf).iFirstCff
                            ilSpots = 0
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    llLnStartDate = tgCffCntr(ilCff).lStartDate
                                    llLnEndDate = tgCffCntr(ilCff).lEndDate
                                    llMoLnStartDate = llLnStartDate
                                    Do While gWeekDayLong(llMoLnStartDate) <> 0
                                        llMoLnStartDate = llMoLnStartDate - 1
                                    Loop
                                    If (llDate >= llMoLnStartDate) And (llDate <= llLnEndDate) Then
                                        ilCffFound = ilCff
                                        If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                        End If
                                        Exit Do
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                            ilDiffCffFound = -1
                            ilDiffCff = tmDiffClf(ilDiffClf).iFirstCff
                            ilDiffSpots = 0
                            Do While ilDiffCff <> -1
                                If (tmDiffCff(ilDiffCff).iStatus = 0) Or (tmDiffCff(ilDiffCff).iStatus = 1) Then
                                    llLnStartDate = tmDiffCff(ilDiffCff).lStartDate
                                    llLnEndDate = tmDiffCff(ilDiffCff).lEndDate
                                    llMoLnStartDate = llLnStartDate
                                    Do While gWeekDayLong(llMoLnStartDate) <> 0
                                        llMoLnStartDate = llMoLnStartDate - 1
                                    Loop
                                    If (llDate >= llMoLnStartDate) And (llDate <= llLnEndDate) Then
                                        ilDiffCffFound = ilDiffCff
                                        If tmDiffCff(ilDiffCff).CffRec.sDyWk = "D" Then
                                            For ilDay = 0 To 6 Step 1
                                                ilDiffSpots = ilDiffSpots + tmDiffCff(ilDiffCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilDiffSpots = tmDiffCff(ilDiffCff).CffRec.iSpotsWk + tmDiffCff(ilDiffCff).CffRec.iXSpotsWk
                                        End If
                                        Exit Do
                                    End If
                                End If
                                ilDiffCff = tmDiffCff(ilDiffCff).iNextCff
                            Loop
                            ilAddWk = 0
                            If (ilCffFound >= 0) And (ilDiffCffFound >= 0) Then
                                If ilSpots <> ilDiffSpots Then
                                    ilAddWk = 1
                                Else
                                    If tgCffCntr(ilCffFound).CffRec.sPriceType <> tmDiffCff(ilDiffCffFound).CffRec.sPriceType Then
                                        ilAddWk = 1
                                    End If
                                    If tgCffCntr(ilCffFound).CffRec.lActPrice <> tmDiffCff(ilDiffCffFound).CffRec.lActPrice Then
                                        ilAddWk = 1
                                    End If
                                End If
                            ElseIf (ilCffFound >= 0) Then
                                ilAddWk = 2
                            ElseIf (ilDiffCffFound >= 0) Then
                                ilAddWk = 3
                            End If
                            If ilAddWk > 0 Then
                                ilCffIndex = UBound(tgCffCntr)
                                tgCffCntr(ilCffIndex).iNextCff = -1
                                If ilFirstCff = -1 Then
                                    ilFirstCff = ilCffIndex
                                Else
                                    tgCffCntr(ilPrevCff).iNextCff = ilCffIndex
                                End If
                                ilPrevCff = ilCffIndex
                                ReDim Preserve tgCffCntr(0 To UBound(tgCffCntr) + 1) As CFFLIST
                                tgCffCntr(UBound(tgCffCntr)).iStatus = -1 'Not Used
                                tgCffCntr(UBound(tgCffCntr)).lRecPos = 0
                                tgCffCntr(UBound(tgCffCntr)).iNextCff = -1
                                tgCffCntr(ilCffIndex).iStatus = 0   'New to not used
                                tgCffCntr(ilCffIndex).CffRec.lChfCode = tgChfCntr.lCode
                                tgCffCntr(ilCffIndex).CffRec.iClfLine = tgClfCntr(ilClf).ClfRec.iLine
                                tgCffCntr(ilCffIndex).CffRec.iCntRevNo = tgClfCntr(ilClf).ClfRec.iCntRevNo
                                tgCffCntr(ilCffIndex).CffRec.iPropVer = tgClfCntr(ilClf).ClfRec.iPropVer
                                If llDate = llStartDate Then
                                    slStartDate = Format$(llDate, "m/d/yy")
                                Else
                                    slStartDate = Format$(llDate, "m/d/yy")
                                    slStartDate = gObtainPrevMonday(slStartDate)
                                End If
                                slEndDate = Format$(llDate, "m/d/yy")
                                slEndDate = gObtainNextSunday(slEndDate)
                                If gDateValue(slEndDate) > llEndDate Then
                                    slEndDate = Format$(llEndDate, "m/d/yy")
                                End If
                                gPackDate slStartDate, tgCffCntr(ilCffIndex).CffRec.iStartDate(0), tgCffCntr(ilCffIndex).CffRec.iStartDate(1)
                                gPackDate slEndDate, tgCffCntr(ilCffIndex).CffRec.iEndDate(0), tgCffCntr(ilCffIndex).CffRec.iEndDate(1)
                                tgCffCntr(ilCffIndex).lStartDate = gDateValue(slStartDate)
                                tgCffCntr(ilCffIndex).lEndDate = gDateValue(slEndDate)
                                If ilAddWk = 1 Then
                                    If tgCffCntr(ilCffFound).CffRec.sDyWk = tmDiffCff(ilDiffCffFound).CffRec.sDyWk Then
                                        If tgCffCntr(ilCffFound).CffRec.sDyWk = "D" Then
                                            tgCffCntr(ilCffIndex).CffRec.sDyWk = "D"
                                            For ilDay = 0 To 6 Step 1
                                                tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tgCffCntr(ilCffFound).CffRec.iDay(ilDay) - tmDiffCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                                            tgCffCntr(ilCffIndex).CffRec.iSpotsWk = tgCffCntr(ilCffFound).CffRec.iSpotsWk - tmDiffCff(ilDiffCffFound).CffRec.iSpotsWk
                                            For ilDay = 0 To 6 Step 1
                                                tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tgCffCntr(ilCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        End If
                                    Else
                                        If tgCffCntr(ilCffFound).CffRec.sDyWk = "D" Then
                                            'Convert to weekly as only showing difference in spot count
                                            tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                                            tgCffCntr(ilCffIndex).CffRec.iSpotsWk = ilSpots - ilDiffSpots
                                            For ilDay = 0 To 6 Step 1
                                                If tgCffCntr(ilCffFound).CffRec.iDay(ilDay) > 0 Then
                                                    tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = 1
                                                Else
                                                    tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = 0
                                                End If
                                            Next ilDay
                                        Else
                                            tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                                            tgCffCntr(ilCffIndex).CffRec.iSpotsWk = ilSpots - ilDiffSpots
                                            For ilDay = 0 To 6 Step 1
                                                tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tgCffCntr(ilCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        End If
                                    End If
                                ElseIf ilAddWk = 2 Then
                                    If tgCffCntr(ilCffFound).CffRec.sDyWk = "D" Then
                                        tgCffCntr(ilCffIndex).CffRec.sDyWk = "D"
                                        For ilDay = 0 To 6 Step 1
                                            tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tgCffCntr(ilCffFound).CffRec.iDay(ilDay)
                                        Next ilDay
                                    Else
                                        tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                                        tgCffCntr(ilCffIndex).CffRec.iSpotsWk = tgCffCntr(ilCffFound).CffRec.iSpotsWk
                                        For ilDay = 0 To 6 Step 1
                                            tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tgCffCntr(ilCffFound).CffRec.iDay(ilDay)
                                        Next ilDay
                                    End If
                                ElseIf ilAddWk = 3 Then
                                    If tmDiffCff(ilDiffCffFound).CffRec.sDyWk = "D" Then
                                        tgCffCntr(ilCffIndex).CffRec.sDyWk = "D"
                                        For ilDay = 0 To 6 Step 1
                                            tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = -tmDiffCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                        Next ilDay
                                    Else
                                        tgCffCntr(ilCffIndex).CffRec.sDyWk = "W"
                                        tgCffCntr(ilCffIndex).CffRec.iSpotsWk = -tmDiffCff(ilDiffCffFound).CffRec.iSpotsWk
                                        For ilDay = 0 To 6 Step 1
                                            tgCffCntr(ilCffIndex).CffRec.iDay(ilDay) = tmDiffCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                        Next ilDay
                                    End If
                                End If
                                tgCffCntr(ilCffIndex).CffRec.sDelete = "N"
                                If ilAddWk = 1 Then
                                    If tgCffCntr(ilCffFound).CffRec.sPriceType <> tmDiffCff(ilDiffCffFound).CffRec.sPriceType Then
                                        If tgCffCntr(ilCffFound).CffRec.sPriceType = "T" Then
                                            tgCffCntr(ilCffIndex).CffRec.sPriceType = tgCffCntr(ilCffFound).CffRec.sPriceType
                                            tgCffCntr(ilCffIndex).CffRec.lActPrice = tgCffCntr(ilCffFound).CffRec.lActPrice
                                        ElseIf tmDiffCff(ilDiffCffFound).CffRec.sPriceType = "T" Then
                                            tgCffCntr(ilCffIndex).CffRec.sPriceType = tmDiffCff(ilDiffCffFound).CffRec.sPriceType
                                            tgCffCntr(ilCffIndex).CffRec.lActPrice = -tmDiffCff(ilDiffCffFound).CffRec.lActPrice
                                        End If
                                    Else
                                        tgCffCntr(ilCffIndex).CffRec.sPriceType = tgCffCntr(ilCffFound).CffRec.sPriceType
                                        tgCffCntr(ilCffIndex).CffRec.lActPrice = tgCffCntr(ilCffFound).CffRec.lActPrice - tmDiffCff(ilDiffCffFound).CffRec.lActPrice
                                    End If
                                ElseIf ilAddWk = 2 Then
                                ElseIf ilAddWk = 3 Then
                                End If
                            End If
                        Next llDate
                        tgClfCntr(ilClf).iFirstCff = ilFirstCff
                        tgClfCntr(ilClf).lCost = tgClfCntr(ilClf).lCost - tmDiffClf(ilDiffClf).lCost
                        tgClfCntr(ilClf).lTSpots = tgClfCntr(ilClf).lTSpots - tmDiffClf(ilDiffClf).lTSpots
                        Exit For
                    End If
                End If
            Next ilDiffClf
        Next ilClf
        For ilDiffClf = LBound(tmDiffClf) To UBound(tmDiffClf) - 1 Step 1
            'Find matching line
            ilFirstCff = -1
            ilFound = False
            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                If tgClfCntr(ilClf).ClfRec.iLine = tmDiffClf(ilDiffClf).ClfRec.iLine Then
                    ilFound = True
                    Exit For
                End If
            Next ilClf
            If Not ilFound Then
                'Add line
                ilClfIndex = UBound(tgClfCntr)
                tgClfCntr(ilClfIndex) = tmDiffClf(ilDiffClf)
                tgClfCntr(ilClfIndex).lCost = -tgClfCntr(ilClfIndex).lCost
                tgClfCntr(ilClfIndex).lTSpots = -tgClfCntr(ilClfIndex).lTSpots
                ilLineNo = 0
                For ilLoop = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                    If tgClfCntr(ilLoop).ClfRec.iLine > ilLineNo Then
                        ilLineNo = tgClfCntr(ilLoop).ClfRec.iLine
                    End If
                Next ilLoop
                'tgClfCntr(ilClfIndex).ClfRec.iLine = ilLineNo + 1
                tgClfCntr(ilClfIndex).iFirstCff = -1
                tgClfCntr(ilClfIndex).iFirstCgf = -1
                ReDim Preserve tgClfCntr(0 To ilClfIndex + 1) As CLFLIST
                ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                tgClfCntr(UBound(tgClfCntr)).iFirstCff = -1
                tgClfCntr(UBound(tgClfCntr)).iFirstCgf = -1
                tgClfCntr(UBound(tgClfCntr)).iStatus = -1
                ilDiffCff = tmDiffClf(ilDiffClf).iFirstCff
                Do While ilDiffCff <> -1
                    tgCffCntr(UBound(tgCffCntr)) = tmDiffCff(ilDiffCff)
                    If tgClfCntr(ilClfIndex).iFirstCff = -1 Then
                        tgClfCntr(ilClfIndex).iFirstCff = UBound(tgCffCntr)
                    Else
                        tgCffCntr(UBound(tgCffCntr) - 1).iNextCff = UBound(tgCffCntr)
                    End If
                    tgCffCntr(UBound(tgCffCntr)).iNextCff = -1
                    tgCffCntr(UBound(tgCffCntr)).lRecPos = 0
                    tgCffCntr(UBound(tgCffCntr)).iStatus = 0
                    If tgCffCntr(UBound(tgCffCntr)).CffRec.sDyWk = "D" Then
                        tgCffCntr(UBound(tgCffCntr)).CffRec.sDyWk = "D"
                        For ilDay = 0 To 6 Step 1
                            tgCffCntr(UBound(tgCffCntr)).CffRec.iDay(ilDay) = -tgCffCntr(UBound(tgCffCntr)).CffRec.iDay(ilDay)
                        Next ilDay
                    Else
                        tgCffCntr(UBound(tgCffCntr)).CffRec.sDyWk = "W"
                        tgCffCntr(UBound(tgCffCntr)).CffRec.iSpotsWk = -tgCffCntr(UBound(tgCffCntr)).CffRec.iSpotsWk
                    End If
                    tgCffCntr(UBound(tgCffCntr)).CffRec.lActPrice = -tgCffCntr(UBound(tgCffCntr)).CffRec.lActPrice
                    ReDim Preserve tgCffCntr(0 To UBound(tgCffCntr) + 1) As CFFLIST
                    tgCffCntr(UBound(tgCffCntr)).iStatus = -1 'Not Used
                    tgCffCntr(UBound(tgCffCntr)).iNextCff = -1
                    tgCffCntr(UBound(tgCffCntr)).lRecPos = 0
                    ilDiffCff = tmDiffCff(ilDiffCff).iNextCff
                Loop
            End If
        Next ilDiffClf
    End If
    '8/6/19: set all lines to show
    If imDiffMode Then
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            tgClfCntr(ilClf).bPkExpanded = True
            tgClfCntr(ilClf).bShowLine = True
        Next ilClf
        mInitPackArrays False
        bgSortByLineNo = Not bgSortByLineNo
        Contract.lacPrice_Click 2
    End If
    'Set cost
    If imLnTotalType <> 1 Then
        imLnTotalType = 1
        Contract.lacPrice(0).Caption = "Total Price"    ' & sgLF & sgCR & "Price"
        Contract.pbcSumType(1).Cls
        Contract.pbcSumType_Paint (1)
    End If
    'Remove unchanged lines
    mMoveLnRecToCtrl imTerminate, tmRegionCode()
    Contract.mSetLineScroll True, True
    Contract.mSetHdDate False, False  'Init Hot values
    Contract.mInitDmShow
    mInitCEShow tmCECtrls()
    Contract.mInitSPShow
    Contract.mInitSBShow
    Contract.mInitRSShow
    Contract.mInitPTShow
    Contract.mInitNoShow
    Contract.mInitLnShow
    For ilLoop = imLBCTCtrls To UBound(tmHdCtrls) Step 1
        If (ilLoop <> CNTRNOINDEX) Or (ilRet = 0) Then
            If ilLoop = MERCHINDEX Then
                slStr = smComment(0)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = PROMOINDEX Then
                slStr = smComment(1)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = INTCOMMINDEX Then
                slStr = smComment(2)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = CANCLAUSEINDEX Then
                slStr = smComment(3)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = CHGREASONINDEX Then
                slStr = smComment(4)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = OTHERCOMMINDEX Then
                slStr = smComment(5)
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            ElseIf ilLoop = STATUSINDEX Then
                If (ilRet = 0) And (Contract.lbcStatus.ListIndex < 0) And (sgCntrScreen <> "Orders") And ((imStep1Index = 1) Or (imStep1Index = 4)) Then
                    imChgMode = True
                    Contract.lbcStatus.ListIndex = 0
                    imChgMode = True
                ElseIf (ilRet = 0) And (Contract.lbcStatus.ListIndex < 0) And (sgCntrScreen = "Orders") And ((imStep1Index = 1) Or (imStep1Index = 2)) Then
                    imChgMode = True
                    Contract.lbcStatus.ListIndex = 1
                    imChgMode = True
                End If
               Contract.mSetShow ilLoop  'Set show strings
            ElseIf ilLoop = REVNOINDEX Then
                slStr = Contract.edcRevNo.Text
                If ((sgCntrScreen = "Orders") And (imStep1Index < 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Then
                    slStr = Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & slStr
                End If
                If (slStr = "0") Or (slStr = "0-0") Then
                    slStr = ""
                End If
                gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
            Else
                Contract.mSetShow ilLoop  'Set show strings
            End If
        Else    'Don't call mSetShow as it sets the contract number of blank
            slStr = Contract.edcCntrNo.Text
            gSetShow Contract.pbcHeader, slStr, tmHdCtrls(CNTRNOINDEX)
        End If
    Next ilLoop
    If imDiffMode Then
        ilSvLnRowNo = imLnRowNo
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            smLnSave(18, ilClf + 1) = gLongToStrDec(tgClfCntr(ilClf).lCost, 2)
            'imLnSave(11, ilClf + 1) = tgClfCntr(ilClf).iTSpots
            smLnSave(23, ilClf + 1) = Trim$(str$(tgClfCntr(ilClf).lTSpots))
            'If tgClfCntr(ilClf).lCost > 0 Then
                slStr = smLnSave(18, ilClf + 1)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            'Else
            '    slStr = ""
            'End If
            'gSetShow Contract.pbcLine, slStr, tmLnCtrls(PRICEINDEX)
            mSetPriceShow slStr, PRICEINDEX
            smLnShow(PRICEINDEX, ilClf + 1) = tmLnCtrls(PRICEINDEX).sShow
            'imLnRowNo = ilClf + 1
            'mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX
        Next ilClf
        imLnRowNo = ilSvLnRowNo
    End If
    '12/11/19
    Contract.lacAudTotals_Click
    ilPaintLnRowNo = imPaintLnRowNo
    imPaintLnRowNo = -1
    Contract.mPkgTotal
    Contract.mCntrTotal
    Contract.pbcHeader_Paint
    Contract.pbcLine_Paint
    Contract.mSetCreditInfo
    mSetSchStatus smOrigStatus
    imPaintLnRowNo = ilPaintLnRowNo
    Contract.cbcDifference.SetFocus
    imChgMode = False
    igFreezeState = ilSvFreezeState
    Screen.MousePointer = vbDefault
End Sub

Public Function mSaveRecFunction(STATUSINDEX As Integer, CNTRNOINDEX As Integer, NOPROPINDEX As Integer, GROSSINDEX As Integer, PRICEINDEX As Integer, PTAGYCOMMINDEX As Integer) As Integer 'L.Bianchi 06/24/2021 TTP 10185
    Dim ilLoop As Integer   'For loop control
    Dim ilIndex As Integer
    Dim ilTest As Integer
    'Dim slName As String    'Name
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilErrRet As Integer
    'Dim ilFilter As Integer 'cbcFilter index
    'Dim ilType As Integer   '1=Contract; 0=Orders
    Dim slNameCode As String  'Name and Code number
    Dim slCode As String    'Code number
    Dim slMsg As String
    'Dim slStamp As String   'Date/Time stamp for file
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilNewChf As Integer
    Dim ilNewClf As Integer
    Dim ilUpdateLine As Integer
    Dim slLineSchStatus As String
    'Dim ilFound As Integer
    'Dim ilRecLen As Integer
    Dim ilChgHdStatus As Integer    'True=Change header status to "A"
    Dim ilAnyLineChgd As Integer
    Dim ilEntryDate0 As Integer
    Dim ilEntryDate1 As Integer
    Dim ilEntryTime0 As Integer
    Dim ilEntryTime1 As Integer
    Dim ilNewHeader As Integer
    Dim ilCntRevNo As Integer
    Dim llChfRecPos As Long
    Dim llTChfRecPos As Long
    Dim ilCxfState As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim llSNTRDate As Long
    Dim llENTRDate As Long
    Dim llSInstallDate As Long
    Dim llEInstallDate As Long
    Dim slChfType As String         'same as tgChfCntr.sType except if scheduling Remnant then its S
    Dim tlClf As CLF
    Dim tlCff As CFF
    Dim ilCgf As Integer
    Dim ilMsf As Integer
    Dim ilMgf As Integer
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim llSbfDate As Long
    Dim llMPEndDate As Long
    Dim slTotalDollars As String
    Dim slLimit As String
    Dim slDollarChgd As String
    '8132
    Dim blTreatAsRep As Boolean
    '8033
    Dim blIsMissingStationRights As Boolean
    Dim ilAgf As Integer
    Dim ilCommPct As Integer
    Dim llCntrChfCode As Long
    Dim llTCntrChfCode As Long
    Dim llRow As Long

    blIsMissingStationRights = False
    Screen.MousePointer = vbHourglass  'Wait
    Contract.mSetShow imHdBoxNo
    If Not imUpdateAllowed Then
        ilRet = MsgBox("Can't Save Changes to Contract", vbOKOnly + vbExclamation, "Message")
        Screen.MousePointer = vbDefault    'Default
        mSaveRecFunction = False
        Exit Function
    End If
    imInSave = True
    If Contract.udcCBill.Visible = True Then
        If Not Contract.udcCBill.Verify() Then
            Screen.MousePointer = vbDefault    'Default
            imInSave = False
            mSaveRecFunction = False
            Exit Function
        End If
        Contract.udcCBill.Action 5
        Contract.mPkgTotal
        Contract.mCntrTotal
    End If
    If Contract.udcCGameInv.Visible = True Then
        Contract.udcCGameInv.Action 5
    End If
    For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
        If tgIBSbf(llSbf).SbfRec.iIhfCode > 0 Then
            If (tgIBSbf(llSbf).iStatus >= 0) And (tgIBSbf(llSbf).SbfRec.iIhfCode > 0) Then
                gUnpackDateLong tgIBSbf(llSbf).SbfRec.iDate(0), tgIBSbf(llSbf).SbfRec.iDate(1), llSbfDate
                If llSbfDate <= 0 Then
                    ilRet = MsgBox("Multi-Media NTR Item missing Billing Date", vbOKOnly + vbExclamation, "Message")
                    Screen.MousePointer = vbDefault    'Default
                    imInSave = False
                    mSaveRecFunction = False
                    Exit Function
                End If
            End If
        '10883 0 = new 1 = old but retain.  2 means to delete so don't test
        'ElseIf tgIBSbf(llSbf).iStatus >= 0 Then
        ElseIf tgIBSbf(llSbf).iStatus = 0 Or tgIBSbf(llSbf).iStatus = 1 Then
                gUnpackDateLong tgIBSbf(llSbf).SbfRec.iDate(0), tgIBSbf(llSbf).SbfRec.iDate(1), llSbfDate
                If llSbfDate <= 0 Then
                    ilRet = MsgBox("Multi-Media NTR Item missing Billing Date", vbOKOnly + vbExclamation, "Message")
                    Screen.MousePointer = vbDefault    'Default
                    imInSave = False
                    mSaveRecFunction = False
                    Exit Function
                End If
        End If
    Next llSbf


    If imSaveStay Then
        imMultiMediaVefCode = Contract.udcCGameInv.MultiMediaVefCode
        smMultiMediaTypeItem = Contract.udcCGameInv.MultiMediaTypeItem
    Else
        imMultiMediaVefCode = -1
        smMultiMediaTypeItem = ""
    End If
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If Not Contract.udcCGameInv.Verify() Then
                'ilRet = MsgBox("Multi-Media Oversold, reduce number units purchased", vbOkOnly + vbExclamation, "Message")
                Screen.MousePointer = vbDefault    'Default
                imInSave = False
                mSaveRecFunction = False
                Exit Function
            End If
        End If
    End If
    If Contract.udcCDigital.Visible = True Then
        Contract.udcCDigital.Action 5
    End If
    ilTest = False
    If (tgUrf(0).sWorkToHold <> "Y") And (tgUrf(0).sWorkToOrder <> "Y") And (tgUrf(0).sCompToHold <> "Y") And (tgUrf(0).sCompToOrder <> "Y") And (tgUrf(0).sIncompToHold <> "Y") And (tgUrf(0).sIncompToOrder <> "Y") Then
        ilTest = True
    ElseIf (tgUrf(0).iSlfCode > 0) And (tgSpf.sGUsePropSys = "N") Then
        ilTest = True
    End If
    If ilTest Then
        'Check Status and if Working or Unapproved- ask if it should be changed to Complete
        ilIndex = Contract.lbcStatus.ListIndex
        If ilIndex < 0 Then
            If (sgCntrScreen <> "Orders") And (imStep1Index = 1) Then
                slStr = "Working Proposal"
            ElseIf (sgCntrScreen <> "Orders") And (imStep1Index = 4) Then
                slStr = "Rev Working"
            ElseIf (sgCntrScreen = "Orders") And (imStep1Index = 1) Then
                slStr = "Approved Order"
            ElseIf (sgCntrScreen = "Orders") And (imStep1Index = 2) Then
                slStr = "Approved Order"
            Else
                slStr = ""  'Don't set as Undo will not come on
            End If
        Else
            slStr = Contract.lbcStatus.List(ilIndex)
        End If
        If (slStr = "Working Proposal") Or (slStr = "Rev Working") Or (slStr = "Unapproved Proposal") Or (slStr = "Rev Unapproved") Then
            For ilLoop = 0 To Contract.lbcStatus.ListCount - 1 Step 1
                slStr = Contract.lbcStatus.List(ilLoop)
                If (slStr = "Completed Proposal") Or (slStr = "Rev Completed") Then
                    ilRet = MsgBox("Is this Proposal Complete", vbYesNo + vbQuestion, "Status")
                    If ilRet = vbYes Then
                        Contract.lbcStatus.ListIndex = ilLoop
                        tmHdCtrls(STATUSINDEX).iChg = True
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    If imStep1Index <= 1 Then
        slStr = Trim$(Contract.edcCntrNo.Text)
        'If (slStr = "") Or (Val(slStr) <= 0) Then
        If (slStr = "") Or (Val(slStr) <= 0) Or (smPropNo = "") Or (Val(smPropNo) <= 0) Then
            Screen.MousePointer = vbDefault    'Default
            CCntrNo.Show vbModal
            If igTerminateReturn = 0 Then
                imInSave = False
                mSaveRecFunction = False
                Exit Function
            End If
            smPropNo = Contract.edcCntrNo.Text
            tmNoCtrls(NOPROPINDEX).iChg = True
            slStr = smPropNo
            gSetShow Contract.pbcHeader, slStr, tmHdCtrls(CNTRNOINDEX)
            slStr = smPropNo
            gSetShow Contract.pbcNumber, slStr, tmNoCtrls(NOPROPINDEX)
            Screen.MousePointer = vbHourglass
        End If
    End If
    'If Contract.tscLine.SelectedItem.Index = 10 Then    'reset to air time to avoid problems with pop daypart
    If Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH) Or Contract.tscLine.SelectedItem.Index = imTabMap(TABPODCASTCPM) Then    'reset to air time to avoid problems with pop daypart
        'Set Contract.tscLine.SelectedItem = Contract.tscLine.Tabs.Item(8)
        Contract.mActiveLnTab
    End If
    
      'L.Bianchi 06/22/2021
    'Test if Agency commision is set when Trade percentage is alloted TTP 10185
    If (Contract.lbcAgency.ListIndex > 1) Then
        If gCompNumberStr(smPctTrade, "0") > 0 And imAgyComm < 0 Then
            'Screen.MousePointer = vbDefault    'Default
                tmPTCtrls(PTAGYCOMMINDEX).iReq = True
        End If
    End If
    
    tmHdCtrls(GROSSINDEX).iReq = True  'Only mandatory when update pressed
    tmHdCtrls(CNTRNOINDEX).iReq = True  'Only mandatory when update pressed
    tmNoCtrls(NOPROPINDEX).iReq = True 'Set at save
    If Contract.mTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        tmHdCtrls(GROSSINDEX).iReq = False  'Only mandatory when update pressed
        tmHdCtrls(CNTRNOINDEX).iReq = False  'Only mandatory when update pressed
        tmNoCtrls(NOPROPINDEX).iReq = False 'Set at save
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        Exit Function
    End If
    tmHdCtrls(GROSSINDEX).iReq = False  'Only mandatory when update pressed
    tmHdCtrls(CNTRNOINDEX).iReq = False  'Only mandatory when update pressed
    tmNoCtrls(NOPROPINDEX).iReq = False 'Set at save
    tmPTCtrls(PTAGYCOMMINDEX).iReq = False 'L.Bianchi 06/24/2021 TTP 10185
    
    'Test if Merchandising/Promotion and no amounts defined
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If (Len(Trim$(smComment(0))) <> 0) And (UBound(tgMBSbf) - 1 < LBound(tgMBSbf)) Then
                Screen.MousePointer = vbDefault    'Default
                MsgBox "Merchandising Amounts not defined, press Merchandising tab and Specify Amounts", vbOKOnly + vbExclamation, "Warning"
                imInSave = False
                mSaveRecFunction = False
                imHdBoxNo = -1
                imLnRowNo = -1
                'cmcBills.SetFocus
                Exit Function
            End If
        End If
    End If
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If (Len(Trim$(smComment(1))) <> 0) And (UBound(tgPBSbf) - 1 < LBound(tgPBSbf)) Then
                Screen.MousePointer = vbDefault    'Default
                MsgBox "Promotion Amounts not defined, press Promotion tab and Specify Amounts", vbOKOnly + vbExclamation, "Warning"
                imInSave = False
                mSaveRecFunction = False
                imHdBoxNo = -1
                imLnRowNo = -1
                'cmcBills.SetFocus
                Exit Function
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass  'Wait
    imBypassEPriceComp = False
    gGetSyncDateTime smSyncDate, smSyncTime
    If Not Contract.mMoveLnCtrlToRec(True) Then
        imBypassEPriceComp = True
        imInSave = False
        Screen.MousePointer = vbDefault    'Default
        mSaveRecFunction = False
        Exit Function
    End If

    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If (Len(Trim$(smComment(0))) <> 0) Then
                If Not Contract.udcCBill.CheckVehicles(0) Then
                    Screen.MousePointer = vbDefault    'Default
                    MsgBox "Merchandising vehicle removed from lines, press Merchandising tab and Create Amounts", vbOKOnly + vbExclamation, "Warning"
                    imInSave = False
                    mSaveRecFunction = False
                    imHdBoxNo = -1
                    imLnRowNo = -1
                    'cmcBills.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            If (Len(Trim$(smComment(1))) <> 0) Then
                If Not Contract.udcCBill.CheckVehicles(1) Then
                    Screen.MousePointer = vbDefault    'Default
                    MsgBox "Promotion vehicle removed from lines, press Promotion tab and Create Amounts", vbOKOnly + vbExclamation, "Warning"
                    imInSave = False
                    mSaveRecFunction = False
                    imHdBoxNo = -1
                    imLnRowNo = -1
                    'cmcBills.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    'Check that Merchandising and Promotion dates within air dates
    If Contract.lbcStatus.ListIndex >= 0 Then
        slStr = Contract.lbcStatus.List(Contract.lbcStatus.ListIndex)
        If (InStr(1, slStr, "Order", 1) > 0) Or (InStr(1, slStr, "Hold", 1) > 0) Then
            ilRet = mCheckMPAgainstCntr(llMPEndDate)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault    'Default
                If ilRet = 1 Then
                    MsgBox "Merchandising dates outside of range of Air Time, press Merchandising tab to correct", vbOKOnly + vbExclamation, "Warning"
                Else
                    MsgBox "Promotion dates outside of range of Air Time, press Promotion tab to correct", vbOKOnly + vbExclamation, "Warning"
                End If
                imInSave = False
                mSaveRecFunction = False
                imHdBoxNo = -1
                imLnRowNo = -1
                'cmcBills.SetFocus
                Exit Function
            End If
        End If
    End If
    imBypassEPriceComp = True
    imInSave = False
    '8033 set each to false
    mStationXmlAllowedOrderStart
    '8132
    blTreatAsRep = False
    'Set line types
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        '8033 only returns false if station xml vehicle, but blocked 'site' wide.
        If mStationXmlAllowedOrder(tgClfCntr(ilClf).ClfRec.iVefCode) Then
        Else
            blIsMissingStationRights = True
        End If
        '8132 get 'vehicle/stations' to go out. Above set igStationXmlChoice
        If bgStationAreVehicles And Not blTreatAsRep And Not imCntrRep And igStationXmlChoice <> STATIONXMLNONE Then
            If gGetStationFromVehicle(tgClfCntr(ilClf).ClfRec.iVefCode) > 0 Then
                blTreatAsRep = True
            End If
        End If
        If (sgCntrScreen <> "Orders") Or (tgSpf.sGUsePropSys = "Y") Then    'And (imStep1Index <> 4) Then
            imInResearch = True 'Only set Books
            mInitBookDemoForLn ilClf + 1, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode
            imInResearch = False
        End If
        mSetPropPriceForNewWks ilClf + 1, tmRcf
        tgClfCntr(ilClf).ClfRec.sType = "S"
        tgClfCntr(ilClf).ClfRec.iPkLineNo = 0
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilLoop).iLineNo Then
                tgClfCntr(ilClf).ClfRec.sType = tgPackLn(ilLoop).sType
                If (tgPackLn(ilLoop).sType <> "O") And (tgPackLn(ilLoop).sType <> "A") And (tgPackLn(ilLoop).sType <> "E") Then
                    For ilTest = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If (tgPackLn(ilTest).sType = "A") Or (tgPackLn(ilTest).sType = "O") Or (tgPackLn(ilTest).sType = "E") Then
                            ilIndex = tgPackLn(ilTest).iIndex
                            Do While ilIndex >= 0
                                If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                                    tgClfCntr(ilClf).ClfRec.iPkLineNo = tgPackLn(ilTest).iLineNo
                                    Exit For
                                End If
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                        End If
                    Next ilTest
                Else
                    mLnTotals ilClf + 1, imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), "N", imNoWkDays()
                    mGetLnPriceInfo ilClf + 1, slStr, imLnTotalType, PRICEINDEX     'Formats Price within Show
                End If
                Exit For
            End If
        Next ilLoop
    Next ilClf
    igLnSPChg = False
    If Contract.mLnTestFields() = NO Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        Exit Function
    End If
    igTabMapIndex = 1
    If Not gCPMTestFields() Then
        igTabMapIndex = 0
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        Exit Function
    End If
    igTabMapIndex = 0
    If imCntrAir Then
        imCntrRep = False
    End If
    Contract.mMoveCtrlToRec True
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        Exit Function
    End If
    
    'Check that the Acquisition cost is always less or equal to the spot rate
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        If (Not tgClfCntr(ilClf).iCancel) And ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) Then
            If tgClfCntr(ilClf).sOwnership = "B" Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If tgClfCntr(ilClf).ClfRec.lAcquisitionCost > tgCffCntr(ilCff).CffRec.lActPrice Then
                        MsgBox "The Acquisition Cost may not exceed Spot Rate as in Line  " & tgClfCntr(ilClf).ClfRec.iLine, vbOKOnly + vbExclamation, "Warning"
                        Screen.MousePointer = vbDefault    'Default
                        imInSave = False
                        mSaveRecFunction = False
                        Exit Function
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
            End If
        End If
    Next ilClf
    
    '9/11/14: This allow Installment and Sports vehicles if using Revenue By Events
    If Not Contract.mRevneueByEventOk() Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        imHdBoxNo = -1
        imLnRowNo = -1
        Exit Function
    End If
    
    '10/7/16: Determine if using installment
    'Rules:
    '   1.  When a contract is commissionable (with agency) and only NTR exists, commission for the NTR items must be all Yes, or all No.
    '   2.  When a contract has a combination of air time and NTR, all NTR commission must be consistent with the  airtime commission.
    '   3.  If a Direct advertiser and NTR only, All NTR items must be set to No Agency Commission.
    '   4.  If a direct advertiser and combination of air time and NTR, all NTR items must be set to No Agency Commission.
    ''12/5/14: Add test to disallow any billing cycle except Standard with Installment billing
    'If Not mCheckInstallment() Then
    ilRet = mCheckInstallment()
    If ilRet = 1 Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        imHdBoxNo = -1
        imLnRowNo = -1
        Exit Function
    ElseIf ilRet = 2 Then
        If tgChfCntr.iAgfCode <= 0 Then
            For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
                If (tgIBSbf(llSbf).iStatus >= 0) Then
                    If tgIBSbf(llSbf).SbfRec.sAgyComm = "Y" Then
                        Screen.MousePointer = vbDefault    'Default
                        MsgBox "NTR agency commission not permitted for Direct Advertiser", vbOKOnly + vbExclamation, "Warning"
                        imInSave = False
                        mSaveRecFunction = False
                        imHdBoxNo = -1
                        imLnRowNo = -1
                        Exit Function
                    End If
                End If
            Next llSbf
        Else
            ilCommPct = -1
            ilAgf = gBinarySearchAgf(tgChfCntr.iAgfCode)
            If ilAgf <> -1 Then
                ilCommPct = tgCommAgf(ilAgf).iCommPct
            End If
            If UBound(tgClfCntr) > LBound(tgClfCntr) Then
                For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
                    If (tgIBSbf(llSbf).iStatus >= 0) Then
                        If ((tgIBSbf(llSbf).SbfRec.sAgyComm = "Y") And (ilCommPct = 0)) Or ((tgIBSbf(llSbf).SbfRec.sAgyComm = "N") And (ilCommPct <> 0)) Then
                            Screen.MousePointer = vbDefault    'Default
                            If ilCommPct = 0 Then
                                MsgBox "NTR agency commission must be set to No as the Agency Commission is zero", vbOKOnly + vbExclamation, "Warning"
                            Else
                                MsgBox "NTR agency commission must be set to Yes as the Agency Commission is non-zero", vbOKOnly + vbExclamation, "Warning"
                            End If
                            imInSave = False
                            mSaveRecFunction = False
                            imHdBoxNo = -1
                            imLnRowNo = -1
                            Exit Function
                        End If
                    End If
                Next llSbf
            Else
                ilRet = -1
                For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
                    If (tgIBSbf(llSbf).iStatus >= 0) Then
                        If ilRet >= 0 Then
                            If tgIBSbf(llSbf).SbfRec.sAgyComm <> tgIBSbf(ilRet).SbfRec.sAgyComm Then
                                Screen.MousePointer = vbDefault    'Default
                                MsgBox "NTR agency commission must be set the same for all transactions", vbOKOnly + vbExclamation, "Warning"
                                imInSave = False
                                mSaveRecFunction = False
                                imHdBoxNo = -1
                                imLnRowNo = -1
                                Exit Function
                            End If
                            ilRet = llSbf
                        Else
                            ilRet = llSbf
                        End If
                    End If
                Next llSbf
            End If
        End If
    End If
    
    If imCntrRep Then
        If tgChfCntr.sBillCycle <> "S" Then
            Screen.MousePointer = vbDefault    'Default
            MsgBox "Standard Broadcast Billing Cycle required for Rep Contract", vbOKOnly + vbExclamation, "Warning"
            imInSave = False
            mSaveRecFunction = False
            imHdBoxNo = -1
            imLnRowNo = -1
            Exit Function
        End If
        If ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMSTARTDATE) <> "")) Then
            Screen.MousePointer = vbDefault    'Default
            MsgBox "Rep and Ad Server Buy Mixture Not Allowed", vbOKOnly + vbExclamation, "Warning"
            imInSave = False
            mSaveRecFunction = False
            imHdBoxNo = -1
            imLnRowNo = -1
            Exit Function
        End If
    End If
    If Val(smPctTrade) > 0 Then
        If ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMSTARTDATE) <> "")) Then
            Screen.MousePointer = vbDefault    'Default
            MsgBox "Trade Contract Not Allowed with Ad Server Buys", vbOKOnly + vbExclamation, "Warning"
            imInSave = False
            mSaveRecFunction = False
            imHdBoxNo = -1
            imLnRowNo = -1
            Exit Function
        End If
    End If
    'Check price discrepancy- Leave as N for Remnant, per Inquire, PSA & Promos
    Contract.mCntrTotal
    tgChfCntr.sNRProcessed = "Y"
    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
        If imCntrRep Then
            If (tgChfCntr.lCntrNo >= tgSpf.lCLowestNo) And (tgChfCntr.lCntrNo <= tgSpf.lCHighestNo) Then
                If (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "O") Then
                    tgChfCntr.sNRProcessed = "N"
                End If
            End If
        End If
        tgChfCntr.sRepDBID = ""
    ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        If imStep1Index <= 1 Then
            tgChfCntr.sRepDBID = ""
        End If
    Else
        tgChfCntr.sRepDBID = ""
    End If
    If ((sgCntrScreen = "Orders") And (imStep1Index = 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 5)) Then
        'Disallow any line changes 7/2/98- Jim
    Else
        If imTotColor <> 2 Then
            Screen.MousePointer = vbDefault    'Default
            If tgSaf(0).sCreditLimitMsg = "C" Then
                ilRet = MsgBox("Gross Amount and Total Line Amount must be in Balance", vbOKOnly + vbExclamation, "Message")
                Screen.MousePointer = vbDefault    'Default
                imHdBoxNo = GROSSINDEX
                imInSave = False
                mSaveRecFunction = False
                Exit Function
            Else
                ilRet = MsgBox("Gross Amount and Total Line Amount not in Balance, Continue with Save", vbYesNo + vbQuestion, "Warning")
                If ilRet = vbNo Then
                    imHdBoxNo = GROSSINDEX
                    imInSave = False
                    mSaveRecFunction = False
                    Exit Function
                End If
            End If
            Screen.MousePointer = vbHourglass
        End If
        If Not mTestMPBalance(hmAgf, smTotalDollars) Then
            Screen.MousePointer = vbDefault    'Default
            imInSave = False
            mSaveRecFunction = False
            Exit Function
        End If
        If Not mTestFBBalance(gAddStr(smTotalDollars, smIBCTotal)) Then
            Screen.MousePointer = vbDefault    'Default
            imInSave = False
            mSaveRecFunction = False
            Exit Function
        End If
        If Not mWarningMsg(imStep1Index, smComment()) Then
            Screen.MousePointer = vbDefault    'Default
            imInSave = False
            mSaveRecFunction = False
            Exit Function
        End If
        
        '6/8/14
        gUnformatStr smTotalDollars, 0, slTotalDollars
        slTotalDollars = gAddStr(slTotalDollars, smIBCTotal)
        slDollarChgd = gSubStr(slTotalDollars, smOrigTotalDollars)
        If Val(smPctTrade) > 0 Then
            '8032
            If mIsStationXmlAllowedForSite(STATIONXMLMARKETRON) And Val(smPctTrade) < 100 Then
                gMsgBox "Because you are using Station Xml-Marketron Electronic Orders, you cannot mix cash/trade in one order.  Save will be cancelled."
                Screen.MousePointer = vbDefault    'Default
                imInSave = False
                mSaveRecFunction = False
                Exit Function
            End If
            slDollarChgd = gDivStr(gMulStr(slDollarChgd, gSubStr("100.00", smPctTrade)), "100.00")
        End If
        If (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "N") And (tgChfCntr.sStatus <> "O") Then
            If (smOrigStatusSaved = "G") Or (smOrigStatusSaved = "H") Or (smOrigStatusSaved = "N") Or (smOrigStatusSaved = "O") Then
                If (smOrigSchdStatusSaved <> "F") And (smOrigSchdStatusSaved <> "I") And (smOrigSchdStatusSaved <> "M") Then
                    'slDollarChgd = "-" & smOrigTotalDollars
                    If Val(smPctTrade) > 0 Then
                        slDollarChgd = "-" & gDivStr(gMulStr(smOrigTotalDollars, gSubStr("100.00", smOrigPctTrade)), "100.00")
                    Else
                        slDollarChgd = "-" & smOrigTotalDollars
                    End If
                End If
            End If
        End If
        If tgSaf(0).sCreditLimitMsg = "C" Then
            If InStr(1, Contract.lacCreditInfo.Caption, "Credit Remaining:", vbBinaryCompare) = 1 Then
                slStr = Mid(Contract.lacCreditInfo.Caption, 19)
                gUnformatStr slStr, 0, slLimit
                gUnformatStr smTotalDollars, 0, slTotalDollars
                slTotalDollars = gAddStr(slTotalDollars, smIBCTotal)
                'ilRet = gCompNumberStr(slLimit, gSubStr(slTotalDollars, smOrigTotalDollars))
                'If ilRet < 0 Then
                slStr = gSubStr(slLimit, slDollarChgd)
                If Val(slStr) < 0 Then
                    slStr = Replace(slStr, "-", "$")
                    If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "O") Then
                        ilRet = MsgBox("Add/Change not allowed as Credit Limit exceeded by " & slStr, vbOKOnly + vbExclamation, "Message")
                        Screen.MousePointer = vbDefault    'Default
                        imInSave = False
                        mSaveRecFunction = False
                        Exit Function
                    Else
                        ilRet = MsgBox("Credit Limit exceeded by " & slStr & ", Continue with Save", vbYesNo + vbQuestion, "Warning")
                        If ilRet = vbNo Then
                            Screen.MousePointer = vbDefault    'Default
                            imInSave = False
                            mSaveRecFunction = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
    End If
    slStr = Trim$(str$(tgChfCntr.lCntrNo))
    If Not Contract.mOKID(slStr) Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        Exit Function
    End If
    'Set Contract Gross Impressions and GRP
    '10258 act one needs these values to be saved to clf
    mSetPriDemoResearch hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating, , , True
    'mSetPriDemoResearch hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating
    Screen.MousePointer = vbHourglass  'Wait
    'NTR Billing
    If smOrigNTRDefined = "Y" Then
        tgChfCntr.sNTRDefined = smOrigNTRDefined
    Else
        tgChfCntr.sNTRDefined = "N"
    End If

    mGetNTRDateRange llSNTRDate, llENTRDate
    mGetInstallmentDateRange llSInstallDate, llEInstallDate
    If llSInstallDate <= 0 Then
        tgChfCntr.sInstallDefined = "N"
    Else
        tgChfCntr.sInstallDefined = "Y"
    End If
    'Set CBS Flag
    tgChfCntr.sCBSOrder = "N"
    If (tgChfCntr.sNTRDefined <> "Y") And (llSInstallDate <= 0) And ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) = "") Or (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) = "")) Then
        ilRet = True    'Assume cancel before start
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            gUnpackDateLong tgClfCntr(ilClf).ClfRec.iStartDate(0), tgClfCntr(ilClf).ClfRec.iStartDate(1), llSDate
            gUnpackDateLong tgClfCntr(ilClf).ClfRec.iEndDate(0), tgClfCntr(ilClf).ClfRec.iEndDate(1), llEDate
            If llSDate <= llEDate Then
                ilRet = False
                Exit For
            End If
        Next ilClf
        If ilRet Then
            tgChfCntr.sCBSOrder = "C"
        End If
    End If
    '12/24/19: Removed transaction logging because of freezing with multi-users
    'ilRet = btrBeginTrans(hmCHF, 1000)
    'Remove lines and any history lines
    If UBound(tgDelClf) > LBound(tgDelClf) Then
        ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, tgChfCntr.lCode, True, tmDiffChf, tmDiffClf(), tmDiffCff(), tmDiffCgf())    'tmChf, tgClfCntr(), tgCffCntr(), tgCffCntr())
        If ilRet Then
            For ilLoop = LBound(tgDelClf) To UBound(tgDelClf) - 1 Step 1
                'Find matching lines and flights and delete
                For ilClf = LBound(tmDiffClf) To UBound(tmDiffClf) - 1 Step 1
                    If tmDiffClf(ilClf).ClfRec.iLine = tgDelClf(ilLoop).ClfRec.iLine Then
                        ilCff = tmDiffClf(ilClf).iFirstCff
                        Do While ilCff <> -1
                            Do
                                'ilRet = btrGetDirect(hmCff, tlCff, imCffRecLen, tmDiffCff(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                tmCffSrchKey1.lCode = tmDiffCff(ilCff).CffRec.lCode
                                ilRet = btrGetEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cff)", Contract
                                On Error GoTo 0
                                ilRet = btrDelete(hmCff)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cff)", Contract
                            On Error GoTo 0
                            ilCff = tmDiffCff(ilCff).iNextCff
                        Loop
                        ilCgf = tmDiffClf(ilClf).iFirstCgf
                        Do While ilCgf <> -1
                            Do
                                tmCgfSrchKey.lCode = tmDiffCgf(ilCgf).CgfRec.lCode
                                ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cgf)", Contract
                                On Error GoTo 0
                                ilRet = btrDelete(hmCgf)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cgf)", Contract
                            On Error GoTo 0
                            ilCgf = tmDiffCgf(ilCgf).iNextCgf
                        Loop
                        Do
                            tmCxfSrchKey.lCode = tmDiffClf(ilClf).ClfRec.lCxfCode
                            If tmCxfSrchKey.lCode <> 0 Then
                                tmCxf.sComment = ""
                                imCxfRecLen = Len(tmCxf) '5027
                                ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cxf)", Contract
                                On Error GoTo 0
                                ilRet = btrDelete(hmCxf)
                                If ilRet <> BTRV_ERR_CONFLICT Then
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cxf)", Contract
                                    On Error GoTo 0
                                End If
                            Else
                                ilRet = BTRV_ERR_NONE
                                Exit Do
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Chf)", Contract
                        On Error GoTo 0
                        Do
                            'ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tmDiffClf(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            tmClfSrchKey2.lCode = tmDiffClf(ilClf).ClfRec.lCode
                            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (GetEqaul- Clf)", Contract
                            On Error GoTo 0
                            ilRet = btrDelete(hmClf)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Clf)", Contract
                        On Error GoTo 0
                    End If
                Next ilClf
            Next ilLoop
        End If
    End If
    ilRet = mCreateVef()
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, "Create Vef", Contract
    On Error GoTo 0
    ilRet = gCreateCPMVef()
    If Not mSetHdVeh(hmVsf, imStep1Index) Then
        Screen.MousePointer = vbDefault    'Default
        'imTerminate = True
        imInSave = False
        mSaveRecFunction = False
        'ilRet = btrAbortTrans(hmCHF)
        Exit Function
    End If
    If (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "N") And (tgChfCntr.sStatus <> "O") Then
        If (smOrigStatusSaved <> "G") And (smOrigStatusSaved <> "H") And (smOrigStatusSaved <> "N") And (smOrigStatusSaved <> "O") Then
            slDollarChgd = ""
        ElseIf (smOrigSchdStatusSaved = "F") Or (smOrigSchdStatusSaved = "I") Or (smOrigSchdStatusSaved = "M") Then
            slDollarChgd = ""
        End If
    End If
    If Not mUpdateAdvtAgy(hmAdf, tmAdf, hmAgf, tmAgf, slDollarChgd, imTerminate) Then
        Screen.MousePointer = vbDefault    'Default
        'imTerminate = True
        imInSave = False
        mSaveRecFunction = False
        'ilRet = btrAbortTrans(hmCHF)
        Exit Function
    End If
    Contract.mSetHdDate False, True
    '6/5/20: Handle case where only NTR defined
    llSDate = 0
    llEDate = 0
    If (UBound(tgClfCntr) > LBound(tgClfCntr)) Or ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMSTARTDATE) <> "")) Then
        gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llSDate
        gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llEDate
    End If
    If llMPEndDate > 0 Then
        'gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llDate
        If llMPEndDate > llEDate Then
            gPackDateLong llMPEndDate, tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1)
            llEDate = llMPEndDate
        End If
    End If
    If tgChfCntr.sNTRDefined = "Y" Then
        'gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llDate
        llSNTRDate = gObtainPrevMondayLong(llSNTRDate)
        If ((llSNTRDate > 0) And (llSNTRDate < llSDate)) Or ((llSNTRDate > 0) And (llSDate = 0)) Then
            gPackDateLong llSNTRDate, tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1)
            llSDate = llSNTRDate
        End If
        'gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llDate
        If ((llENTRDate > 0) And (llENTRDate > llEDate)) Or ((llENTRDate > 0) And (llEDate = 0)) Then
            gPackDateLong llENTRDate, tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1)
            llEDate = llENTRDate
        End If
    End If
        
    If llSInstallDate > 0 Then
        'gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llDate
        llSInstallDate = gObtainPrevMondayLong(llSInstallDate)
        If ((llSInstallDate > 0) And (llSInstallDate < llSDate)) Or ((llSInstallDate > 0) And (llSDate = 0)) Then
            gPackDateLong llSInstallDate, tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1)
            llSDate = llSInstallDate
        End If
        'gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llDate
        If ((llEInstallDate > 0) And (llEInstallDate > llEDate)) Or ((llEInstallDate > 0) And (llEDate = 0)) Then
            gPackDateLong llEInstallDate, tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1)
            llEDate = llEInstallDate
        End If
    End If
    ''If dates zero, reset to previous dates
    'gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llSDate
    'gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llEDate
    If (llSDate = 0) Or (llEDate = 0) Then
        tmChfSrchKey1.lCntrNo = lmOrigCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmOrigCntrNo) Then
            If llSDate = 0 Then
                tgChfCntr.iStartDate(0) = tmChf.iStartDate(0)
                tgChfCntr.iStartDate(1) = tmChf.iStartDate(1)
            End If
            If llEDate = 0 Then
                tgChfCntr.iEndDate(0) = tmChf.iEndDate(0)
                tgChfCntr.iEndDate(1) = tmChf.iEndDate(1)
            End If
        End If
    Else
        '8/23/18: Test that contract dates don't excede 104 weeks
        If (llEDate - llSDate) / 7 > 103 Then
            If (gDateValue(smChfEndDate) - gDateValue(smChfStartDate)) / 7 < 104 Then
                Screen.MousePointer = vbDefault    'Default
                MsgBox "Contract cannot exceed 2 year span", vbOKOnly + vbExclamation + vbApplicationModal, "Too Many Weeks"
                imInSave = False
                mSaveRecFunction = False
                'ilRet = btrAbortTrans(hmCHF)
                Exit Function
            End If
        End If
    End If
    'Check price discrepancy- Leave as N for Remnant, per Inquire, PSA & Promos
    Contract.mCntrTotal
    tgChfCntr.sDiscrep = "N"
    If (tgChfCntr.sType <> "T") And (tgChfCntr.sType <> "Q") And (tgChfCntr.sType <> "S") And (tgChfCntr.sType <> "M") Then
        If imTotColor <> 2 Then
            tgChfCntr.sDiscrep = "Y"
        End If
    End If
    '3/1/18: handle formatted dollars
    'tgChfCntr.lNTRGross = gStrDecToLong(smIBCTotal, 2)
    gUnformatStr smIBCTotal, 0, slStr
    tgChfCntr.lNTRGross = gStrDecToLong(slStr, 2)
    '3/1/18: handle formatted dollars
    'tgChfCntr.lAirTimeGross = gStrDecToLong(smTotalDollars, 2)
    gUnformatStr smTotalDollars, 0, slStr
    tgChfCntr.lAirTimeGross = gStrDecToLong(slStr, 2)
    ilChgHdStatus = False   'Set to true when processing lines
    ilAnyLineChgd = False
    ilCntRevNo = tgChfCntr.iCntRevNo
    If (sgCntrScreen = "Orders") And (imStep1Index = 2) Then
        If (smOrigStatus = "O") Or (smOrigStatus = "H") Then
            ilCntRevNo = -1    'Force new revision number
        End If
        ''Increment revision # if required
        'If tgChfCntr.sPrint = "P" Then 'Printed
        '    'No resources included as part of contract/proposal check
    End If
    If ((sgCntrScreen = "Orders") And (imStep1Index = 2)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4) And ((tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N"))) Then
        'If (smOrigStatus = "O") Or (smOrigStatus = "H") Then
            Screen.MousePointer = vbDefault    'Default
            CRevNo.Show vbModal
            DoEvents
            If igTerminateReturn = 0 Then
                imInSave = False
                mSaveRecFunction = False
                'ilRet = btrAbortTrans(hmCHF)
                Exit Function
            End If
            Screen.MousePointer = vbHourglass
        'End If
    End If
    If (sgCntrScreen = "Orders") Then
        'imStep1Index: 0=New Scratch; 1=New Model; 2=Revise Order 3= Change Hold to Order
        If (imStep1Index <= 1) Or ((imStep1Index = 2) And (ilCntRevNo <> tgChfCntr.iCntRevNo)) Then
            ilNewHeader = True
            If imStep1Index <= 1 Then
                tgChfCntr.iCntRevNo = 0
                tgChfCntr.iPropVer = 0
                tgChfCntr.iExtRevNo = 0
                If mEDSSetNeg2(0, "", bmAnyVehicleOnInsertion) Then
                    tgChfCntr.iEDSSentExtRevNo = -2
                Else
                    tgChfCntr.iEDSSentExtRevNo = -1
                End If
                gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                tgChfCntr.iEDSSentUrfCode = 0
                tgChfCntr.sSource = ""
            Else
                'Get oldest header and determine next number
                tmChfSrchKey1.lCntrNo = lmOrigCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmOrigCntrNo) Then
                    If (tmChf.sStatus <> "H") And (tmChf.sStatus <> "O") Then
                        Screen.MousePointer = vbDefault    'Default
                        MsgBox "Contract # " & str$(lmOrigCntrNo) & " has been Revised to a Rev Proposal, it can't be modified", vbOKOnly + vbExclamation + vbApplicationModal, "Status Change"
                        imInSave = False
                        mSaveRecFunction = False
                        'ilRet = btrAbortTrans(hmCHF)
                        Exit Function
                    End If
                Else
                    Screen.MousePointer = vbDefault    'Default
                    MsgBox "Contract # " & str$(lmOrigCntrNo) & " can't be found to get latest revision #, Save Stopped", vbOKOnly + vbExclamation + vbApplicationModal, "# Missing"
                    imInSave = False
                    mSaveRecFunction = False
                    'ilRet = btrAbortTrans(hmCHF)
                    Exit Function
                End If
                If (tmChf.sSchStatus = "A") Or (tmChf.sSchStatus = "N") Then
                    ilNewHeader = False
                    tgChfCntr.iCntRevNo = tmChf.iCntRevNo
                    tgChfCntr.iPropVer = tmChf.iPropVer
                Else
                    tgChfCntr.iCntRevNo = tmChf.iCntRevNo + 1
                    tgChfCntr.iPropVer = tmChf.iPropVer
                End If
                If ilNewHeader Then
                    If mEDSSetNeg2(0, "", bmAnyVehicleOnInsertion) Then
                        tgChfCntr.iEDSSentExtRevNo = -2 '-2 indicates not used
                        gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                        gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = 0
                    ElseIf mEDSSetNeg1(tmChf.iEDSSentExtRevNo, 0, "", bmAnyVehicleOnInsertion) Then
                        tgChfCntr.iEDSSentExtRevNo = -1
                        gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                        gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = 0
                    Else
                        tgChfCntr.iEDSSentExtRevNo = tmChf.iEDSSentExtRevNo
                        tgChfCntr.iEDSSentDate(0) = tmChf.iEDSSentDate(0)
                        tgChfCntr.iEDSSentDate(1) = tmChf.iEDSSentDate(1)
                        tgChfCntr.iEDSSentTime(0) = tmChf.iEDSSentTime(0)
                        tgChfCntr.iEDSSentTime(1) = tmChf.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = tmChf.iEDSSentUrfCode
                    End If
                End If
                tgChfCntr.sSource = tmChf.sSource
            End If
        Else
            ilNewHeader = False
        End If
    Else
        'imStep1Index: 0=New Scratch; 1= New Model; 2=Chg proposal; 3=New Version of Existing Proposal;
        '4=Revise an Existing Order; 5=Change Hold to Order
        If (imStep1Index <= 1) Or (imStep1Index = 3) Or (imStep1Index = 4) Then
            If imStep1Index <= 1 Then
                ilNewHeader = True
                tgChfCntr.iCntRevNo = 0
                tgChfCntr.iPropVer = 1
                tgChfCntr.iExtRevNo = 0
                If mEDSSetNeg2(0, "", bmAnyVehicleOnInsertion) Then
                    tgChfCntr.iEDSSentExtRevNo = -2
                Else
                    tgChfCntr.iEDSSentExtRevNo = -1
                End If
                gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                tgChfCntr.iEDSSentUrfCode = 0
            Else
                'Get oldest header and determine next number
                tmChfSrchKey1.lCntrNo = lmOrigCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmOrigCntrNo) Then
                    'With the separation of proposals and orders this code shouldn't be required
                    If (imStep1Index = 3) And ((tmChf.sStatus = "H") Or (tmChf.sStatus = "O")) Then
                        Screen.MousePointer = vbDefault    'Default
                        MsgBox "Contract # " & str$(lmOrigCntrNo) & " has been Revised to a Order, it can't be modified", vbOKOnly + vbExclamation + vbApplicationModal, "Status Change"
                        imInSave = False
                        mSaveRecFunction = False
                        'ilRet = btrAbortTrans(hmCHF)
                        Exit Function
                    End If
                Else
                    Screen.MousePointer = vbDefault    'Default
                    MsgBox "Contract # " & str$(lmOrigCntrNo) & " can't be found to get latest Revision/Version #, Save Stopped", vbOKOnly + vbExclamation + vbApplicationModal, "# Missing"
                    imInSave = False
                    mSaveRecFunction = False
                    'ilRet = btrAbortTrans(hmCHF)
                    Exit Function
                End If
                If imStep1Index = 3 Then
                    ilNewHeader = True
                    tgChfCntr.iCntRevNo = tmChf.iCntRevNo
                    tgChfCntr.iPropVer = tmChf.iPropVer + 1
                Else
                    If (smOrigStatus = "O") Or (smOrigStatus = "H") Then
                        ilNewHeader = True
                        tgChfCntr.iCntRevNo = tmChf.iCntRevNo + 1
                        'tgChfCntr.iPropVer = 1
                        tgChfCntr.iPropVer = tmChf.iPropVer 'Retain proposal version number
                    Else
                        ilNewHeader = False
                    End If
                End If
                If ilNewHeader Then
                    If mEDSSetNeg2(0, "", bmAnyVehicleOnInsertion) Then
                        tgChfCntr.iEDSSentExtRevNo = -2 '-2 indicates not used
                        gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                        gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = 0
                    ElseIf mEDSSetNeg1(tmChf.iEDSSentExtRevNo, 0, "", bmAnyVehicleOnInsertion) Then
                        tgChfCntr.iEDSSentExtRevNo = -1
                        gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                        gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = 0
                    Else
                        tgChfCntr.iEDSSentExtRevNo = tmChf.iEDSSentExtRevNo
                        tgChfCntr.iEDSSentDate(0) = tmChf.iEDSSentDate(0)
                        tgChfCntr.iEDSSentDate(1) = tmChf.iEDSSentDate(1)
                        tgChfCntr.iEDSSentTime(0) = tmChf.iEDSSentTime(0)
                        tgChfCntr.iEDSSentTime(1) = tmChf.iEDSSentTime(1)
                        tgChfCntr.iEDSSentUrfCode = tmChf.iEDSSentUrfCode
                    End If
                End If
            End If
        Else
            ilNewHeader = False
        End If
    End If

    Screen.MousePointer = vbHourglass  'Wait
    For ilLoop = 0 To 5 Step 1
        ilRet = BTRV_ERR_NONE
        ilCxfState = 0
        Do  'Loop until record updated or added
            If Not ilNewHeader Then
                'Reposition
                Select Case ilLoop
                    Case 0
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfMerch
                    Case 1
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfProm
                    Case 2
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfInt
                    Case 3
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfCanc
                    Case 4
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfChgR
                    Case 5
                        tmCxfSrchKey.lCode = tgChfCntr.lCxfCode
                End Select
                If tmCxfSrchKey.lCode <> 0 Then
                    tmCxf.sComment = ""
                    imCxfRecLen = Len(tmCxf) '5027
                    ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, "mSaveRec (btrGetEqual:Comment)", Contract
                    On Error GoTo 0
                Else
                    tmCxf.lCode = 0
                End If
            End If
            tmCxf.sComment = ""
            Select Case ilLoop
                Case 0
                    tmCxf.sComType = "M"
                Case 1
                    tmCxf.sComType = "P"
                Case 2
                    tmCxf.sComType = "I"
                Case 3
                    tmCxf.sComType = "C"
                Case 4
                    tmCxf.sComType = "R"
                Case 5
                    tmCxf.sComType = "O"
            End Select
            If (imCommShowOn(ilLoop, 0) = 0) And (ilLoop <> 2) Then
                tmCxf.sShProp = "Y"
            Else
                tmCxf.sShProp = "N"
            End If
            If (imCommShowOn(ilLoop, 1) = 0) And (ilLoop <> 2) Then
                tmCxf.sShOrder = "Y"
            Else
                tmCxf.sShOrder = "N"
            End If
            If (imCommShowOn(ilLoop, 2) = 0) Then
                tmCxf.sShSpot = "Y"
            Else
                tmCxf.sShSpot = "N"
            End If
            If (imCommShowOn(ilLoop, 3) = 0) And (ilLoop <> 2) Then
                tmCxf.sShInv = "Y"
            Else
                tmCxf.sShInv = "N"
            End If
            If (imCommShowOn(ilLoop, 4) = 0) And (ilLoop = 5) Then
                tmCxf.sShInsertion = "Y"
            Else
                tmCxf.sShInsertion = "N"
            End If
            'tmCxf.iStrLen = Len(smComment(ilLoop))
            tmCxf.sComment = Trim$(smComment(ilLoop)) & Chr$(0) '& Chr$(0) 'sgTB
            'imCxfRecLen = 25 + Len(Trim$(tmCxf.sComment)) + 2    '25 = fixed record length; 2=Length value which is part of the variable record
            imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment)) ' + 2    '25 = fixed record length; 2=Length value which is part of the variable record
            If ilNewHeader Then  'New selected
                ''If imCxfRecLen - 2 > 27 Then  '-2 so control character at end not counted
                'If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so control character at end not counted
                If Trim$(smComment(ilLoop)) <> "" Then  '-2 so control character at end not counted
                    ilCxfState = 1
                    tmCxf.lCode = 0 'Autoincrement
                    tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmCxf.lAutoCode = tmCxf.lCode
                    ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                Else
                    tmCxf.lCode = 0
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrInsert: Comment)"
            Else 'Old record-Update
                ''If imCxfRecLen - 2 > 27 Then  '-2 so the control character at end is not counted
                'If Len(Trim$(tmCxf.sComment)) > 2 Then  '-2 so the control character at end is not counted
                If Trim$(smComment(ilLoop)) <> "" Then
                    If tmCxf.lCode = 0 Then
                        ilCxfState = 1
                        tmCxf.lCode = 0 'Autoincrement
                        tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                        tmCxf.lAutoCode = tmCxf.lCode
                        ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                    Else
                        ilCxfState = 2
                        tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                        gPackDate smSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                        gPackTime smSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                        ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                    End If
                Else
                    If tmCxf.lCode <> 0 Then
                        ilCxfState = 3
                        ilRet = btrDelete(hmCxf)
                    End If
                    tmCxf.lCode = 0
                    ilRet = BTRV_ERR_NONE
                End If
                slMsg = "mSaveRec (btrUpdate: Comment)"
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Contract
        On Error GoTo 0
        Select Case ilLoop
            Case 0
                tgChfCntr.lCxfMerch = tmCxf.lCode
            Case 1
                tgChfCntr.lCxfProm = tmCxf.lCode
            Case 2
                tgChfCntr.lCxfInt = tmCxf.lCode
            Case 3
                tgChfCntr.lCxfCanc = tmCxf.lCode
            Case 4
                tgChfCntr.lCxfChgR = tmCxf.lCode
            Case 5
                tgChfCntr.lCxfCode = tmCxf.lCode
        End Select
        If ilCxfState = 1 Then
            Do
                tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                tmCxf.lAutoCode = tmCxf.lCode
                tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate smSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                gPackTime smSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                slMsg = "mSaveRec (btrUpdate: Comment RU ID)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
        ElseIf ilCxfState = 3 Then
        End If
    Next ilLoop
    If Trim$(tgChfCntr.sProduct) <> "" Then
        gFindMatch tgChfCntr.sProduct, 0, Contract.lbcProd
        If gLastFound(Contract.lbcProd) < 0 Then
            imPrfRecLen = Len(tmPrf)
            Do  'Loop until record updated or added
                tmPrf.lCode = 0
                tmPrf.iAdfCode = tmAdf.iCode
                tmPrf.sName = tgChfCntr.sProduct
                tmPrf.iMnfComp(0) = tgChfCntr.iMnfComp(0)
                tmPrf.iMnfComp(1) = tgChfCntr.iMnfComp(1)
                tmPrf.iMnfExcl(0) = tgChfCntr.iMnfExcl(0)
                tmPrf.iMnfExcl(1) = tgChfCntr.iMnfExcl(1)
                tmPrf.iPnfBuyer = tmAdf.iPnfBuyer
                tmPrf.sCppCpm = tgChfCntr.sCppCpm
                For ilLoop = 0 To 3
                    tmPrf.iMnfDemo(ilLoop) = tgChfCntr.iMnfDemo(ilLoop)
                    tmPrf.lTarget(ilLoop) = tgChfCntr.lTarget(ilLoop)
                    tmPrf.lLastCPP(ilLoop) = 0
                    tmPrf.lLastCPM(ilLoop) = 0
                Next ilLoop
                tmPrf.sState = "A"
                tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert:Product)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
            Do
                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmPrf.lAutoCode = tmPrf.lCode
                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate smSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                gPackTime smSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                slMsg = "mSaveRec (btrUpdate:Product)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
        Else
            slNameCode = tmProdCode(gLastFound(Contract.lbcProd) - 1).sKey    'Traffic!Contract.lbcProdCode.List(gLastFound(Contract.lbcProd) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            slCode = Trim$(slCode)
            Do
                tmPrfSrchKey0.lCode = Val(slCode)
                ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                tmPrf.iMnfComp(0) = tgChfCntr.iMnfComp(0)
                tmPrf.iMnfComp(1) = tgChfCntr.iMnfComp(1)
                tmPrf.iMnfExcl(0) = tgChfCntr.iMnfExcl(0)
                tmPrf.iMnfExcl(1) = tgChfCntr.iMnfExcl(1)
                'tmPrf.iPnfBuyer = tmAdf.iPnfBuyer
                tmPrf.sCppCpm = tgChfCntr.sCppCpm
                For ilLoop = 0 To 3
                    tmPrf.iMnfDemo(ilLoop) = tgChfCntr.iMnfDemo(ilLoop)
                    tmPrf.lTarget(ilLoop) = tgChfCntr.lTarget(ilLoop)
                Next ilLoop
                tmPrf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                gPackDate smSyncDate, tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                gPackTime smSyncTime, tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                slMsg = "mSaveRec (btrUpdate:Product)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
            smProdCodeTag = ""  'Force repopulation
        End If
    End If
    If (Trim$(smShtTitle) <> "") And (Trim$(smShtTitle) <> "[None]") And (tgSpf.sUseProdSptScr = "P") Then
        gFindMatch smShtTitle, 0, Contract.lbcShtTitle
        If gLastFound(Contract.lbcShtTitle) < 0 Then
            imSifRecLen = Len(tmSif)
            Do  'Loop until record updated or added
                tmSif.lCode = 0
                tmSif.iAdfCode = tmAdf.iCode
                tmSif.sName = smShtTitle
                tmSif.sState = "A"
                tmSif.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                tmSif.lAutoCode = tmSif.lCode
                ilRet = btrInsert(hmSif, tmSif, imSifRecLen, INDEXKEY0)
                slMsg = "mSaveRec (btrInsert:Short Title)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
            tgChfCntr.lSifCode = tmSif.lCode
            sgShtTitleCodeTag = ""  'Force repopulation
            Do
                tmSif.iRemoteID = tgUrf(0).iRemoteUserID
                tmSif.lAutoCode = tmSif.lCode
                gPackDate smSyncDate, tmSif.iSyncDate(0), tmSif.iSyncDate(1)
                gPackTime smSyncTime, tmSif.iSyncTime(0), tmSif.iSyncTime(1)
                ilRet = btrUpdate(hmSif, tmSif, imSifRecLen)
                slMsg = "mSaveRec (btrUpdate:Short Title)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
        Else
            slNameCode = tgShtTitleCode(gLastFound(Contract.lbcShtTitle)).sKey ' - 1).sKey    'Contract.lbcShtTitleCode.List(gLastFound(Contract.lbcShtTitle) - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tgChfCntr.lSifCode = Val(slCode)
        End If
    Else
        tgChfCntr.lSifCode = 0
    End If
    If (Trim$(smEDIClientCode) <> "") Or (Trim$(smEDIProductCode) <> "") Then
    'If (Trim$(smInvProjCode) <> "") Or (Trim$(smInvSMSNo) <> "") Then
        If (tgChfCntr.lEffCode > 0) And (Not ilNewHeader) Then
            tmEffSrchKey.lCode = tgChfCntr.lEffCode
            ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmEff)
            End If
        End If
        tmEff.lCode = 0
        tmEff.sType = "I"
        tmEff.sString(0) = smEDIClientCode
        tmEff.sString(1) = smEDIProductCode
        'tmEff.sString(0) = smInvProjCode
        'tmEff.sString(1) = smInvSMSNo
        tmEff.sString(2) = ""
        tmEff.lLong(0) = 0
        tmEff.lLong(1) = 0
        tmEff.lLong(2) = 0
        tmEff.sUnused = ""
        ilRet = btrInsert(hmEff, tmEff, imEffRecLen, INDEXKEY0)
        If ilRet = BTRV_ERR_NONE Then
            tgChfCntr.lEffCode = tmEff.lCode
        Else
            tgChfCntr.lEffCode = 0
        End If
    Else
        If Not ilNewHeader Then
            If tgChfCntr.lEffCode > 0 Then
                tmEffSrchKey.lCode = tgChfCntr.lEffCode
                ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    ilRet = btrDelete(hmEff)
                End If
            End If
        End If
        tgChfCntr.lEffCode = 0
    End If
    slStr = Format$(gNow(), "m/d/yy")
    gPackDate slStr, ilEntryDate0, ilEntryDate1
    slStr = Format$(gNow(), "h:m:s AM/PM")
    gPackTime slStr, ilEntryTime0, ilEntryTime1
    'Determine date range and Vehicle for contract
    If (tgSpf.sSchdRemnant = "Y") And (tgChfCntr.sType = "T") Then
        slChfType = "C" 'Act as if standard
    ElseIf (tgSpf.sSchdPromo = "Y") And (tgChfCntr.sType = "M") Then
        slChfType = "C" 'Act as if standard
    ElseIf (tgSpf.sSchdPSA = "Y") And (tgChfCntr.sType = "S") Then
        slChfType = "C" 'Act as if standard
    Else
        slChfType = tgChfCntr.sType
    End If
    Do  'Loop until record updated or added
        If (ilNewHeader) Then 'New selected
            ilNewChf = True
            tgChfCntr.lCode = 0  'Autoincrement
            If (tgSpf.sInvAirOrder = "O") Or (tgSpf.sInvAirOrder = "S") Then
                tgChfCntr.sMGMiss = "M"
            Else
                tgChfCntr.sMGMiss = "G"
            End If
            If ((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Or ((sgCntrScreen = "Orders") And (imStep1Index = 2)) Then
                'Leave value
            Else
                tgChfCntr.iHdChg = 0
            End If
            'tgChfCntr.iCntRevNo = 0
            'tgChfCntr.iPropVer = 0
            tgChfCntr.sDelete = "N"
            If (sgCntrScreen = "Orders") Then
                tgChfCntr.iOHDDate(0) = ilEntryDate0
                tgChfCntr.iOHDDate(1) = ilEntryDate1
                tgChfCntr.iOHDTime(0) = ilEntryTime0
                tgChfCntr.iOHDTime(1) = ilEntryTime1
                Select Case slChfType    'tgChfCntr.sType
                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        tgChfCntr.sSchStatus = "M"
                        If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                            tgChfCntr.sStatus = "O"
                        End If
                    Case Else   'Standard or Reservation
                        If imStep1Index <= 1 Then
                            tgChfCntr.sSchStatus = "N"
                        Else    'imStep2Index = 2 (revise order) and Rev incremented
                            If tgChfCntr.sSchStatus = "M" Then
                                tgChfCntr.iHdChg = tgChfCntr.iHdChg Or HDSCHSTATUSMCHG
                            End If
                            'If (tgChfCntr.iHdChg <> 0) Then
                            If (tgChfCntr.iHdChg <> 0) Or ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "")) Then
                                tgChfCntr.sSchStatus = "A" 'Altered
                            End If
                        End If
                End Select
            Else
                Select Case slChfType  'tgChfCntr.sType
                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        tgChfCntr.sSchStatus = "M"
                        If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                            tgChfCntr.sStatus = "O"
                        End If
                End Select
                If (tgChfCntr.sStatus <> "O") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "N") Then
                    If tgChfCntr.sSchStatus = "M" Then

                    End If
                    tgChfCntr.sSchStatus = "P"  'Prevent Scheduling
                    If tgChfCntr.sStatus = "D" Then
                        tgChfCntr.iOHDDate(0) = ilEntryDate0
                        tgChfCntr.iOHDDate(1) = ilEntryDate1
                        tgChfCntr.iOHDTime(0) = ilEntryTime0
                        tgChfCntr.iOHDTime(1) = ilEntryTime1
                    Else
                        tgChfCntr.iPropDate(0) = ilEntryDate0
                        tgChfCntr.iPropDate(1) = ilEntryDate1
                        tgChfCntr.iPropTime(0) = ilEntryTime0
                        tgChfCntr.iPropTime(1) = ilEntryTime1
                        tgChfCntr.iOHDDate(0) = ilEntryDate0
                        tgChfCntr.iOHDDate(1) = ilEntryDate1
                        tgChfCntr.iOHDTime(0) = ilEntryTime0
                        tgChfCntr.iOHDTime(1) = ilEntryTime1
                    End If
                Else
                    tgChfCntr.iOHDDate(0) = ilEntryDate0
                    tgChfCntr.iOHDDate(1) = ilEntryDate1
                    tgChfCntr.iOHDTime(0) = ilEntryTime0
                    tgChfCntr.iOHDTime(1) = ilEntryTime1
                    Select Case slChfType    'tgChfCntr.sType
                        Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        Case Else   'Standard or Reservation
                            If tgChfCntr.sSchStatus = "M" Then
                                tgChfCntr.iHdChg = tgChfCntr.iHdChg Or HDSCHSTATUSMCHG
                            End If
                            'If tgChfCntr.iHdChg <> 0 Then
                            If (tgChfCntr.iHdChg <> 0) Or ((Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) <> "")) Then
                                tgChfCntr.sSchStatus = "A" 'Altered
                            End If
                    End Select
                End If
            End If
            '                          Altered                        Interrupted                  New                               Manual
'            tgChfCntr.sSellNet = tgSpf.sSSellNet
            If ((sgCntrScreen = "Orders") And (imStep1Index = 2)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Then
                'Retain as previously set or as set in CRevNo Jim:5/19/97
            Else
                tgChfCntr.sPrint = "N"  'New
            End If
            'tgChfCntr.sExpired = "N"


            tgChfCntr.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            'Dan 11/16/15 insertion orders
            gPackDate "", tgChfCntr.iXMLSentDate(0), tgChfCntr.iXMLSentDate(1)
            gPackTime "12am", tgChfCntr.iXMLSentTime(0), tgChfCntr.iXMLSentTime(1)
            tgChfCntr.iXMLSentExtRevNo = 0
            'tgChfCntr.iRemoteID = tgUrf(0).iRemoteUserID
            'tgChfCntr.lAutoCode = tgChfCntr.lCode
            '10967
            mAdjustDeliveryStatus
            ilRet = btrInsert(hmCHF, tgChfCntr, imCHFRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert: Contract)"
        Else 'Old record-Update
            'Read contract in
            'ilRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, lmChfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            tmChfSrchKey.lCode = lmCntrChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual Chf)", Contract
            On Error GoTo 0
            ilNewChf = False
            tgChfCntr.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
            tgChfCntr.sDelete = "N"
            If mEDSSetNeg2(0, "", bmAnyVehicleOnInsertion) Then
                tgChfCntr.iEDSSentExtRevNo = -2 '-2 indicates not used
                gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                tgChfCntr.iEDSSentUrfCode = 0
            ElseIf mEDSSetNeg1(tmChf.iEDSSentExtRevNo, 0, "", bmAnyVehicleOnInsertion) Then
                tgChfCntr.iEDSSentExtRevNo = -1
                gPackDate "1/1/1970", tgChfCntr.iEDSSentDate(0), tgChfCntr.iEDSSentDate(1)
                gPackTime "12am", tgChfCntr.iEDSSentTime(0), tgChfCntr.iEDSSentTime(1)
                tgChfCntr.iEDSSentUrfCode = 0
            Else
                tgChfCntr.iEDSSentExtRevNo = tmChf.iEDSSentExtRevNo
                tgChfCntr.iEDSSentDate(0) = tmChf.iEDSSentDate(0)
                tgChfCntr.iEDSSentDate(1) = tmChf.iEDSSentDate(1)
                tgChfCntr.iEDSSentTime(0) = tmChf.iEDSSentTime(0)
                tgChfCntr.iEDSSentTime(1) = tmChf.iEDSSentTime(1)
                tgChfCntr.iEDSSentUrfCode = tmChf.iEDSSentUrfCode
            End If
            If (sgCntrScreen = "Orders") Then
                Select Case slChfType    'tgChfCntr.sType
                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        tgChfCntr.sSchStatus = "M"
                        If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                            tgChfCntr.sStatus = "O"
                        End If
                    Case Else   'Standard or Reservation
                        If tgChfCntr.sSchStatus = "P" Then
                            tgChfCntr.iHdChg = tgChfCntr.iHdChg Or HDSCHSTATUSPCHG
                        End If
                        If tgChfCntr.sSchStatus = "M" Then
                            tgChfCntr.iHdChg = tgChfCntr.iHdChg Or HDSCHSTATUSMCHG
                        End If
                        If (tgChfCntr.iHdChg <> 0) And (tgChfCntr.sSchStatus <> "N") Then
                            tgChfCntr.sSchStatus = "A" 'Altered
                        End If
                End Select
            Else
                Select Case slChfType    'tgChfCntr.sType
                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        tgChfCntr.sSchStatus = "M"
                        If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                            tgChfCntr.sStatus = "O"
                        End If
                End Select
                If (tgChfCntr.sStatus <> "O") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "N") Then
                    tgChfCntr.sSchStatus = "P"  'Prevent Scheduling
                    tgChfCntr.iPropDate(0) = ilEntryDate0
                    tgChfCntr.iPropDate(1) = ilEntryDate1
                    tgChfCntr.iPropTime(0) = ilEntryTime0
                    tgChfCntr.iPropTime(1) = ilEntryTime1
                Else
                    Select Case slChfType    'tgChfCntr.sType
                        Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                            tgChfCntr.sSchStatus = "M"
                            If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                                tgChfCntr.sStatus = "O"
                            End If
                       Case Else   'Standard or Reservation
                            'tgChfCntr.sSchStatus = "A"
                            If tgChfCntr.sSchStatus = "P" Then
                                tgChfCntr.iHdChg = tgChfCntr.iHdChg Or HDSCHSTATUSPCHG
                            End If
                            If tgChfCntr.iHdChg <> 0 Then
                                tgChfCntr.sSchStatus = "A" 'Altered
                            End If
                    End Select
                End If
            End If
            If (tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
                If (smOrigStatus <> "O") And (smOrigStatus <> "H") And (smOrigStatus <> "G") And (smOrigStatus <> "N") Then
                    '5/20/08:  Retain original date so that the Sale activity report will reflict the change in the original week.
                    '          This was changed for Disney because of Sales Activity report and Advertiser credit status.
                    '          While waiting for the credit status to be changed, the contract was changed causing the dollar to move
                    '          from one week to another
                    '6/30/08:  re instated until a new solution to designed see TTP 3433
                    tgChfCntr.iOHDDate(0) = ilEntryDate0
                    tgChfCntr.iOHDDate(1) = ilEntryDate1
                    tgChfCntr.iOHDTime(0) = ilEntryTime0
                    tgChfCntr.iOHDTime(1) = ilEntryTime1
                End If
            End If


            tgChfCntr.iUrfCode = tgUrf(0).iCode
            '10967 if this is a contract that's been sent to vendor it may need to be flagged to be resent.
            mAdjustDeliveryStatus
            ilRet = btrUpdate(hmCHF, tgChfCntr, imCHFRecLen)
            If ilRet = BTRV_ERR_CONFLICT Then
                tmChfSrchKey.lCode = tgChfCntr.lCode
                ilErrRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilErrRet, "mSaveRec (GetEqual- Chf)", Contract
                On Error GoTo 0
            End If
            slMsg = "mSaveRec (btrUpdate: Contract)"
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, Contract
    On Error GoTo 0

    If ilNewHeader Then
        Do
            tgChfCntr.lSpotChfCode = 0
            If ((sgCntrScreen = "Orders") And (imStep1Index < 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index < 2)) Then
                'tgChfCntr.iOrigRemoteID = tgChfCntr.iRemoteID
                'tgChfCntr.lOrigAutoCode = tgChfCntr.lCode
                If tgUrf(0).iRemoteUserID > 0 Then
                    tgChfCntr.sNoAssigned = "N"
                Else
                    tgChfCntr.sNoAssigned = "Y"
                End If
            End If
            ilRet = btrUpdate(hmCHF, tgChfCntr, imCHFRecLen)
            slMsg = "mSaveRec (btrUpdate:Contract RU ID)"
        Loop While ilRet = BTRV_ERR_CONFLICT
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Contract
        On Error GoTo 0
    End If
    mAddSbfRec hmSbf, hmRvf, hmPhf, imTerminate
    igTabMapIndex = 1
    gAddPcfRec
    igTabMapIndex = 0
    If imTerminate Then
        Screen.MousePointer = vbDefault    'Default
        imInSave = False
        mSaveRecFunction = False
        'ilRet = btrAbortTrans(hmCHF)
        Exit Function
    End If
    ilUpdateLine = False
    If ((sgCntrScreen = "Orders") And (imStep1Index = 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 5)) Then
        'Disallow any line changes 7/2/98- Jim
    Else
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            ilUpdateLine = False
            If tgClfCntr(ilClf).iStatus >= 0 Then   '(0=New; 1=Old & retain; 2=Old & delete;-1=ignore)
                For ilLoop = 0 To 4 Step 1
                    If (imLnCommShowOn(ilLoop, ilClf + 1) <> imLnInitCommShowOn(ilLoop, ilClf + 1)) Then
                        ilUpdateLine = True
                    End If
                Next ilLoop
                If smLnComment(ilClf + 1) <> smLnInitComment(ilClf + 1) Then
                    ilUpdateLine = True
                End If
                Select Case slChfType    'tgChfCntr.sType
                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                        Select Case smOrigType
                            Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                            Case Else
                                If smOrigType <> "" Then
                                    ilUpdateLine = True
                                End If
                        End Select
                    Case Else
                        Select Case smOrigType
                            Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                                ilUpdateLine = True
                            Case Else
                        End Select
                End Select
                If (imStep1Index <= 1) Or tgClfCntr(ilClf).iStatus = 0 Then 'New selected
                    ilUpdateLine = True
                Else 'Old record-Update
                    If tgClfCntr(ilClf).iStatus = 2 Then    'Remove line only
                        ilUpdateLine = True
                    ElseIf (sgCntrScreen <> "Orders") And (imStep1Index = 2) And ((tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N")) Then
                        ilUpdateLine = True 'Force schedule status change
                    ElseIf (((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Or ((sgCntrScreen = "Orders") And (imStep1Index = 2))) And (tgClfCntr(ilClf).ClfRec.sSchStatus = "P") And ((tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N")) Then
                        ilUpdateLine = True 'Force schedule status change
                    Else
                            ilUpdateLine = True
                    End If
                End If
            End If
            'If "Creating New Version of An Existing Proposal" all lines of the existing
            'proposal must be added as new to this proposal
            If (sgCntrScreen <> "Orders") And (imStep1Index = 3) Then
                If tgClfCntr(ilClf).iStatus = 1 Then   '(0=New; 1=Old & retain; 2=Old & delete;-1=ignore)
                    ilUpdateLine = True
                    tgClfCntr(ilClf).iStatus = 0
                    'Set Flights as new
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            tgCffCntr(ilCff).iStatus = 0
                        Else
                            tgCffCntr(ilCff).iStatus = -1
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                End If
            End If
            If Not ilUpdateLine Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If tgCffCntr(ilCff).iStatus >= 0 Then
                        If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 0) Then 'New selected
                            ilUpdateLine = True
                            Exit Do
                        Else 'Old record-Update
                            If tgCffCntr(ilCff).iStatus = 2 Then
                                ilUpdateLine = True
                                Exit Do
                            Else
                                    ilUpdateLine = True
                                    Exit Do
                            End If
                        End If
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
            End If
            If ilUpdateLine Then
                'If revise and Old line from other order, then add as new
                If (((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Or ((sgCntrScreen = "Orders") And (imStep1Index = 2))) And (tgClfCntr(ilClf).iStatus = 1) And (tgClfCntr(ilClf).ClfRec.lChfCode <> tgChfCntr.lCode) Then
                    tgClfCntr(ilClf).iStatus = 0
                    'Set Flights as new
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            tgCffCntr(ilCff).iStatus = 0
                        Else
                            tgCffCntr(ilCff).iStatus = -1
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                End If
            End If
            If ilUpdateLine Then
                ilAnyLineChgd = True
                ilCxfState = 0
                Do
                    'Update comment if required
                    tmCxf.lCode = 0 'Must create a new comment to be associated with the line
                    tmCxf.sComType = "L"
                    If (imLnCommShowOn(0, ilClf + 1) = 0) Then
                        tmCxf.sShProp = "Y"
                    Else
                        tmCxf.sShProp = "N"
                    End If
                    If (imLnCommShowOn(1, ilClf + 1) = 0) Then
                        tmCxf.sShOrder = "Y"
                    Else
                        tmCxf.sShOrder = "N"
                    End If
                    If (imLnCommShowOn(2, ilClf + 1) = 0) Then
                        tmCxf.sShSpot = "Y"
                    Else
                        tmCxf.sShSpot = "N"
                    End If
                    If (imLnCommShowOn(3, ilClf + 1) = 0) Then
                        tmCxf.sShInv = "Y"
                    Else
                        tmCxf.sShInv = "N"
                    End If
                    If (imLnCommShowOn(4, ilClf + 1) = 0) Then
                        tmCxf.sShInsertion = "Y"
                    Else
                        tmCxf.sShInsertion = "N"
                    End If
                    'tmCxf.iStrLen = Len(Trim$(smLnComment(ilClf + 1)))
                    tmCxf.sComment = Trim$(smLnComment(ilClf + 1)) & Chr$(0) ' & Chr$(0) 'sgTB
                    'imCxfRecLen = 25 + Len(Trim$(tmCxf.sComment)) + 2    '25 = fixed record length; 2=Length value which is part of the variable record
                    imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))     '25 = fixed record length; 2=Length value which is part of the variable record
                    If tmCxf.lCode = 0 Then  'New selected
                        ''If imCxfRecLen - 2 > 27 Then  '-2 so control character at end not counted
                        'If Len(Trim$(tmCxf.sComment)) > 2 Then  '2 so control character at end not counted
                        If Trim$(smLnComment(ilClf + 1)) <> "" Then
                            ilCxfState = 1
                            tmCxf.lCode = 0 'Autoincrement
                            tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmCxf.lAutoCode = tmCxf.lCode
                            ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                            ilUpdateLine = True
                        Else
                            tmCxf.lCode = 0
                            ilRet = BTRV_ERR_NONE
                        End If
                        slMsg = "mSaveRec (btrInsert: Comment)"
                    Else 'Old record-Update
                        ''If imCxfRecLen - 2 > 27 Then  '-2 so the control character at end is not counted
                        'If Len(Trim$(tmCxf.sComment)) > 2 Then  '2 so the control character at end is not counted
                        If Trim$(smLnComment(ilClf + 1)) <> "" Then
                            If tmCxf.lCode = 0 Then
                                ilCxfState = 1
                                tmCxf.lCode = 0 'Autoincrement
                                tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                                tmCxf.lAutoCode = tmCxf.lCode
                                ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                                ilUpdateLine = True
                            Else
                                ilCxfState = 2
                                tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                                gPackDate smSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                                gPackTime smSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                                ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                            End If
                        Else
                            If tmCxf.lCode <> 0 Then
                                ilCxfState = 3
                                ilRet = btrDelete(hmCxf)
                                ilUpdateLine = True
                            End If
                            tmCxf.lCode = 0
                        End If
                        slMsg = "mSaveRec (btrUpdate: Comment)"
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Contract
                On Error GoTo 0
                tgClfCntr(ilClf).ClfRec.lCxfCode = tmCxf.lCode
                If ilCxfState = 1 Then
                    Do
                        tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                        tmCxf.lAutoCode = tmCxf.lCode
                        tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                        gPackDate smSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                        gPackTime smSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                        imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                        ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                        slMsg = "mSaveRec (btrUpdate:Comment RU ID)"
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, Contract
                    On Error GoTo 0
                ElseIf ilCxfState = 3 Then
                End If
                If (imStep1Index <= 1) Or tgClfCntr(ilClf).iStatus = 0 Then 'New selected
                Else
                    'ilRet = btrGetDirect(hmClf, tlClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                    ilRet = btrGetEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Clf)", Contract
                    On Error GoTo 0
                    slLineSchStatus = tgClfCntr(ilClf).ClfRec.sSchStatus
                End If
                Do  'Loop until record updated or added
                    If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 3)) Then 'New selected
                        ilNewClf = True
                        tgClfCntr(ilClf).ClfRec.lChfCode = tgChfCntr.lCode
                        tgClfCntr(ilClf).ClfRec.iCntRevNo = tgChfCntr.iCntRevNo
                        tgClfCntr(ilClf).ClfRec.iPropVer = tgChfCntr.iPropVer
                        tgClfCntr(ilClf).ClfRec.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                        tgClfCntr(ilClf).ClfRec.iEntryDate(0) = ilEntryDate0
                        tgClfCntr(ilClf).ClfRec.iEntryDate(1) = ilEntryDate1
                        tgClfCntr(ilClf).ClfRec.iEntryTime(0) = ilEntryTime0
                        tgClfCntr(ilClf).ClfRec.iEntryTime(1) = ilEntryTime1
                        tgClfCntr(ilClf).ClfRec.sDelete = "N"
                        If (sgCntrScreen = "Orders") Then
                            Select Case slChfType    'tgChfCntr.sType
                                Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                                    tgClfCntr(ilClf).ClfRec.sSchStatus = "M"
                                Case Else   'Standard or Reservation
                                    If imStep1Index <= 1 Then
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "N"
                                    Else
                                        'It could be new or altered, set as altered
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "A" 'Altered
                                        ilChgHdStatus = True
                                    End If
                            End Select
                        Else
                            If (tgChfCntr.sStatus <> "O") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "N") Then
                                tgClfCntr(ilClf).ClfRec.sSchStatus = "P"  'Prevent Scheduling
                            Else
                                Select Case slChfType    'tgChfCntr.sType
                                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "M"
                                    Case Else   'Standard or Reservation
                                        'It could be new or altered, set as altered
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "A"
                                        ilChgHdStatus = True
                                End Select
                            End If
                        End If

                        If (imCntrCluster) Or (imCntrRep) Then     'Force swap only
                            tgClfCntr(ilClf).ClfRec.sSchStatus = "F"
                        End If
                        
                        If mEDSSetNeg2(tgClfCntr(ilClf).ClfRec.iVefCode, tgClfCntr(ilClf).ClfRec.sType, bmAnyVehicleOnInsertion) Then
                            tgClfCntr(ilClf).ClfRec.iEDSSentExtRevNo = -2 '-2 indicates not used
                        Else
                            tgClfCntr(ilClf).ClfRec.iEDSSentExtRevNo = -1 '-2 indicates not used
                        End If
                        
                        tgClfCntr(ilClf).ClfRec.lCode = 0
                        ilRet = btrInsert(hmClf, tgClfCntr(ilClf).ClfRec, imClfRecLen, INDEXKEY0)
                        slMsg = "mSaveRec (btrInsert: Contract Line)"
                    Else 'Old record-Update
                        ilNewClf = False
                        slMsg = "mSaveRec (btrGetEqual: Contract Line)"
                        'ilRet = btrGetDirect(hmClf, tlClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                        ilRet = btrGetEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                        'Replace altered lines that have not been scheduled
                        If (sgCntrScreen = "Orders") Then
                            Do
                                tmCxfSrchKey.lCode = tlClf.lCxfCode
                                If tmCxfSrchKey.lCode <> 0 Then
                                    tmCxf.sComment = ""
                                    imCxfRecLen = Len(tmCxf) '5027
                                    ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cxf)", Contract
                                    On Error GoTo 0
                                    ilRet = btrDelete(hmCxf)
                                    If ilRet <> BTRV_ERR_CONFLICT Then
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cxf)", Contract
                                        On Error GoTo 0
                                    End If
                                Else
                                    Exit Do
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            slMsg = "mSaveRec (btrDelete: Contract Line)"
                            ilRet = btrDelete(hmClf)
                        Else
                            Do
                                tmCxfSrchKey.lCode = tlClf.lCxfCode
                                If tmCxfSrchKey.lCode <> 0 Then
                                    tmCxf.sComment = ""
                                    imCxfRecLen = Len(tmCxf) '5027
                                    ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cxf)", Contract
                                    On Error GoTo 0
                                    ilRet = btrDelete(hmCxf)
                                    If ilRet <> BTRV_ERR_CONFLICT Then
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cxf)", Contract
                                        On Error GoTo 0
                                    End If
                                Else
                                    Exit Do
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            slMsg = "mSaveRec (btrDelete: Contract Line)"
                            ilRet = btrDelete(hmClf)
                        End If
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                        If (sgCntrScreen = "Orders") Then
                            Select Case slChfType    'tgChfCntr.sType
                                Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                                    tgClfCntr(ilClf).ClfRec.sSchStatus = "M"
                                Case Else   'Standard or Reservation
                                    'It could be new or altered, set as altered
                                    tgClfCntr(ilClf).ClfRec.sSchStatus = "A"    'Altered
                                    ilChgHdStatus = True
                            End Select
                        Else
                            If (tgChfCntr.sStatus <> "O") And (tgChfCntr.sStatus <> "H") And (tgChfCntr.sStatus <> "G") And (tgChfCntr.sStatus <> "N") Then
                                tgClfCntr(ilClf).ClfRec.sSchStatus = "P"  'Prevent Scheduling
                            Else
                                Select Case slChfType    'tgChfCntr.sType
                                    Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "M"
                                    Case Else   'Standard or Reservation
                                        'It could be new or altered, set as altered
                                        tgClfCntr(ilClf).ClfRec.sSchStatus = "A"    'Altered
                                        ilChgHdStatus = True
                                End Select
                            End If
                        End If
                        If tgClfCntr(ilClf).iStatus = 2 Then    'Remove line only
                            tgClfCntr(ilClf).ClfRec.sDelete = "Y"
                        Else
                            tgClfCntr(ilClf).ClfRec.sDelete = "N"
                        End If
                        tgClfCntr(ilClf).ClfRec.lChfCode = tgChfCntr.lCode
                        tgClfCntr(ilClf).ClfRec.iCntRevNo = tgChfCntr.iCntRevNo
                        tgClfCntr(ilClf).ClfRec.iPropVer = tgChfCntr.iPropVer
                        tgClfCntr(ilClf).ClfRec.iEntryDate(0) = ilEntryDate0
                        tgClfCntr(ilClf).ClfRec.iEntryDate(1) = ilEntryDate1
                        tgClfCntr(ilClf).ClfRec.iEntryTime(0) = ilEntryTime0
                        tgClfCntr(ilClf).ClfRec.iEntryTime(1) = ilEntryTime1
                        tgClfCntr(ilClf).ClfRec.iUrfCode = tgUrf(0).iCode 'Use first record retained for user

                        If (imCntrCluster) Or (imCntrRep) Then     'Force swap only
                            tgClfCntr(ilClf).ClfRec.sSchStatus = "F"
                        End If
                        If mEDSSetNeg2(tgClfCntr(ilClf).ClfRec.iVefCode, tgClfCntr(ilClf).ClfRec.sType, bmAnyVehicleOnInsertion) Then
                            tgClfCntr(ilClf).ClfRec.iEDSSentExtRevNo = -2 '-2 indicates not used
                        ElseIf mEDSSetNeg1(tlClf.iEDSSentExtRevNo, tgClfCntr(ilClf).ClfRec.iVefCode, tgClfCntr(ilClf).ClfRec.sType, bmAnyVehicleOnInsertion) Then
                            tgClfCntr(ilClf).ClfRec.iEDSSentExtRevNo = -1
                        Else
                            tgClfCntr(ilClf).ClfRec.iEDSSentExtRevNo = tlClf.iEDSSentExtRevNo
                        End If
                        tgClfCntr(ilClf).ClfRec.lCode = 0
                        ilRet = btrInsert(hmClf, tgClfCntr(ilClf).ClfRec, imClfRecLen, INDEXKEY0)
                        slMsg = "mSaveRec (Insert- Clf)"
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Contract
                On Error GoTo 0
                ilRet = btrGetPosition(hmClf, tgClfCntr(ilClf).lRecPos)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, "mSaveRec (GetPosition- Clf)", Contract
                On Error GoTo 0
                'Delete all old first if changing existing
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If tgCffCntr(ilCff).iStatus >= 0 Then
                        Do  'Loop until record updated or added
                            If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 0) Then 'New selected
                            Else 'Old record-Update
                                slMsg = "mSaveRec (btrGetEqual: Contract Line Flight)"
                                'ilRet = btrGetDirect(hmCff, tlCff, imCffRecLen, tgCffCntr(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                tmCffSrchKey1.lCode = tgCffCntr(ilCff).CffRec.lCode
                                ilRet = btrGetEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, slMsg, Contract
                                On Error GoTo 0
                                'Replace altered lines that have not been scheduled
                                If (sgCntrScreen = "Orders") Then
                                    slMsg = "mSaveRec (btrDelete: Contract Line Flight)"
                                    ilRet = btrDelete(hmCff)
                                Else
                                    slMsg = "mSaveRec (btrDelete: Contract Line Flight)"
                                    ilRet = btrDelete(hmCff)
                                End If
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, slMsg, Contract
                                On Error GoTo 0
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                        'ilRet = btrGetPosition(hmCff, tgCffCntr(ilCff).lRecPos)
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop

                ilCgf = tgClfCntr(ilClf).iFirstCgf
                Do While ilCgf <> -1
                    If tgCgfCntr(ilCgf).iStatus >= 0 Then
                        Do  'Loop until record updated or added
                            If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or (tgCgfCntr(ilCgf).iStatus = 0) Then 'New selected
                            Else 'Old record-Update
                                slMsg = "mSaveRec (btrGetEqual: Contract Line Events)"
                                tmCgfSrchKey.lCode = tgCgfCntr(ilCgf).CgfRec.lCode
                                ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, slMsg, Contract
                                On Error GoTo 0
                                If (sgCntrScreen = "Orders") Then
                                    slMsg = "mSaveRec (btrDelete: Contract Line Events)"
                                    ilRet = btrDelete(hmCgf)
                                Else
                                    slMsg = "mSaveRec (btrDelete: Contract Line Events)"
                                    ilRet = btrDelete(hmCgf)
                                End If
                                On Error GoTo mSaveRecErr
                                gBtrvErrorMsg ilRet, slMsg, Contract
                                On Error GoTo 0
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                    End If
                    ilCgf = tgCgfCntr(ilCgf).iNextCgf
                Loop

                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus >= 0) And (tgCffCntr(ilCff).iStatus <= 1) Then
                        Do  'Loop until record updated or added
                            If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 0) Then 'New selected
                                tgCffCntr(ilCff).CffRec.lChfCode = tgChfCntr.lCode
                                tgCffCntr(ilCff).CffRec.iClfLine = tgClfCntr(ilClf).ClfRec.iLine
                                tgCffCntr(ilCff).CffRec.iCntRevNo = tgClfCntr(ilClf).ClfRec.iCntRevNo
                                tgCffCntr(ilCff).CffRec.iPropVer = tgClfCntr(ilClf).ClfRec.iPropVer
                                tgCffCntr(ilCff).CffRec.sDelete = "N"
                                tgCffCntr(ilCff).CffRec.lCode = 0
                                ilRet = btrInsert(hmCff, tgCffCntr(ilCff).CffRec, imCffRecLen, INDEXKEY1)
                                slMsg = "mSaveRec (btrInsert: Contract Line Flight)"
                            Else 'Old record-Update
                                tgCffCntr(ilCff).CffRec.lChfCode = tgChfCntr.lCode
                                tgCffCntr(ilCff).CffRec.iCntRevNo = tgClfCntr(ilClf).ClfRec.iCntRevNo
                                tgCffCntr(ilCff).CffRec.iPropVer = tgClfCntr(ilClf).ClfRec.iPropVer
                                If (tgCffCntr(ilCff).iStatus = 2) Or (tgClfCntr(ilClf).iStatus = 2) Then
                                    ilRet = BTRV_ERR_NONE   'Don't retain flights if line is deleted
                                Else
                                    tgCffCntr(ilCff).CffRec.lCode = 0
                                    ilRet = btrInsert(hmCff, tgCffCntr(ilCff).CffRec, imCffRecLen, INDEXKEY1)
                                    slMsg = "mSaveRec (btrInsert: Contract Line Flight)"
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                        ilRet = btrGetPosition(hmCff, tgCffCntr(ilCff).lRecPos)
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, "mSaveRec (GetPosition-Cff)", Contract
                        On Error GoTo 0
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop

                ilCgf = tgClfCntr(ilClf).iFirstCgf
                Do While ilCgf <> -1
                    If (tgCgfCntr(ilCgf).iStatus >= 0) And (tgCgfCntr(ilCgf).iStatus <= 1) And (tgCgfCntr(ilCgf).CgfRec.iNoSpots > 0) Then
                        Do  'Loop until record updated or added
                            If (imStep1Index <= 1) Or (tgClfCntr(ilClf).iStatus = 0) Or (tgCgfCntr(ilCgf).iStatus = 0) Then 'New selected
                                tgCgfCntr(ilCgf).CgfRec.lClfCode = tgClfCntr(ilClf).ClfRec.lCode
                                tgCgfCntr(ilCgf).CgfRec.lCode = 0
                                ilRet = btrInsert(hmCgf, tgCgfCntr(ilCgf).CgfRec, imCgfRecLen, INDEXKEY0)
                                slMsg = "mSaveRec (btrInsert: Contract Line Event)"
                            Else 'Old record-Update
                                tgCgfCntr(ilCgf).CgfRec.lClfCode = tgClfCntr(ilClf).ClfRec.lCode
                                tgCgfCntr(ilCgf).CgfRec.lCode = 0
                                If (tgCgfCntr(ilCgf).iStatus = 2) Or (tgClfCntr(ilClf).iStatus = 2) Then
                                    ilRet = BTRV_ERR_NONE   'Don't retain flights if line is deleted
                                Else
                                    ilRet = btrInsert(hmCgf, tgCgfCntr(ilCgf).CgfRec, imCgfRecLen, INDEXKEY0)
                                    slMsg = "mSaveRec (btrInsert: Contract Line Event)"
                                End If
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                    End If
                    ilCgf = tgCgfCntr(ilCgf).iNextCgf
                Loop
                'Update region if defined
                If tgClfCntr(ilClf).ClfRec.lRafCode > 0 Then
                    tmRafSrchKey.lCode = tgClfCntr(ilClf).ClfRec.lRafCode
                    ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tmRaf.sAssigned = "Y"
                        ilRet = btrUpdate(hmRaf, tmRaf, imRafRecLen)
                    End If
                End If
            End If
        Next ilClf
        If Not ilNewHeader Then
            Do
                tmMsfSrchKey2.lChfCode = tgChfCntr.lCode
                ilRet = btrGetGreaterOrEqual(hmMsf, tmMsf, imMsfRecLen, tmMsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet <> BTRV_ERR_NONE) Or (tmMsf.lChfCode <> tgChfCntr.lCode) Then
                    Exit Do
                End If
                Do
                    tmMgfSrchKey1.lMsfCode = tmMsf.lCode
                    tmMgfSrchKey1.iGameNo = 0
                    ilRet = btrGetGreaterOrEqual(hmMgf, tmMgf, imMgfRecLen, tmMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet <> BTRV_ERR_NONE) Then
                        Exit Do
                    End If
                    If tmMgf.lMsfCode <> tmMsf.lCode Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hmMgf)
                Loop
                Do
                    tmCxfSrchKey.lCode = tmMsf.lCxfCode
                    If tmCxfSrchKey.lCode <> 0 Then
                        tmCxf.sComment = ""
                        imCxfRecLen = Len(tmCxf) '5027
                        ilRet = gCXFGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrDelete(hmCxf)
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = btrDelete(hmMsf)
            Loop While (ilRet = BTRV_ERR_CONFLICT) Or (ilRet = BTRV_ERR_NONE)
        End If
        For ilMsf = 0 To UBound(tgMsfCntr) - 1 Step 1
            If (tgMsfCntr(ilMsf).iStatus >= 0) And (tgMsfCntr(ilMsf).iFirstMgf <> -1) Then
                If tgMsfCntr(ilMsf).iCxfIndex >= 0 Then
                    If Trim$(sgMsfCntrCxf(tgMsfCntr(ilMsf).iCxfIndex)) <> "" Then
                        tmCxf.lCode = 0 'Must create a new comment to be associated with the line
                        tmCxf.sComType = "G"    'Game Multi-Media
                        tmCxf.sShProp = "N"
                        tmCxf.sShOrder = "N"
                        tmCxf.sShSpot = "N"
                        tmCxf.sShInv = "N"
                        tmCxf.sShInsertion = "N"
                        'tmCxf.iStrLen = Len(Trim$(sgMsfCntrCxf(tgMsfCntr(ilMsf).iCxfIndex)))
                        tmCxf.sComment = Trim$(sgMsfCntrCxf(tgMsfCntr(ilMsf).iCxfIndex)) & Chr$(0) '& Chr$(0) 'sgTB
                        imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))     '25 = fixed record length; 2=Length value which is part of the variable record
                        'If Len(Trim$(tmCxf.sComment)) > 2 Then  '2 so control character at end not counted
                        If Trim$(sgMsfCntrCxf(tgMsfCntr(ilMsf).iCxfIndex)) <> "" Then
                            tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmCxf.lAutoCode = tmCxf.lCode
                            ilRet = gCXFInsert(hmCxf, tmCxf, imCxfRecLen, INDEXKEY0)
                            If ilRet = BTRV_ERR_NONE Then
                                tmCxf.iRemoteID = tgUrf(0).iRemoteUserID
                                tmCxf.lAutoCode = tmCxf.lCode
                                tmCxf.iSourceID = tgUrf(0).iRemoteUserID
                                gPackDate smSyncDate, tmCxf.iSyncDate(0), tmCxf.iSyncDate(1)
                                gPackTime smSyncTime, tmCxf.iSyncTime(0), tmCxf.iSyncTime(1)
                                imCxfRecLen = Len(tmCxf) '- Len(tmCxf.sComment) + Len(Trim$(tmCxf.sComment))
                                ilRet = gCXFUpdate(hmCxf, tmCxf, imCxfRecLen)
                            End If
                        Else
                            tmCxf.lCode = 0
                            ilRet = BTRV_ERR_NONE
                        End If
                        tgMsfCntr(ilMsf).MsfRec.lCxfCode = tmCxf.lCode
                    End If
                End If
                Do  'Loop until record updated or added
                    tgMsfCntr(ilMsf).MsfRec.lChfCode = tgChfCntr.lCode
                    tgMsfCntr(ilMsf).MsfRec.lCode = 0
                    ilRet = btrInsert(hmMsf, tgMsfCntr(ilMsf).MsfRec, imMsfRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Contract Event Media)"
                Loop While ilRet = BTRV_ERR_CONFLICT
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, Contract
                On Error GoTo 0
                ilMgf = tgMsfCntr(ilMsf).iFirstMgf
                Do While ilMgf <> -1
                    If tgMgfCntr(ilMgf).iStatus >= 0 Then
                        Do
                            tgMgfCntr(ilMgf).MgfRec.lCode = 0
                            tgMgfCntr(ilMgf).MgfRec.lMsfCode = tgMsfCntr(ilMsf).MsfRec.lCode
                            ilRet = btrInsert(hmMgf, tgMgfCntr(ilMgf).MgfRec, imMgfRecLen, INDEXKEY0)
                            slMsg = "mSaveRec (btrInsert: Contract Media Events)"
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        On Error GoTo mSaveRecErr
                        gBtrvErrorMsg ilRet, slMsg, Contract
                        On Error GoTo 0
                    End If
                    ilMgf = tgMgfCntr(ilMgf).iNextMgf
                Loop
            End If
        Next ilMsf
    End If


    If (ilChgHdStatus) Then  'Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N")) And (tgChfCntr.sSchStatus = "F") Then
        Do
            tmChfSrchKey.lCode = tgChfCntr.lCode
            ilRet = btrGetEqual(hmCHF, tgChfCntr, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual-Chf)", Contract
            On Error GoTo 0
            tgChfCntr.iUrfCode = tgUrf(0).iCode
            tgChfCntr.sSchStatus = "A"
            If tgChfCntr.sStatus = "O" Then
                tgChfCntr.sStatus = "N"
            End If
            If tgChfCntr.sStatus = "H" Then
                tgChfCntr.sStatus = "G"
            End If
            ilRet = btrUpdate(hmCHF, tgChfCntr, imCHFRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mSaveRec (btrUpdate: Contract Status)"
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, Contract
        On Error GoTo 0
    End If
    'Clear up other proposal of this contract
    If (sgCntrScreen <> "Orders") Then      'Number can't be changed
        If imStep1Index <> 4 And imStep1Index <> 5 Then   '4=Revise an Order, 5 = change hold to order dont need to do anything to history
            'Was it changed to Hold or Order or Dead- if so remove other proposals
            If (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "D") Then
                'Read in total contract history- remove other proposals and set other reversion
                'to deleted
                ilCntRevNo = tgChfCntr.iCntRevNo
                tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
                tmChfSrchKey1.iCntRevNo = ilCntRevNo
                tmChfSrchKey1.iPropVer = 32000
                ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo)
                    ilRet = btrGetPosition(hmCHF, llChfRecPos)
                    llCntrChfCode = tmChf.lCode
                    If (tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Then
                        If tmChf.lCode <> tgChfCntr.lCode Then
                            ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, tmChf.lCode, False, tmChf, tgClfCntr(), tgCffCntr(), tgCgfCntr())
                            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                If (tgClfCntr(ilClf).ClfRec.iCntRevNo = tmChf.iCntRevNo) And (tgClfCntr(ilClf).ClfRec.iPropVer = tmChf.iPropVer) Then
                                    ilCff = tgClfCntr(ilClf).iFirstCff
                                    Do While ilCff <> -1
                                        'ilRet = btrGetDirect(hmCff, tmCff, imCffRecLen, tgCffCntr(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        tmCffSrchKey1.lCode = tgCffCntr(ilCff).CffRec.lCode
                                        ilRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "GetEqual- Cff", Contract
                                        On Error GoTo 0
                                        Do
                                            ilRet = btrDelete(hmCff)
                                            If ilRet = BTRV_ERR_CONFLICT Then
                                                'ilCRet = btrGetDirect(hmCff, tmCff, imCffRecLen, tgCffCntr(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                tmCffSrchKey1.lCode = tgCffCntr(ilCff).CffRec.lCode
                                                ilCRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                'tmRec = tmCff
                                                'ilCRet = gGetByKeyForUpdate("Cff", hmCff, tmRec)
                                                'tmCff = tmRec
                                            End If
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cff)", Contract
                                        On Error GoTo 0
                                        ilCff = tgCffCntr(ilCff).iNextCff
                                    Loop
                                    ilCgf = tgClfCntr(ilClf).iFirstCgf
                                    Do While ilCgf <> -1
                                        tmCgfSrchKey.lCode = tgCgfCntr(ilCgf).CgfRec.lCode
                                        ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "GetEqual- Cgf", Contract
                                        On Error GoTo 0
                                        Do
                                            ilRet = btrDelete(hmCgf)
                                            If ilRet = BTRV_ERR_CONFLICT Then
                                                tmCgfSrchKey.lCode = tgCgfCntr(ilCgf).CgfRec.lCode
                                                ilCRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                            End If
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        On Error GoTo mSaveRecErr
                                        gBtrvErrorMsg ilRet, "mSaveRec (Delete- Cgf)", Contract
                                        On Error GoTo 0
                                        ilCgf = tgCgfCntr(ilCgf).iNextCgf
                                    Loop
                                    'ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                                    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (GetEqual-Clf)", Contract
                                    On Error GoTo 0
                                    Do
                                        ilRet = btrDelete(hmClf)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            'ilCRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                                            ilCRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (Delete-Clf)", Contract
                                    On Error GoTo 0

                                End If
                            Next ilClf
                            'ilRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, llChfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            tmChfSrchKey.lCode = llCntrChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Chf)", Contract
                            On Error GoTo 0
                            ilRet = gAlertContractErase(tmChf.lCode)
                            ilRet = mDelPjf(tmChf.lCode, hmDsf)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (mDelPjf)", Contract
                            On Error GoTo 0
                            ilRet = mDelSbf(tmChf.lCode, hmSbf)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (mDelSbf)", Contract
                            On Error GoTo 0
                            Do
                                ilRet = btrDelete(hmCHF)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    'ilCRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, llChfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                                    tmChfSrchKey.lCode = llCntrChfCode
                                    ilCRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (Delete- Chf)", Contract
                            On Error GoTo 0
                        End If
                    End If
                    'ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    If tmChf.iPropVer > 1 Then
                        tmChf.iPropVer = tmChf.iPropVer - 1
                    Else
                        tmChf.iPropVer = 32000
                        ilCntRevNo = ilCntRevNo - 1
                    End If
                    tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
                    tmChfSrchKey1.iCntRevNo = ilCntRevNo
                    tmChfSrchKey1.iPropVer = tmChf.iPropVer '32000
                    ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrGetPosition(hmCHF, llTChfRecPos)
                        llTCntrChfCode = tmChf.lCode
                        If (ilRet = BTRV_ERR_NONE) And (llTCntrChfCode = llCntrChfCode) Then   '(llTChfRecPos = llChfRecPos) Then
                            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                    End If
                Loop
                'Change any previous Order to Delete (sSchStatus = "M" Manual)
                If (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "O") Then
                    ilRet = gObtainCntrPlusGame(hmCHF, hmClf, hmCff, hmCgf, tgChfCntr.lCode, True, tmChf, tgClfCntr(), tgCffCntr(), tgCgfCntr())
                    'ilRet = gObtainCntr(hmChf, hmClf, hmCff, tgChfCntr.lCode, False, tmChf, tgClfCntr(), tgCffCntr())
                    If ilRet Then
                        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                            'ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Clf)", Contract
                            On Error GoTo 0
                            tmChfSrchKey.lCode = tmClf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Chf)", Contract
                            On Error GoTo 0
                            If (tmChf.lCode <> tgChfCntr.lCode) And ((tmChf.sStatus = "H") Or (tmChf.sStatus = "O")) Then
                                If tmChf.sDelete <> "Y" Then
                                    Do
                                        tmChf.sDelete = "Y"
                                        ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
                                        If ilRet = BTRV_ERR_CONFLICT Then
                                            tmChfSrchKey.lCode = tmClf.lChfCode
                                            ilCRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                                        End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    On Error GoTo mSaveRecErr
                                    gBtrvErrorMsg ilRet, "mSaveRec (Update- Chf)", Contract
                                    On Error GoTo 0
                                End If
                                'Check if line is replaced- if so, then set line as delete
                                If tmClf.sDelete <> "Y" Then
                                    For ilLoop = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                        If (tgClfCntr(ilLoop).ClfRec.iLine = tmClf.iLine) And (tgClfCntr(ilLoop).ClfRec.iCntRevNo > tmClf.iCntRevNo) Then
                                            ilCff = tgClfCntr(ilClf).iFirstCff
                                            Do While ilCff <> -1
                                                'ilRet = btrGetDirect(hmCff, tmCff, imCffRecLen, tgCffCntr(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                tmCffSrchKey1.lCode = tgCffCntr(ilCff).CffRec.lCode
                                                ilRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                On Error GoTo mSaveRecErr
                                                gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Cff)", Contract
                                                On Error GoTo 0
                                                Do
                                                    tmCff.sDelete = "Y"
                                                    ilRet = btrUpdate(hmCff, tmCff, imCffRecLen)
                                                    If ilRet = BTRV_ERR_CONFLICT Then
                                                        'ilCRet = btrGetDirect(hmCff, tmCff, imCffRecLen, tgCffCntr(ilCff).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        tmCffSrchKey1.lCode = tgCffCntr(ilCff).CffRec.lCode
                                                        ilCRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                    End If
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                                On Error GoTo mSaveRecErr
                                                gBtrvErrorMsg ilRet, "mSaveRec (Update- Cff)", Contract
                                                On Error GoTo 0
                                                ilCff = tgCffCntr(ilCff).iNextCff
                                            Loop
                                            'ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                                            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                            On Error GoTo mSaveRecErr
                                            gBtrvErrorMsg ilRet, "mSaveRec (GetEqual- Clf)", Contract
                                            On Error GoTo 0
                                            Do
                                                tmClf.sDelete = "Y"
                                                ilRet = btrUpdate(hmClf, tmClf, imClfRecLen)
                                                If ilRet = BTRV_ERR_CONFLICT Then
                                                    'ilCRet = btrGetDirect(hmClf, tmClf, imClfRecLen, tgClfCntr(ilClf).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    tmClfSrchKey2.lCode = tgClfCntr(ilClf).ClfRec.lCode
                                                    ilCRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                                End If
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            On Error GoTo mSaveRecErr
                                            gBtrvErrorMsg ilRet, "mSaveRec (Update- Clf)", Contract
                                            On Error GoTo 0
                                            Exit For
                                        End If
                                    Next ilLoop
                                End If
                            End If
                        Next ilClf
                    End If
                End If
                ilRet = mDelPjf(tgChfCntr.lCode, hmDsf)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, "mSaveRec (mDelPjf)", Contract
                On Error GoTo 0
            Else
                'Was the potential set- if so, clear potential from other proposals
                'If set to Complete, set all others to Working (Removed 6/11/99- Jim request)
                If (tgChfCntr.iMnfPotnType <> 0) Then   'Or (tgChfCntr.sStatus = "C") Or (tgChfCntr.sStatus = "I") Then
                    ilCntRevNo = tgChfCntr.iCntRevNo
                    tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
                    tmChfSrchKey1.iCntRevNo = ilCntRevNo
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo) And (tmChf.iCntRevNo = tgChfCntr.iCntRevNo)
                        If tmChf.lCode <> tgChfCntr.lCode Then
                            ilRet = btrGetPosition(hmCHF, llChfRecPos)
                            llCntrChfCode = tmChf.lCode
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (GetPosition- Chf)", Contract
                            On Error GoTo 0
                            Do
                                'If (tgChfCntr.sStatus = "C") Or (tgChfCntr.sStatus = "I") Then
                                '    tmChf.sStatus = "W"
                                'End If
                                If tgChfCntr.iMnfPotnType <> 0 Then
                                    tmChf.iMnfPotnType = 0
                                End If
                                ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    'ilCRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, llChfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                                    tmChfSrchKey.lCode = llCntrChfCode
                                    ilCRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                               End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            On Error GoTo mSaveRecErr
                            gBtrvErrorMsg ilRet, "mSaveRec (Update- Chf)", Contract
                            On Error GoTo 0
                        End If
                        If tmChf.iPropVer > 1 Then
                            tmChf.iPropVer = tmChf.iPropVer - 1
                        Else
                            tmChf.iPropVer = 32000
                            ilCntRevNo = ilCntRevNo - 1
                        End If
                        tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
                        tmChfSrchKey1.iCntRevNo = ilCntRevNo
                        tmChfSrchKey1.iPropVer = tmChf.iPropVer
                        ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                        If ilRet = BTRV_ERR_NONE Then
                            ilRet = btrGetPosition(hmCHF, llTChfRecPos)
                            llTCntrChfCode = tmChf.lCode
                            If (ilRet = BTRV_ERR_NONE) And (llTCntrChfCode = llCntrChfCode) Then   '(llTChfRecPos = llChfRecPos) Then
                                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            End If
                        End If
                    Loop
                End If
            End If
        End If
    Else
        'If contract number changed, change all other contracts
        If (tgChfCntr.lCntrNo <> lmOrigCntrNo) And (lmOrigCntrNo <> -1) Then
            tmChfSrchKey1.lCntrNo = lmOrigCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmOrigCntrNo)
                If tmChf.lCode <> tgChfCntr.lCode Then
                    ilRet = btrGetPosition(hmCHF, llChfRecPos)
                    llCntrChfCode = tmChf.lCode
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, "mSaveRec (GetPosition- Chf)", Contract
                    On Error GoTo 0
                    Do
                        tmChf.lCntrNo = tgChfCntr.lCntrNo
                        ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
                        If ilRet = BTRV_ERR_CONFLICT Then
                            'ilCRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, llChfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                            tmChfSrchKey.lCode = llCntrChfCode
                            ilCRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, "mSaveRec (Update- Chf)", Contract
                    On Error GoTo 0
                End If
                'ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE)
                tmChfSrchKey1.lCntrNo = lmOrigCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Loop
            ilRet = mUpdateCntrNoInRvf(lmOrigCntrNo, tgChfCntr.lCntrNo, hmRvf, hmPhf)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (mUpdateCntrNoInRvf)", Contract
            On Error GoTo 0
        End If
    End If
    'If ((imCntrCluster) Or (imCntrRep) Or ((tgChfCntr.sNTRDefined = "Y") And (UBound(tgClfCntr) <= LBound(tgClfCntr)))) And ((ilChgHdStatus) Or (tgChfCntr.sSchStatus = "N") Or (tgChfCntr.sSchStatus = "A")) Then   'Force swap only
    If (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMVEHICLE) = "") Or (Contract.grdCPMBuy.TextMatrix(Contract.grdCPMBuy.FixedRows, CPMDAYPART) = "") Then
        If ((imCntrCluster) Or (imCntrRep) Or ((tgChfCntr.sNTRDefined = "Y") And (UBound(tgClfCntr) <= LBound(tgClfCntr)))) And ((ilChgHdStatus) Or (tgChfCntr.sSchStatus = "N") Or (tgChfCntr.sSchStatus = "A") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "G")) Then   'Force swap only
            Do
                tmChfSrchKey.lCode = tgChfCntr.lCode
                ilRet = btrGetEqual(hmCHF, tgChfCntr, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, "mSaveRec (GetEqual-Chf)", Contract
                On Error GoTo 0
                tgChfCntr.iUrfCode = tgUrf(0).iCode
                tgChfCntr.sSchStatus = "F"
                tgChfCntr.iHdChg = 0
                If (tgChfCntr.sStatus = "N") Then
                    tgChfCntr.sStatus = "O"
                End If
                If (tgChfCntr.sStatus = "G") Then
                    tgChfCntr.sStatus = "H"
                End If
                ilRet = btrUpdate(hmCHF, tgChfCntr, imCHFRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            slMsg = "mSaveRec (btrUpdate: Contract Status)"
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, Contract
            On Error GoTo 0
            ilRet = mSwapCodes(False, hmCHF, hmClf, hmCff, hmDsf, hmSbf, hmMsf)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (mSwapCodes)", Contract
            On Error GoTo 0
            ilRet = mCntrSetDeleteFlag(hmCHF, hmClf, hmCff)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (mCntrSetDeleteFlag)", Contract
            On Error GoTo 0
            ilChgHdStatus = False
            ilAnyLineChgd = True
        End If
    End If
    If imCntrRep Or blTreatAsRep Then
        '8033 8132
        If (mIsStationXmlAllowedForOrder()) Then
            tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo) Then
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChf.lCode, False, tgChfCntr, tgClfCntr(), tgCffCntr())
                mMoveLnRecToCtrl imTerminate, tmRegionCode()
                Contract.mExportAsFile True
            Else
                gMsgBox "Station Xml-could not create files.  mSaverec-could not obtain contract."
            End If
        ElseIf blIsMissingStationRights Then
            gMsgBox "Station Xml-could not create files.  Some vehicles are marked as station xml, but they are not allowed."
            If Not ogStationXmlErrors Is Nothing Then
                ogStationXmlErrors.WriteWarning "Station Xml-could not create files.  Some vehicles are marked as station xml, but they are not allowed.", True
            End If
        End If
    ElseIf mIsStationXmlAllowedForOrder() Then
        gMsgBox "Station Xml-could not create files.  Some vehicles are marked as station xml, but they are not properly defined."
        If Not ogStationXmlErrors Is Nothing Then
            ogStationXmlErrors.WriteWarning "Station Xml-could not create files.  Some vehicles are marked as station xml, but they are not properly defined.", True
        End If
    End If
    If (tgChfCntr.sSchStatus = "M") Then
        If (tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "H") Then
            ilRet = mSwapCodes(False, hmCHF, hmClf, hmCff, hmDsf, hmSbf, hmMsf)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (mSwapCodes)", Contract
            On Error GoTo 0
            ilRet = mCntrSetDeleteFlag(hmCHF, hmClf, hmCff)
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, "mSaveRec (mCntrSetDeleteFlag)", Contract
            On Error GoTo 0
        End If
    Else
        If (Not ilAnyLineChgd) And (tgChfCntr.iHdChg = 0) And ((smOrigStatus = "O") Or (smOrigStatus = "H")) And (((sgCntrScreen = "Orders") And (imStep1Index = 2)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4))) Then
            If (tgChfCntr.sSchStatus <> "P") And (tgChfCntr.sSchStatus <> "I") And (tgChfCntr.sSchStatus <> "A") And (tgChfCntr.sSchStatus <> "N") Then
                ilRet = mSwapCodes(True, hmCHF, hmClf, hmCff, hmDsf, hmSbf, hmMsf)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, "mSaveRec (mSwapCodes)", Contract
                On Error GoTo 0
                ilRet = mCntrSetDeleteFlag(hmCHF, hmClf, hmCff)
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, "mSaveRec (mCntrSetDeleteFlag)", Contract
                On Error GoTo 0
            End If
        End If
    End If
    '10910
    If Not mSaveContractDocuments(tgChfCntr.lCntrNo) Then
        GoTo mSaveRecErr
    End If
    mSaveRecFunction = True
    'ilRet = btrEndTrans(hmCHF)
    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
        If imCntrRep Then
            If tgChfCntr.sNRProcessed = "N" Then
                tmGif.lCode = 0
                tmGif.sNRProcessed = "N"
                tmGif.sType = "C"
                tmGif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                tmGif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                Select Case Left$(gGetLocalTZName(), 1)
                    Case "E"
                    Case "C"
                        tmGif.lGenTime = tmGif.lGenTime + 3600
                    Case "M"
                        tmGif.lGenTime = tmGif.lGenTime + 7200
                    Case "P"
                        tmGif.lGenTime = tmGif.lGenTime + 10800
                End Select
                tmGif.lCntrNo = tgChfCntr.lCntrNo
                tmGif.iVefCode = 0
                tmGif.iStartDate(0) = tgChfCntr.iStartDate(0)
                tmGif.iStartDate(1) = tgChfCntr.iStartDate(1)
                tmGif.iEndDate(0) = tgChfCntr.iEndDate(0)
                tmGif.iEndDate(1) = tgChfCntr.iEndDate(1)
                tmGif.sMessageRepDBID = ""
                tmGif.sMessageNetDBID = ""
                tmGif.lMessageCefCode = 0
                gPackDate "", tmGif.iProcessedDate(0), tmGif.iProcessedDate(1)
                gPackTime "", tmGif.iProcessedTime(0), tmGif.iProcessedTime(1)
                ilRet = btrInsert(hmGif, tmGif, imGifRecLen, INDEXKEY0)
                On Error GoTo mSaveRecErr
                gCPErrorMsg ilRet, "mSaveRec GIF (btrInsert)", Contract
                On Error GoTo 0
            End If
        End If
    End If
    If ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        If Trim$(tgChfCntr.sRepDBID) <> "" Then
            If tgChfCntr.sStatus <> smOrigStatus Then
                If (smOrigStatus <> "N") And (smOrigStatus <> "G") And (smOrigStatus <> "O") And (smOrigStatus <> "H") And (smOrigStatus <> "D") And (smOrigStatus <> "I") Then
                    If (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "O") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "D") Or (tgChfCntr.sStatus = "I") Then
                        slStr = "C|O|" & sgNetDBID & "|" & tgChfCntr.lCntrNo & "|" & Trim$(tmAdf.sName) & "||"
                        slStr = slStr & "Contract "
                        If tgChfCntr.sStatus = "D" Then
                            slStr = slStr & "Rejected"
                        ElseIf (tgChfCntr.sStatus = "N") Then
                            slStr = slStr & "Approved Order"
                        ElseIf (tgChfCntr.sStatus = "G") Then
                            slStr = slStr & "Approved Hold"
                        ElseIf (tgChfCntr.sStatus = "I") Then
                            slStr = slStr & "Unapproved"
                        Else
                            slStr = slStr & "Scheduled"
                        End If
                        'tmCef.iStrLen = Len(Trim$(slStr))
                        tmCef.sComment = Trim$(slStr) & Chr$(0) '& Chr$(0) 'sgTB
                        'imCefRecLen = 5 + Len(Trim$(tmCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
                        imCefRecLen = Len(tmCef)
                        tmCef.lCode = 0 'Autoincrement
                        ilRet = btrInsert(hmCef, tmCef, imCefRecLen, INDEXKEY0)
                        tmGif.lCode = 0
                        tmGif.sNRProcessed = "N"
                        tmGif.sType = "M"
                        tmGif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
                        tmGif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
                        Select Case Left$(gGetLocalTZName(), 1)
                            Case "E"
                            Case "C"
                                tmGif.lGenTime = tmGif.lGenTime + 3600
                            Case "M"
                                tmGif.lGenTime = tmGif.lGenTime + 7200
                            Case "P"
                                tmGif.lGenTime = tmGif.lGenTime + 10800
                        End Select
                        tmGif.lCntrNo = tgChfCntr.lCntrNo
                        tmGif.iVefCode = 0
                        tmGif.iStartDate(0) = 0
                        tmGif.iStartDate(1) = 0
                        tmGif.iEndDate(0) = 0
                        tmGif.iEndDate(1) = 0
                        tmGif.sMessageRepDBID = tgChfCntr.sRepDBID
                        tmGif.sMessageNetDBID = sgNetDBID
                        tmGif.lMessageCefCode = tmCef.lCode
                        gPackDate "", tmGif.iProcessedDate(0), tmGif.iProcessedDate(1)
                        gPackTime "", tmGif.iProcessedTime(0), tmGif.iProcessedTime(1)
                        ilRet = btrInsert(hmGif, tmGif, imGifRecLen, INDEXKEY0)
                    End If
                End If
            End If
        End If
    End If
    'Clear alerts
    tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet - BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo)
        ilRet = gAlertClear("A", "C", "C", tmChf.lCode, 0, "")
        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    If tgChfCntr.sStatus = "C" Then
        If tgChfCntr.sSource <> "P" Then
            ilRet = gAlertAdd("C", "C", tgChfCntr.lCode, 0, "")
        Else
            ilRet = gAlertAdd("C", "P", tgChfCntr.lCode, 0, "")
        End If
    'ElseIf smOrigStatus = "C" Then
    '    ilRet = gAlertClear("A", "C", "C", tgChfCntr.lCode, 0, "")
    End If
    If (tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "N") Then
        If tgChfCntr.sSource <> "P" Then
            ilRet = gAlertAdd("C", "S", tgChfCntr.lCode, 0, "")
        Else
            ilRet = gAlertAdd("C", "P", tgChfCntr.lCode, 0, "")
        End If
    'ElseIf (smOrigStatus = "G") Or (smOrigStatus = "N") Then
    '    ilRet = gAlertClear("A", "C", "S", tgChfCntr.lCode, 0, "")
    End If
    '8/15/18: Add Unapproved alerts
    If tgChfCntr.sStatus = "I" Then
        ilRet = gAlertAdd("C", "I", tgChfCntr.lCode, 0, "")
    End If
    Contract.udcCBill.Action 4
    imInSave = False
    Screen.MousePointer = vbDefault    'Default
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    'imTerminate = True
    imInSave = False
    mSaveRecFunction = False
    'ilRet = btrAbortTrans(hmCHF)
    '8723
    If ogContractCreator.ContractToProcess = Order Then
        ogContractCreator.ProcessResult = ErrorResult
    End If
    Exit Function
End Function

Public Sub pbcLnTabSub(LINECOMMINDEX As Integer, PRICEINDEX As Integer, OV1INDEX As Integer, LENINDEX As Integer, OV2INDEX As Integer, VEHINDEX As Integer, DPINDEX As Integer, WKSPOTSINDEX As Integer)
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    'Dim llLnStartDate As Long
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim llAllowedChgWkStartDate As Long
    Dim ilSpots As Integer
    Dim ilVef As Integer
    '10347
    Dim slLineType As String
    
    If bmSettingPriceInTab Then
        Exit Sub
    End If
    If Contract.plcPower.Visible Then
        If Contract.plcPower.Enabled Then
            Contract.edcPBDate(0).SetFocus
        Else
            '2/7/09: Added to handle case where focus can't be set
            On Error Resume Next
            Contract.pbcClickFocus.SetFocus
            On Error GoTo 0
        End If
    End If
    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        '2/7/09: Added to handle case where focus can't be set
        On Error Resume Next
        Contract.pbcClickFocus.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    If GetFocus() <> Contract.pbcLnTab.HWnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    If imDirProcess >= 0 Then
        Contract.mLnDirection imDirProcess
        imDirProcess = -1
        Exit Sub
    End If
    Contract.mAllSetShow 2
    'If (imSelectedIndex > 0) And (tgChfCntr.sStatus = "I") Then
    If (tgChfCntr.sSchStatus = "I") Then
        ilRet = MsgBox("Contract partially scheduled, finish scheduling", vbOKOnly + vbExclamation, "Error")
        Contract.cmcHold.Enabled = True
        Contract.cmcHold.SetFocus
        Exit Sub
    End If
    If (imCreditRestr = 1) And ((imStep1Index = 0) Or (imStep1Index = 1)) Then
        ilRet = MsgBox("NO NEW ORDERS", vbOKOnly + vbExclamation, "Restriction")
        Contract.cmcCancel.SetFocus
        Exit Sub
    End If
    If imLnBoxNo = LINECOMMINDEX Then
        If Contract.mLnCommentBranch() Then
            Exit Sub
        End If
    End If
    If Contract.grdCPMBuy.Visible Then
        gCPMTab
        Exit Sub
    End If
    Select Case imLnBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            
            imLnRowNo = UBound(tgClfCntr) + 1
            
            '1/11/15: Add Hide/Show Hidden Lines
            imLnRowNo = mRemapRowNo(imLnRowNo)
            
            Contract.mCntrTotal 'Obtain total for new line
            imSettingValue = True
            If (imLnRowNo <= Contract.vbcLines.LargeChange + 1) Or (imLnRowNo - Contract.vbcLines.LargeChange - 1 < 0) Then
                Contract.vbcLines.Value = 0
            Else
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                Contract.vbcLines.Value = imLnRowNo - Contract.vbcLines.LargeChange - 1
                On Error GoTo 0
            End If
            mGetFlightDays imLnRowNo, imOvDays()
            imSettingValue = False
            ilBox = 1
        '10/20/15: Changed from Comment to Price becaus ethe expansion of package lines moved from Comment to Price
        Case -LINECOMMINDEX, -PRICEINDEX 'From flight screen
            smLineType = mGetLineType(imLnRowNo)
            If (smLineType = "S") Then
                ilBox = -imLnBoxNo   'LINECOMMINDEX
                gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) > 0 Then
                    slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCode = CInt(slCode)
                    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilLoop).iCode = ilCode Then
                        ilLoop = gBinarySearchVef(ilCode)
                        If ilLoop <> -1 Then
                            If (tgMVef(ilLoop).sType = "P") And (tgMVef(ilLoop).lPvfCode > 0) Then
                                ilBox = PRICEINDEX
                            End If
                    '        Exit For
                        End If
                    'Next ilLoop
                End If
            Else
                ilBox = -imLnBoxNo   'LINECOMMINDEX
            End If
            If (ilBox = PRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                ilBox = LINECOMMINDEX
            End If
            imLnBoxNo = ilBox
            imMoveLnRowNo = imLnRowNo
            Contract.mLnEnableBox ilBox
            Screen.MousePointer = vbDefault
            Exit Sub
        Case OV1INDEX
            If Contract.edcAct1.Visible Then
                imOVBoxNo = ACT1LINEUPINDEX
                ilBox = ACT1LINEUPINDEX
                Contract.mOvEnableBox ilBox
                Contract.edcAct1.Visible = False
                Exit Sub
            End If
            Select Case imOVBoxNo
                Case -1
                    If smAlterHidden = "C" Then
                        imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                        ilBox = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                    Else
                        If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                imOVBoxNo = OVBBOPENINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                ilBox = OVBBOPENINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            Else
                                'imOVBoxNo = OVBBCLOSEINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                'ilBox = OVBBCLOSEINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                If bmAct1PackageEnable Then
                                    If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                                        imOVBoxNo = ACT1LINEUPINDEX
                                        ilBox = ACT1LINEUPINDEX
                                    Else
                                        imOVBoxNo = ACT1DAYPARTFILTERINDEX
                                        ilBox = ACT1DAYPARTFILTERINDEX
                                    End If
                                Else
                                    imOVBoxNo = OVCOPYINDEX
                                    ilBox = OVCOPYINDEX
                                End If
                            End If
                        Else
                            'imOVBoxNo = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            'ilBox = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            If bmAct1PackageEnable Then
                                If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                                    imOVBoxNo = ACT1LINEUPINDEX
                                    ilBox = ACT1LINEUPINDEX
                                Else
                                    imOVBoxNo = ACT1DAYPARTFILTERINDEX
                                    ilBox = ACT1DAYPARTFILTERINDEX
                                End If
                            Else
                                imOVBoxNo = OVCOPYINDEX
                                ilBox = OVCOPYINDEX
                            End If
                        End If
                    End If
                    If (Asc(tgSpf.sUsingFeatures8) And LRMANDATORY) = LRMANDATORY Then
                        If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                            imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            ilBox = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                        End If
                    End If
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                'Case OVPREEMPTINDEX
                Case OVDAYSINDEX + 6
                    If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                        Contract.mOvSetShow imOVBoxNo
                        'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                        '    ilBox = OVCOPYINDEX
                        '    imOVBoxNo = ilBox
                        '    Contract.mOvEnableBox ilBox
                        '    Exit Sub
                        'Else
                            If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                                ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") Then
                                    'Test if daily spots defined
                                    ilSpots = 0
                                    For ilLoop = 0 To 6 Step 1
                                        ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                                    Next ilLoop
                                    If ilSpots = 0 Then
                                        Beep
                                        imOVBoxNo = OVDAYSINDEX
                                        Contract.mOvEnableBox imOVBoxNo
                                        Exit Sub
                                    End If
                                Else
                                    'Weekly buy
                                End If
                                If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                                    ilBox = OVBBOPENINDEX 'OVPREEMPTINDEX
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                End If
                                'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                                    ilBox = OVCOPYINDEX
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                'End If
                                'imOVBoxNo = -1
                                'Contract.mSetTmOV2Ctrl
                                'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                                '    ilBox = OV2INDEX    'LENINDEX
                                'Else
                                '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                '    ilBox = ilFound
                                '    imLnBoxNo = ilBox
                                '    Contract.mLnEnableBox ilBox
                                '    Screen.MousePointer = vbDefault
                                '    Exit Sub
                                'End If
                            End If
                        'End If
                    Else
                        ilBox = imOVBoxNo + 1
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                Case OVWDINDEX
                    ilBox = imOVBoxNo + 2
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OVHIGHINDEX
                    If (tgSpf.sCPkEqual <> "Y") Then
                        Contract.mOvSetShow imOVBoxNo
                        If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                            'Test if daily spots defined
                            ilSpots = 0
                            For ilLoop = 0 To 6 Step 1
                                ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                            Next ilLoop
                            If ilSpots = 0 Then
                                Beep
                                imOVBoxNo = OVDAYSINDEX
                                Contract.mOvEnableBox imOVBoxNo
                                Exit Sub
                            End If
                        Else
                        End If
                        If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                            ilBox = OVBBOPENINDEX 'OVPREEMPTINDEX
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        End If
                        ilBox = OVCOPYINDEX
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                        
                        'imOVBoxNo = -1
                        'Contract.mSetTmOV2Ctrl
                        'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                        '    ilBox = OV2INDEX    'LENINDEX
                        'Else
                        '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                        '    ilBox = ilFound
                        '    imLnBoxNo = ilBox
                        '    Contract.mLnEnableBox ilBox
                        '    Screen.MousePointer = vbDefault
                        '    Exit Sub
                        'End If
                    Else
                        ilBox = OVRANKINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                Case OVRANKINDEX
                    Contract.mOvSetShow imOVBoxNo
                    If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                        'Test if daily spots defined
                        ilSpots = 0
                        For ilLoop = 0 To 6 Step 1
                            ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                        Next ilLoop
                        If ilSpots = 0 Then
                            Beep
                            imOVBoxNo = OVDAYSINDEX
                            Contract.mOvEnableBox imOVBoxNo
                            Exit Sub
                        End If
                    Else
                        'Weekly buy
                    End If
                    If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                        ilBox = OVBBOPENINDEX 'OVPREEMPTINDEX
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                    ilBox = OVCOPYINDEX
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                    
                    'imOVBoxNo = -1
                    'Contract.mSetTmOV2Ctrl
                    'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                    '    ilBox = OV2INDEX    'LENINDEX
                    'Else
                    '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                    '    ilBox = ilFound
                    '    imLnBoxNo = ilBox
                    '    Contract.mLnEnableBox ilBox
                    '    Screen.MousePointer = vbDefault
                    '    Exit Sub
                    'End If
                Case OVSTARTTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                        ilBox = OVENDTIMEINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        smDGSave(14) = ""
                        Contract.mOvSetShow imOVBoxNo
                        'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                        '    ilBox = OVCOPYINDEX
                        '    imOVBoxNo = ilBox
                        '    Contract.mOvEnableBox ilBox
                        '    Exit Sub
                        'Else
                            If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                                ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                If (tgSpf.sCPkEqual = "Y") Then
                                    ilBox = OVRANKINDEX
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                End If
                                If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") Then
                                    'Test if daily spots defined
                                    ilSpots = 0
                                    For ilLoop = 0 To 6 Step 1
                                        ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                                    Next ilLoop
                                    If ilSpots = 0 Then
                                        Beep
                                        imOVBoxNo = OVDAYSINDEX
                                        Contract.mOvEnableBox imOVBoxNo
                                        Exit Sub
                                    End If
                                Else
                                    'Weekly buy
                                End If
                                If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                                    ilBox = OVBBOPENINDEX 'OVPREEMPTINDEX
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                End If
                                ilBox = OVCOPYINDEX
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                                'imOVBoxNo = -1
                                'Contract.mSetTmOV2Ctrl
                                'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                                '    ilBox = OV2INDEX    'LENINDEX
                                'Else
                                '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                '    ilBox = ilFound
                                '    imLnBoxNo = ilBox
                                '    Contract.mLnEnableBox ilBox
                                '    Screen.MousePointer = vbDefault
                                '    Exit Sub
                                'End If
                            End If
                        'End If
                    End If
                Case OVENDTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                    Contract.mOvSetShow imOVBoxNo
                    'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                    '    ilBox = OVCOPYINDEX
                    '    imOVBoxNo = ilBox
                    '    Contract.mOvEnableBox ilBox
                    '    Exit Sub
                    'Else
                        If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                            ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If (tgSpf.sCPkEqual = "Y") Then
                                ilBox = OVRANKINDEX
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            End If
                            Contract.mOvSetShow imOVBoxNo
                            If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") Then
                                'Test if daily spots defined
                                ilSpots = 0
                                For ilLoop = 0 To 6 Step 1
                                    ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                                Next ilLoop
                                If ilSpots = 0 Then
                                    Beep
                                    imOVBoxNo = OVDAYSINDEX
                                    Contract.mOvEnableBox imOVBoxNo
                                    Exit Sub
                                End If
                            Else
                                'Weekly buy
                            End If
                            If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                                ilBox = OVBBOPENINDEX 'OVPREEMPTINDEX
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            End If
                            ilBox = OVCOPYINDEX
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                            'imOVBoxNo = -1
                            'Contract.mSetTmOV2Ctrl
                            'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                            '    ilBox = OV2INDEX    'LENINDEX
                            'Else
                            '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                            '    ilBox = ilFound
                            '    imLnBoxNo = ilBox
                            '    Contract.mLnEnableBox ilBox
                            '    Screen.MousePointer = vbDefault
                            '    Exit Sub
                            'End If
                        End If
                    'End If
                Case OVBBOPENINDEX
                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                        Contract.mOvSetShow imOVBoxNo
                        If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                            'Test if daily spots defined
                            ilSpots = 0
                            For ilLoop = 0 To 6 Step 1
                                ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                            Next ilLoop
                            If ilSpots = 0 Then
                                Beep
                                imOVBoxNo = OVDAYSINDEX
                                Contract.mOvEnableBox imOVBoxNo
                                Exit Sub
                            End If
                        Else
                            'Weekly buy
                        End If
                        ilBox = OVCOPYINDEX
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                        'imOVBoxNo = -1
                        'Contract.mSetTmOV2Ctrl
                        'If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                        '    ilBox = OV2INDEX    'LENINDEX
                        'Else
                        '    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                        '    ilBox = ilFound
                        '    imLnBoxNo = ilBox
                        '    Contract.mLnEnableBox ilBox
                        '    Screen.MousePointer = vbDefault
                        '    Exit Sub
                        'End If
                    Else
                        ilBox = imOVBoxNo + 1
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                Case OVCOPYINDEX    'OVBBCLOSEINDEX
                    If tmOV1Ctrls(OVCOPYINDEX).iReq = True Then
                        If imDGSave(7) < 0 Then
                            Beep
                            imOVBoxNo = OVCOPYINDEX
                            Contract.mOvEnableBox imOVBoxNo
                            Exit Sub
                        End If
                    End If
                    Contract.mOvSetShow imOVBoxNo
                    If smAlterHidden <> "C" Then
                        If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                            'Test if daily spots defined
                            ilSpots = 0
                            For ilLoop = 0 To 6 Step 1
                                ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                            Next ilLoop
                            If ilSpots = 0 Then
                                Beep
                                imOVBoxNo = OVDAYSINDEX
                                Contract.mOvEnableBox imOVBoxNo
                                Exit Sub
                            End If
                        Else
                            'Weekly buy
                        End If
                        If bmAct1PackageEnable Then
                            imOVBoxNo = ACT1LINEUPINDEX
                            ilBox = ACT1LINEUPINDEX
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        End If
                        imOVBoxNo = -1
                        Contract.mSetTmOV2Ctrl
                        If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                            ilBox = OV2INDEX    'LENINDEX
                        Else
                            Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                            ilBox = ilFound
                            imLnBoxNo = ilBox
                            Contract.mLnEnableBox ilBox
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    Else
                        'imOVBoxNo = -1
                        'Contract.mSetTmOV2Ctrl
                        'Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                        Contract.mLnSetShow imLnBoxNo
                        ilBox = LINECOMMINDEX
                        imLnBoxNo = LINECOMMINDEX
                        Contract.mLnEnableBox ilBox
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                Case ACT1LINEUPINDEX
                    Contract.mOvSetShow imOVBoxNo
                    If Trim$(smDGSave(15)) <> "" Then
                        imOVBoxNo = ACT1STOREDTIMESINDEX
                        ilBox = ACT1STOREDTIMESINDEX
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                            'Test if daily spots defined
                            ilSpots = 0
                            For ilLoop = 0 To 6 Step 1
                                ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                            Next ilLoop
                            If ilSpots = 0 Then
                                Beep
                                imOVBoxNo = OVDAYSINDEX
                                Contract.mOvEnableBox imOVBoxNo
                                Exit Sub
                            End If
                        Else
                            'Weekly buy
                        End If
                        imOVBoxNo = -1
                        Contract.mSetTmOV2Ctrl
                        If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                            ilBox = OV2INDEX    'LENINDEX
                        Else
                            Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                            ilBox = ilFound
                            imLnBoxNo = ilBox
                            Contract.mLnEnableBox ilBox
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                Case ACT1DAYPARTFILTERINDEX
                    Contract.mOvSetShow imOVBoxNo
                    If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") And (Not mIsLnGame(imLnRowNo)) Then
                        'Test if daily spots defined
                        ilSpots = 0
                        For ilLoop = 0 To 6 Step 1
                            ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                        Next ilLoop
                        If ilSpots = 0 Then
                            Beep
                            imOVBoxNo = OVDAYSINDEX
                            Contract.mOvEnableBox imOVBoxNo
                            Exit Sub
                        End If
                    Else
                        'Weekly buy
                    End If
                    imOVBoxNo = -1
                    Contract.mSetTmOV2Ctrl
                    If (imShowOV2 >= 0) And (smLineTypeOv <> "P") Then
                        ilBox = OV2INDEX    'LENINDEX
                    Else
                        Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                        ilBox = ilFound
                        imLnBoxNo = ilBox
                        Contract.mLnEnableBox ilBox
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                Case Else
                    ilBox = imOVBoxNo + 1
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
            End Select
        Case OV2INDEX
            If imOVBoxNo = OV2SPLITINDEX Then
                If Contract.mSplitBranch() Then
                    Exit Sub
                End If
            End If
            Select Case imOVBoxNo
                Case -1
                    ilBox = imShowOV2
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OV2STARTTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                        ilBox = OV2ENDTIMEINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        smLnSave(26, imLnRowNo) = ""
                        slStr = ""
                        gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2ENDTIMEINDEX)
                        If mRegionAllowed(imLnRowNo) Then  'Allocation
                            ilBox = OV2SPLITINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                                ilBox = OV2ALLOCATIONINDEX
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                                If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                                    ilBox = OV2ACQUISITIONINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                Else
                                    If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                                        ilBox = OV2SOLOINDEX
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = ilBox
                                        Contract.mOvEnableBox ilBox
                                        Exit Sub
                                    Else
                                        If (Asc(tgSpf.sOverrideOptions) And SP1STPOSITION) = SP1STPOSITION Then
                                            ilBox = OV21STINDEX
                                            Contract.mOvSetShow imOVBoxNo
                                            imOVBoxNo = ilBox
                                            Contract.mOvEnableBox ilBox
                                            Exit Sub
                                        Else
                                            Contract.mOvSetShow imOVBoxNo
                                            imOVBoxNo = -1
                                            Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                            ilBox = ilFound
                                            imLnBoxNo = ilBox
                                            Contract.mLnEnableBox ilBox
                                            Screen.MousePointer = vbDefault
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case OV2ENDTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                    Else
                        smLnSave(25, imLnRowNo) = ""
                        slStr = ""
                        gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2STARTTIMEINDEX)
                    End If
                    If mRegionAllowed(imLnRowNo) Then  'Allocation
                        ilBox = OV2SPLITINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                            ilBox = OV2ALLOCATIONINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                                ilBox = OV2ACQUISITIONINDEX
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                                    ilBox = OV2SOLOINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                Else
                                    If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                                        ilBox = OV21STINDEX
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = ilBox
                                        Contract.mOvEnableBox ilBox
                                        Exit Sub
                                    Else
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = -1
                                        Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                        ilBox = ilFound
                                        imLnBoxNo = ilBox
                                        Contract.mLnEnableBox ilBox
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case OV2SPLITINDEX
                    If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                        ilBox = OV2ALLOCATIONINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                        If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                            ilBox = OV2ACQUISITIONINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                                ilBox = OV2SOLOINDEX
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                                    ilBox = OV21STINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                Else
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = -1
                                    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                    ilBox = ilFound
                                    imLnBoxNo = ilBox
                                    Contract.mLnEnableBox ilBox
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Case OV2ALLOCATIONINDEX
                    'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                    If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                        ilBox = OV2ACQUISITIONINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                            ilBox = OV2SOLOINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If (Asc(tgSpf.sOverrideOptions) And SP1STPOSITION) = SP1STPOSITION Then
                                ilBox = OV21STINDEX
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                                ilBox = ilFound
                                imLnBoxNo = ilBox
                                Contract.mLnEnableBox ilBox
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                        End If
                    End If
                Case OV2ACQUISITIONINDEX
                    If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                        ilBox = OV2SOLOINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If (Asc(tgSpf.sOverrideOptions) And SP1STPOSITION) = SP1STPOSITION Then
                            ilBox = OV21STINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                            ilBox = ilFound
                            imLnBoxNo = ilBox
                            Contract.mLnEnableBox ilBox
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                Case OV2SOLOINDEX
                    If (Asc(tgSpf.sOverrideOptions) And SP1STPOSITION) = SP1STPOSITION Then
                        ilBox = OV21STINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = -1
                        Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                        ilBox = ilFound
                        imLnBoxNo = ilBox
                        Contract.mLnEnableBox ilBox
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                Case OV21STINDEX
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = -1
                    Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
                    ilBox = ilFound
                    imLnBoxNo = ilBox
                    Contract.mLnEnableBox ilBox
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Case Else
                    ilBox = imOVBoxNo + 1
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
            End Select
        Case 0  'At arrow
            If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                ilBox = DPINDEX
            Else
                ilBox = VEHINDEX
            End If
        Case VEHINDEX
            If (imLnRowNo >= UBound(tgClfCntr) + 1) And (Contract.lbcLnVehicle(igTabMapIndex).ListIndex < 0) Then
                Contract.mLnSetShow imLnBoxNo
                For ilLoop = VEHINDEX To LINECOMMINDEX Step 1
                    slStr = ""
                    gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilLoop)
                    smLnShow(ilLoop, imLnRowNo) = tmLnCtrls(ilLoop).sShow
                Next ilLoop
                tgClfCntr(UBound(tgClfCntr)).ClfRec.iLine = 0
                smLnSave(1, imLnRowNo) = ""
                imLnBoxNo = -1
                imLnRowNo = -1
                Contract.mCntrTotal 'Obtain total for new line
                Contract.pbcLine_Paint
                If Contract.cmcUpdate.Enabled Then
                    'Contract.cmcUpdate.SetFocus
                    Contract.cmcSaveStay.SetFocus
                Else
                    Contract.cmcDone.SetFocus
                End If
                Exit Sub
            End If
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                If Contract.lbcLnVehicle(igTabMapIndex).ListIndex < 0 Then
                    Exit Sub
                End If
            Else
                If Trim$(Contract.edcLNDropDown.Text) = "" Then
                    Exit Sub
                End If
            End If
            '9/11/14: If installment defined, disallow sports vehicle
            If ((Asc(tgSpf.sSportInfo) And USINGSPORTS) = USINGSPORTS) And ((Asc(tgSaf(0).sFeatures2) And EVENTREVENUE) = EVENTREVENUE) Then
                If ((Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT) And (UBound(tgFBSbf) > LBound(tgFBSbf)) Then
                    If Contract.lbcLnVehicle(igTabMapIndex).ListIndex >= 0 Then
                        slNameCode = tmVehicleCode(Contract.lbcLnVehicle(igTabMapIndex).ListIndex).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVef = gBinarySearchVef(CInt(slCode))
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType = "G" Then
                                MsgBox "Sports and Installment disallowed when using Revenue By Events on the same contract", vbOKOnly + vbExclamation, "Warning"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
            ilBox = imLnBoxNo + 1
        Case DPINDEX
            Contract.mLnSetShow imLnBoxNo
            'If Contract.tscLine.SelectedItem.Index = 10 Then
            If Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH) Then
                ilBox = LINECOMMINDEX    'OV1INDEX
                imLnBoxNo = ilBox
                Contract.mLnEnableBox ilBox
                Exit Sub
            End If
            ilFound = False
            smLineType = mGetLineType(imLnRowNo)
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                For ilLoop = 0 To Contract.lbcPrg(igTabMapIndex).ListCount - 1 Step 1
                    If imLnSave(1, imLnRowNo) = Contract.lbcPrg(igTabMapIndex).ItemData(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
            Else
                If smLnSave(20, imLnRowNo) <> "" Then
                    ilFound = True
                End If
            End If
            If Not ilFound Then
                Exit Sub
            End If
            ilBox = LENINDEX    'OV1INDEX
            imLnBoxNo = ilBox
            Contract.mLnEnableBox ilBox
            Exit Sub
        Case WKSPOTSINDEX + 12

            If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 12) And (imOVType = 3) And (Contract.pbcOVWD(0).Visible) And (imOVBoxNo >= imLBCTCtrls) And (imOVBoxNo <= UBound(tmDailyCtrls)) Then
                Select Case imOVBoxNo
                    Case OV2DAYSINDEX + 6
                        Contract.mOvSetShow imOVBoxNo
                    Case Else
                        ilBox = imOVBoxNo + 1
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                End Select
            End If
            'slStr = UCase$(Left$(Contract.edcLNDropDown.Text, 1))
            slCode = Trim$(Contract.edcLNDropDown.Text)
            ilIndex = InStr(1, slCode, "P", 1)
            If ilIndex = 0 Then
                ilIndex = InStr(1, slCode, "Q", 1)
            End If
            If ilIndex = 0 Then
                ilIndex = InStr(1, slCode, "C", 1)
            End If
            If ilIndex > 0 Then
                'If ilIndex > 1 Then
                '10/18/11: Bypass call to Contract.mLnSetShow to aviod week being set to blank
                If (ilIndex > 1) And (Not mIsLnGame(imLnRowNo)) Then
                    Contract.edcLNDropDown.Text = Left$(slCode, ilIndex - 1)
                    If Trim$(Contract.edcLNDropDown.Text) <> "" Then
                        If (imOVType = 3) And (Contract.pbcOVWD(0).Visible) And (imOVBoxNo >= imLBCTCtrls) And (imOVBoxNo <= UBound(tmDailyCtrls)) Then
                            Contract.mOvSetShow imOVBoxNo
                        End If
                        Contract.mLnSetShow imLnBoxNo
                    End If
                End If
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
                smLineType = mGetLineType(imLnRowNo)
                If (smLineType <> "O") And (smLineType <> "A") Then
                    If smLineType <> "H" Then
                        ilBox = PRICEINDEX
                    Else
                        slStr = mGetPkLineType(imLnRowNo)
                        If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                            ilBox = LINECOMMINDEX
                        Else
                            ilBox = PRICEINDEX
                        End If
                    End If
                Else
                    If (tgSpf.sCAudPkg = "Y") Then
                        ilBox = PRICEINDEX
                    Else
                        ilBox = LINECOMMINDEX
                    End If
                End If
                If Contract.plcDemoBar.Visible Then
                    imDBBoxNo = -1
                End If
                '10843 remove dead code
'                If (ilBox = PRICEINDEX) And (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                    mSetPriceInTab PRICEINDEX
'                    ilBox = LINECOMMINDEX
'                End If
            End If
            If (ilBox = PRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                mSetPriceInTab PRICEINDEX
                ilBox = LINECOMMINDEX
            End If
            'slStr = UCase$(Left$(Contract.edcLNDropDown.Text, 1))
            If InStr(1, slCode, "P", 1) > 0 Then
                imLnBoxNo = ilBox
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf (InStr(1, slCode, "Q", 1) > 0) Or (InStr(1, slCode, "C", 1) > 0) Then
                ilBox = LINECOMMINDEX
                imLnBoxNo = ilBox
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                Contract.mLnSetShow imLnBoxNo
                ilBox = WKSPOTSINDEX
                Contract.pbcLine_MouseUp 0, 0, CSng(lmLnHotSpot(3, 1)), CSng(lmLnHotSpot(3, 2))
                imLnBoxNo = ilBox
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        '7/29/15: Swapped Length and override columns
        'Case LENINDEX, WKSPOTSINDEX To WKSPOTSINDEX + 11    '12
        Case WKSPOTSINDEX To WKSPOTSINDEX + 11
            ''If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 12) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (igFreezeState = 1) And (Contract.pbcOvWD(0).Visible) And (imOVBoxNo >= LBound(tmDailyCtrls)) And (imOVBoxNo <= UBound(tmDailyCtrls)) Then
            'If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 12) And (imLnSave(8, imLnRowNo) = 0) And (imOVType = 3) And (Contract.pbcOvWD(0).Visible) And (imOVBoxNo >= LBound(tmDailyCtrls)) And (imOVBoxNo <= UBound(tmDailyCtrls)) Then
            If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 12) And (imOVType = 3) And (Contract.pbcOVWD(0).Visible) And (imOVBoxNo >= imLBCTCtrls) And (imOVBoxNo <= UBound(tmDailyCtrls)) Then
                Select Case imOVBoxNo
                    Case OV2DAYSINDEX + 6
                        Contract.mOvSetShow imOVBoxNo
                    Case Else
                        ilBox = imOVBoxNo + 1
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                End Select
            End If
            slCode = Trim$(Contract.edcLNDropDown.Text)
            If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 11) And (InStr(1, slCode, "P", 1) > 0) Then
                ilIndex = InStr(1, slCode, "P", 1)
                '10/18/11: Bypass call to Contract.mLnSetShow to aviod week being set to blank
                'If ilIndex >= 1 Then
                If (ilIndex >= 1) And (Not mIsLnGame(imLnRowNo)) Then
                    If imDGSave(2) = 0 Then
                        Contract.mLnSetShow imLnBoxNo
                    Else
                        Contract.edcLNDropDown.Text = Left$(slCode, ilIndex - 1)
                        If (Trim$(Contract.edcLNDropDown.Text) <> "") Or (lmWkSpots <> 0) Then
                            Contract.mLnSetShow imLnBoxNo
                        End If
                    End If
                End If
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
                smLineType = mGetLineType(imLnRowNo)
                If (smLineType <> "O") And (smLineType <> "A") Then
                    If smLineType <> "H" Then
                        ilFound = PRICEINDEX
                    Else
                        slStr = mGetPkLineType(imLnRowNo)
                        If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                            ilFound = LINECOMMINDEX
                        Else
                            ilFound = PRICEINDEX
                        End If
                    End If
                Else
                    If (tgSpf.sCAudPkg = "Y") Then
                        ilFound = PRICEINDEX
                    Else
                        ilFound = LINECOMMINDEX
                    End If
                End If
            ElseIf (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 11) And ((InStr(1, slCode, "Q", 1) > 0) Or (InStr(1, slCode, "C", 1) > 0)) Then
                ilIndex = InStr(1, slCode, "Q", 1)
                If ilIndex = 0 Then
                    ilIndex = InStr(1, slCode, "C", 1)
                End If
                If ilIndex >= 1 Then
                    If imDGSave(2) = 0 Then
                        Contract.mLnSetShow imLnBoxNo
                    Else
                        Contract.edcLNDropDown.Text = Left$(slCode, ilIndex - 1)
                        If (Trim$(Contract.edcLNDropDown.Text) <> "") Or (lmWkSpots <> 0) Then
                            Contract.mLnSetShow imLnBoxNo
                        End If
                    End If
                End If
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
                ilFound = LINECOMMINDEX
            Else
                Contract.mNextAllowedWeek ilFound, ilBox, llAllowedChgWkStartDate, llDate
            End If
            'Contract.mCntrTotal 'Obtain total for new line
            ilBox = ilFound
           '10843 remove dead code
'            If (ilBox = PRICEINDEX) And (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                mSetPriceInTab PRICEINDEX
'                ilBox = LINECOMMINDEX
'            End If
            If (ilBox = PRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                mSetPriceInTab PRICEINDEX
                ilBox = LINECOMMINDEX
            End If
            imLnBoxNo = ilBox
            Contract.mLnEnableBox ilBox
            bmSettingPriceInTab = False
            Screen.MousePointer = vbDefault
            Exit Sub
        Case LINECOMMINDEX  'PRICEINDEX, LINECOMMINDEX 'Last control within header
            'smLineType = mGetLineType(imLnRowNo)
            'If ((imLnBoxNo = LINECOMMINDEX) And (smLineType = "O") Or (smLineType = "A")) Or (imLnBoxNo = PRICEINDEX) Then
                    Contract.mLnSetShow imLnBoxNo
                If Contract.mLnTestSaveFields(True) = NO Then
                    Contract.mLnEnableBox imLnBoxNo
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                imLnRowNo = mRemapRowNo(imMoveLnRowNo)
                If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                    If tgClfCntr(imLnRowNo - 1).iStatus = -1 Then
                        tgClfCntr(imLnRowNo - 1).iStatus = 0
                        imLnChg = True
                        imComputeReq = True
                    End If
                    ReDim Preserve tgClfCntr(0 To imLnRowNo) As CLFLIST
                    ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
                    ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
                    ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
                    ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
                    ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
                    ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                    ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                    For ilLoop = LBound(imLnSave, FIELDBOUNDINDEX) To UBound(imLnSave, FIELDBOUNDINDEX) Step 1
                        imLnSave(ilLoop, UBound(tgClfCntr) + 1) = -1
                    Next ilLoop
                    'mSetHdDate False, False
                End If
                '12/16/14: Add Hide/Show Hidden Lines
                'Do While (imLnRowNo < UBound(tgClfCntr)) And (tgClfCntr(imLnRowNo).iCancel)
                Do While (imLnRowNo < UBound(tgClfCntr)) And ((tgClfCntr(imLnRowNo).iCancel) Or (tgClfCntr(imLnRowNo).bPkExpanded = False) Or (tgClfCntr(imLnRowNo).bShowLine = False))
                    imLnRowNo = imLnRowNo + 1
                    mBuildMapLinePaint
                    imMoveLnRowNo = mFindPaintLnRowNo(imLnRowNo)
                Loop
                
                If (imLnRowNo >= UBound(tgClfCntr)) Then
                    'Disallow new lines in on Research tab since columns replaced
                    'If (Contract.tscLine.SelectedItem.Index = 10) Then
                    If (Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH)) Then
                        'Vehicle->smLnSave(1, imLnRowNo); Non-Package Daypart->imLnSave(1, imLnRowNo); Package Daypart->mLnSave(20, imLnRowNo); Len->smLnSave(16, imLnRowNo)
                        Contract.pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    Contract.mInitNewLine
                    smLnSave(1, imLnRowNo + 1) = "" 'Vehicle list index
                    '10865 didn't change first variable to true to save time...should be ok
                    Contract.mSetLineScroll False, False
                    Contract.pbcLine_Paint
                End If
                imLnRowNo = imLnRowNo + 1
                
                mBuildMapLinePaint
                imMoveLnRowNo = mFindPaintLnRowNo(imLnRowNo)
                
                Contract.mCntrTotal 'Obtain total for new line, placed back into service 7/15/97
                mSetSchStatus smOrigStatus
                '12/16/14: A dd Hide/Show Hidden Lines
                'If imLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                If imMoveLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                    '9/14/16: Handle case where all lines are CBS
                    If Contract.vbcLines.Value + 1 > Contract.vbcLines.Max Then
                        Contract.pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    imSettingValue = True
                    Contract.vbcLines.Value = Contract.vbcLines.Value + 1
                    imSettingValue = False
                End If
                If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                    If Contract.plcDemoBar.Visible Then
                        If imDBTotals < 5 Then
                            Contract.pbcDemoBar(0).Cls
                            Contract.pbcDemoBar(1).Cls
                            Contract.pbcDemoBar_Paint 0
                            Contract.pbcDemoBar_Paint 1
                        Else
                            Contract.pbcDemoBar(1).Cls
                            Contract.pbcDemoBar_Paint 1
                        End If
                    End If
                    Screen.MousePointer = vbDefault
                    Contract.mSetCommands
                    imLnBoxNo = 0 '1
                    Contract.mSetArrowFrame
                    If Contract.pbcLineArrow.Visible = True Then
                        Contract.pbcLineArrow.SetFocus
                    End If
                    Exit Sub
                Else
                    If imLnBoxNo = LINECOMMINDEX Then
                        If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                            ilBox = VEHINDEX
                        Else
                            '10347 added R and C testing
                            slLineType = mGetLineType(imLnRowNo)
                            If slLineType = "H" And smAlterHidden = "R" Then
                                ilBox = PRICEINDEX
                            ElseIf slLineType = "H" And smAlterHidden = "C" Then
                                ilBox = LINECOMMINDEX
                            Else
                                ilBox = WKSPOTSINDEX
                                '12/17/14: If week disallowed, set focus to comment field
                                If smLnSave(14, imLnRowNo) <> "" Then
                                    'llAllowedChgWkStartDate = mLnDate(ilRowNo, smLnLLD, False)
                                    llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 1)
                                    If llAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
                                        llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
                                    End If
                                    llDate = lmHSStartDate + 7 * (imHSCurWkNo + ilBox - WKSPOTSINDEX - 1)
                                    ''If (llAllowedChgWkStartDate > llDate) Or (llDate < gDateValue(gObtainPrevMonday(smLnSave(14, ilRowNo)))) Or ((llDate > gDateValue(gObtainNextSunday(smLnSave(15, ilRowNo))))) Then
                                    'If (llAllowedChgWkStartDate > llDate) Then
                                    If (llAllowedChgWkStartDate > llDate) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "A") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "N") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M")) Then
                                        ilBox = LINECOMMINDEX
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If tgClfCntr(imLnRowNo - 1).iInPast Then
                            ilBox = DPINDEX
                        Else
                            ilBox = VEHINDEX
                        End If
                    End If
                End If
                mGetFlightDays imLnRowNo, imOvDays()
                imLnBoxNo = ilBox
                If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                    imComputeReq = True
                    'cmcCompute.Enabled = True
                    If imDBTotals < 5 Then
                        Contract.pbcDemoBar(0).Cls
                        Contract.pbcDemoBar(1).Cls
                        Contract.pbcDemoBar_Paint 0
                        Contract.pbcDemoBar_Paint 1
                    Else
                        Contract.pbcDemoBar(1).Cls
                        Contract.pbcDemoBar_Paint 1
                    End If
                Else
                    If Contract.plcDemoBar.Visible Then
                        imComputeReq = True
                        'cmcCompute.Enabled = True
                        If imDBTotals < 5 Then
                            mSetDemoBarBooks imLnRowNo, ilFound, imCppCpm, 2, imCntrVefCode, imVpfIndex, imTerminate, imChgMode 'Using ilfound for ilDnfCode (mBuildWkInfo sets tgClf to DnfCode)
                            Contract.mBuildWkInfo True, True, False
                            Contract.pbcDemoBar(0).Cls
                            Contract.pbcDemoBar(1).Cls
                            Contract.pbcDemoBar_Paint 0
                            Contract.pbcDemoBar_Paint 1
                        Else
                            Contract.mBuildWkInfo True, True, False
                            Contract.pbcDemoBar(1).Cls
                            Contract.pbcDemoBar_Paint 1
                        End If
                        DoEvents
                    End If
                End If
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            'Else
            '    ilBox = imLnBoxNo + 1
            'End If
        Case Else
            ilBox = imLnBoxNo + 1
    End Select
    Contract.mLnSetShow imLnBoxNo
    imLnBoxNo = ilBox
    Contract.mLnEnableBox ilBox
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Sub pbcSTabSub(LINECOMMINDEX As Integer, PRICEINDEX As Integer, OV1INDEX As Integer, LENINDEX As Integer, OV2INDEX As Integer, VEHINDEX As Integer, DPINDEX As Integer, WKSPOTSINDEX As Integer)
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slCode As String
    Dim ilFound As Integer
    'Dim llLnStartDate As Long
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim llAllowedChgWkStartDate As Long
    Dim ilLastBox As Integer
    Dim ilSpots As Integer

    If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
        If imLnBoxNo <> -1 Then
            Contract.pbcClickFocus.SetFocus
        End If
        Exit Sub
    End If
    If GetFocus() <> Contract.pbcLnSTab.HWnd Then
        Exit Sub
    End If
    If Contract.grdCPMBuy.Visible Then
        gCPMSTab
        Exit Sub
    End If
    'If from Rate bar, then process override
    If (Contract.plcPower.Visible = True) And (Trim$(smPowerCaption) = "Rate Bar") Then
        Select Case imOVBoxNo
            Case -1
                If smAlterHidden = "C" Then
                    imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                    ilBox = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                Else
                    imOVBoxNo = OVDAYSINDEX
                    ilBox = OVDAYSINDEX
                End If
                Contract.mOvEnableBox ilBox
                Exit Sub
            Case OVDAYSINDEX
                Contract.mOvSetShow imOVBoxNo
                imOVBoxNo = -1
                Contract.plcOv.Visible = False
                Contract.pbcOv.Visible = False
                Contract.mRBSetShow True
                Exit Sub
            Case OVSTARTTIMEINDEX
                slStr = Contract.edcDGDropDown.Text
                If slStr <> "" Then
                    If Not gValidTime(slStr) Then
                        Beep
                        Contract.edcDGDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = OVDAYSINDEX + 6
                Contract.mOvSetShow imOVBoxNo
                imOVBoxNo = ilBox
                Contract.mOvEnableBox ilBox
                Exit Sub
            Case OVENDTIMEINDEX
                slStr = Contract.edcDGDropDown.Text
                If slStr <> "" Then
                    If Not gValidTime(slStr) Then
                        Beep
                        Contract.edcDGDropDown.SetFocus
                        Exit Sub
                    End If
                End If
                ilBox = OVSTARTTIMEINDEX
                Contract.mOvSetShow imOVBoxNo
                imOVBoxNo = ilBox
                Contract.mOvEnableBox ilBox
                Exit Sub
            Case OVCOPYINDEX
                If smAlterHidden <> "C" Then
                    ilBox = imOVBoxNo - 1
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Else
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = -1
                    Contract.plcOv.Visible = False
                    Contract.pbcOv.Visible = False
                    Contract.mRBSetShow True
                    Exit Sub
                End If
            Case Else
                ilBox = imOVBoxNo - 1
                Contract.mOvSetShow imOVBoxNo
                imOVBoxNo = ilBox
                Contract.mOvEnableBox ilBox
                Exit Sub
        End Select
    End If
    If Contract.plcPower.Visible Then
        If Contract.plcPower.Enabled Then
            Contract.edcPBDate(0).SetFocus
        Else
            Contract.pbcClickFocus.SetFocus
        End If
    End If
    imTabDirection = -1  'Set-Right to left
    Contract.mAllSetShow 2
    If (imCreditRestr = 1) And ((imStep1Index = 0) Or (imStep1Index = 1)) Then
        ilRet = MsgBox("NO NEW ORDERS", vbOKOnly + vbExclamation, "Restriction")
        Contract.cmcCancel.SetFocus
        Exit Sub
    End If
    If Contract.plcPower.Visible Then
        Contract.edcPBDate(0).SetFocus
        Exit Sub
    End If
    If imLnBoxNo = LINECOMMINDEX Then
        If Contract.mLnCommentBranch() Then
            Exit Sub
        End If
    End If
    Select Case imLnBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            If (UBound(tgClfCntr) = 0) And (tgClfCntr(0).ClfRec.iLine <> 1) Then
                imLnRowNo = 0
                Contract.mInitNewLine
                smLnSave(1, imLnRowNo + 1) = "" 'Vehicle list index
                Contract.mSetHdDate False, False
                Contract.pbcLine_Paint
                Contract.vbcLines.Value = 0
                imLnRowNo = 1
                imMoveLnRowNo = imLnRowNo
                Contract.mSetLineScroll False, False
                imSettingValue = True
                Contract.vbcLines.Value = Contract.vbcLines.Min
                imSettingValue = False
            Else
                Contract.mSetLineScroll False, False
                imLnRowNo = 0
                imSettingValue = True
                Contract.vbcLines.Value = Contract.vbcLines.Min
                imSettingValue = False
                Do While (imLnRowNo < UBound(tgClfCntr)) And ((tgClfCntr(imLnRowNo).iCancel) Or (tgClfCntr(imLnRowNo).bPkExpanded = False) Or (tgClfCntr(imLnRowNo).bShowLine = False))
                    imLnRowNo = imLnRowNo + 1
                    'If imLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                    imMoveLnRowNo = mFindPaintLnRowNo(imLnRowNo)
                    If imMoveLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                        '9/14/16: Handle case where all lines are CBS
                        If Contract.vbcLines.Value + 1 > Contract.vbcLines.Max Then
                            Contract.pbcClickFocus.SetFocus
                            Exit Sub
                        End If
                        imSettingValue = True
                        Contract.vbcLines.Value = Contract.vbcLines.Value + 1
                        imSettingValue = False
                    End If
                Loop
                If imLnRowNo >= UBound(tgClfCntr) Then
                    If (smLnSave(1, imLnRowNo + 1) = "") Or (imLnSave(1, imLnRowNo + 1) <= 0) Or (smLnSave(16, imLnRowNo + 1) = "") Then
                        Contract.mInitNewLine
                        smLnSave(1, imLnRowNo + 1) = "" 'Vehicle list index
                        Contract.pbcLine_Paint
                    End If
                End If
                imLnRowNo = imLnRowNo + 1
                'If imLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                imMoveLnRowNo = mFindPaintLnRowNo(imLnRowNo)
                If imMoveLnRowNo > Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1 Then
                    '9/14/16: Handle case where all lines are CBS
                    If Contract.vbcLines.Value + 1 > Contract.vbcLines.Max Then
                        Contract.pbcClickFocus.SetFocus
                        Exit Sub
                    End If
                    imSettingValue = True
                    Contract.vbcLines.Value = Contract.vbcLines.Value + 1
                    imSettingValue = False
                End If
            End If
            mGetFlightDays imLnRowNo, imOvDays()
            'Contract.mCntrTotal 'Obtain total for new line
            If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                ilBox = DPINDEX
            Else
                ilBox = VEHINDEX
            End If
            imLnBoxNo = ilBox
            Contract.mLnEnableBox ilBox
            Exit Sub
        Case VEHINDEX 'Name (first control within header)
            Contract.mLnSetShow imLnBoxNo
            'ilBox = PRICEINDEX
            Do
                If imLnRowNo <= 1 Then
                    imLnBoxNo = -1
                    Contract.pbcHdTab.SetFocus
                    Exit Sub
                End If
                imLnRowNo = imLnRowNo - 1
                smLineType = mGetLineType(imLnRowNo)
                If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
                    If (tgClfCntr(imLnRowNo).bPkExpanded = True) Then
                        imMoveLnRowNo = imMoveLnRowNo + 1
                    End If
                Else
                    If (tgClfCntr(imLnRowNo).bShowLine = True) Then
                        imMoveLnRowNo = imMoveLnRowNo + 1
                    End If
                End If
                If imLnRowNo < Contract.vbcLines.Value + 1 Then
                    imSettingValue = True
                    Contract.vbcLines.Value = Contract.vbcLines.Value - 1
                    imSettingValue = False
                End If
                '12/16/14
            'Loop While tgClfCntr(imLnRowNo).iCancel
            Loop While tgClfCntr(imLnRowNo).iCancel Or (tgClfCntr(imLnRowNo).bPkExpanded = False) Or (tgClfCntr(imLnRowNo).bShowLine = False)
            Screen.MousePointer = vbHourglass
            mGetFlightDays imLnRowNo, imOvDays()
            Contract.mCntrTotal 'Obtain total for new line, placed back into service 7/15/97
            smLineType = mGetLineType(imLnRowNo)
            'If (smLineType <> "O") And (smLineType <> "A") Then
            '    ilBox = PRICEINDEX
            'Else
                ilBox = LINECOMMINDEX
            'End If
            imLnBoxNo = ilBox
            If Contract.plcDemoBar.Visible Then
                imComputeReq = True
                'cmcCompute.Enabled = True
                If imDBTotals < 5 Then
                    mSetDemoBarBooks imLnRowNo, ilFound, imCppCpm, 2, imCntrVefCode, imVpfIndex, imTerminate, imChgMode 'Using ilfound for ilDnfCode (Contract.mBuildWkInfo sets tgClf to DnfCode)
                    Contract.mBuildWkInfo True, True, False
                    Contract.pbcDemoBar(0).Cls
                    Contract.pbcDemoBar(1).Cls
                    Contract.pbcDemoBar_Paint 0
                    Contract.pbcDemoBar_Paint 1
                    DoEvents
                Else
                    Contract.mBuildWkInfo True, True, False
                    Contract.pbcDemoBar(1).Cls
                    Contract.pbcDemoBar_Paint 1
                End If
            End If
            Contract.mLnEnableBox ilBox
            Screen.MousePointer = vbDefault
            Exit Sub
        Case DPINDEX 'Name (first control within header)
            If tgClfCntr(imLnRowNo - 1).iInPast Then    'bypass vehicle
                Contract.mLnSetShow imLnBoxNo
                'ilBox = PRICEINDEX
                Do
                    If imLnRowNo <= 1 Then
                        imLnBoxNo = -1
                        Contract.pbcHdTab.SetFocus
                        Exit Sub
                    End If
                    imLnRowNo = imLnRowNo - 1
                    If imLnRowNo < Contract.vbcLines.Value + 1 Then
                        imSettingValue = True
                        Contract.vbcLines.Value = Contract.vbcLines.Value - 1
                        imSettingValue = False
                    End If
                Loop While tgClfCntr(imLnRowNo).iCancel
                Screen.MousePointer = vbHourglass
                Contract.mCntrTotal 'Obtain total for new line
                smLineType = mGetLineType(imLnRowNo)
                'If (smLineType <> "O") And (smLineType <> "A") Then
                '    ilBox = PRICEINDEX
                'Else
                    ilBox = LINECOMMINDEX
                'End If
                imLnBoxNo = ilBox
                If Contract.plcDemoBar.Visible Then
                    If imDBTotals < 5 Then
                        Contract.mBuildWkInfo True, True, False
                        Contract.pbcDemoBar(0).Cls
                        Contract.pbcDemoBar(1).Cls
                        Contract.pbcDemoBar_Paint 0
                        Contract.pbcDemoBar_Paint 1
                    Else
                        Contract.mBuildWkInfo True, True, False
                        Contract.pbcDemoBar(1).Cls
                        Contract.pbcDemoBar_Paint 1
                    End If
                End If
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ilBox = imLnBoxNo - 1
            End If
        Case OV1INDEX
            If Contract.edcAct1.Visible Then
                Contract.edcAct1.Visible = False
                imOVBoxNo = -1
            End If
            Select Case imOVBoxNo
                Case -1
                    If smAlterHidden = "C" Then
                        imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                        ilBox = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                    Else
                        If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                            If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                                imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                ilBox = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            Else
                            '    imOVBoxNo = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                            '    ilBox = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                If bmAct1PackageEnable Then
                                    If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                                        imOVBoxNo = ACT1LINEUPINDEX
                                        ilBox = ACT1LINEUPINDEX
                                    Else
                                        imOVBoxNo = ACT1DAYPARTFILTERINDEX
                                        ilBox = ACT1DAYPARTFILTERINDEX
                                    End If
                                Else
                                    imOVBoxNo = OVCOPYINDEX
                                    ilBox = OVCOPYINDEX
                                End If
                            End If
                        Else
                            If (Not mIsLnGame(imLnRowNo)) Then
                                imOVBoxNo = OVDAYSINDEX
                                ilBox = OVDAYSINDEX
                            Else
                                If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
                                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                        imOVBoxNo = OVBBOPENINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                        ilBox = OVBBOPENINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                    Else
                                        'imOVBoxNo = OVBBCLOSEINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                        'ilBox = OVBBCLOSEINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                        If bmAct1PackageEnable Then
                                            If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                                                imOVBoxNo = ACT1LINEUPINDEX
                                                ilBox = ACT1LINEUPINDEX
                                            Else
                                                imOVBoxNo = ACT1DAYPARTFILTERINDEX
                                                ilBox = ACT1DAYPARTFILTERINDEX
                                            End If
                                        Else
                                            imOVBoxNo = OVCOPYINDEX
                                            ilBox = OVCOPYINDEX
                                        End If
                                    End If
                                Else
                                    'imOVBoxNo = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                    'ilBox = OVHIGHINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                                    If bmAct1PackageEnable Then
                                        If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                                            imOVBoxNo = ACT1LINEUPINDEX
                                            ilBox = ACT1LINEUPINDEX
                                        Else
                                            imOVBoxNo = ACT1DAYPARTFILTERINDEX
                                            ilBox = ACT1DAYPARTFILTERINDEX
                                        End If
                                    Else
                                        imOVBoxNo = OVCOPYINDEX
                                        ilBox = OVCOPYINDEX
                                    End If
                                End If
                            End If
                        End If
                    End If
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OVWDINDEX
                    Contract.mOvSetShow imOVBoxNo
                    If (imDGSave(2) = 0) And (tgSpf.sAllowDailyBuys = "Y") Then
                        'Test if daily spots defined
                        ilSpots = 0
                        For ilLoop = 0 To 6 Step 1
                            ilSpots = ilSpots + Val(smDGSave(3 + ilLoop))
                        Next ilLoop
                        If ilSpots = 0 Then
                            Beep
                            imOVBoxNo = OVDAYSINDEX
                            Contract.mOvEnableBox imOVBoxNo
                            Exit Sub
                        End If
                    Else
                        'Weekly buy
                    End If
                    imOVBoxNo = -1
                    '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                    'ilBox = DPINDEX
                    ilBox = LENINDEX
                Case OVDAYSINDEX
                    If (tgSpf.sAllowDailyBuys = "Y") And (Not tgClfCntr(imLnRowNo - 1).iInPast) Then
                        Contract.mOvSetShow imOVBoxNo
                        ilBox = imOVBoxNo - 2
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = -1
                        '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                        'ilBox = DPINDEX
                        ilBox = LENINDEX
                    End If
                Case OVSTARTTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                    ilBox = OVDAYSINDEX + 6
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OVENDTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                    End If
                    ilBox = OVSTARTTIMEINDEX
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OVCOPYINDEX
                    If smAlterHidden = "C" Then
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = -1
                        ilBox = LINECOMMINDEX
                    Else
                        If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                            'ilBox = DPINDEX
                            ilBox = LENINDEX
                        Else
                            If (Not mIsLnGame(imLnRowNo)) Then
                                If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                                    ilBox = OVDAYSINDEX + 6
                                Else
                                    ilBox = OVENDTIMEINDEX
                                End If
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                If (tgSpf.sCPkEqual = "Y") Then
                                    ilBox = OVRANKINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                End If
                                If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                                    ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                End If
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                                'ilBox = DPINDEX
                                ilBox = LENINDEX
                            End If
                        End If
                    End If
                Case OVHIGHINDEX
                    'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                    '    ilBox = OVCOPYINDEX '
                    '    Contract.mOvSetShow imOVBoxNo
                    '    imOVBoxNo = ilBox
                    '    Contract.mOvEnableBox ilBox
                    '    Exit Sub
                    'Else
                        If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                            'ilBox = DPINDEX
                            ilBox = LENINDEX
                        Else
                            If (Not mIsLnGame(imLnRowNo)) Then
                                If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                                    ilBox = OVDAYSINDEX + 6
                                Else
                                    ilBox = OVENDTIMEINDEX
                                End If
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                                'ilBox = DPINDEX
                                ilBox = LENINDEX
                            End If
                        End If
                    'End If
                Case OVRANKINDEX
                    If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                        ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                    'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                    '    ilBox = OVCOPYINDEX '
                    '    Contract.mOvSetShow imOVBoxNo
                    '    imOVBoxNo = ilBox
                    '    Contract.mOvEnableBox ilBox
                    '    Exit Sub
                    'Else
                        If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                            'ilBox = DPINDEX
                            ilBox = LENINDEX
                        Else
                            If (Not mIsLnGame(imLnRowNo)) Then
                                If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                                    ilBox = OVDAYSINDEX + 6
                                Else
                                    ilBox = OVENDTIMEINDEX
                                End If
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                                'ilBox = DPINDEX
                                ilBox = LENINDEX
                            End If
                        End If
                    'End If
                Case OVBBOPENINDEX
                    If (tgSpf.sCPkEqual = "Y") Then
                        ilBox = OVRANKINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                    If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
                        ilBox = OVHIGHINDEX 'OVPREEMPTINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    End If
                    'If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                    '    ilBox = OVCOPYINDEX '
                    '    Contract.mOvSetShow imOVBoxNo
                    '    imOVBoxNo = ilBox
                    '    Contract.mOvEnableBox ilBox
                    '    Exit Sub
                    'Else
                        If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                            'ilBox = DPINDEX
                            ilBox = LENINDEX
                        Else
                            If (Not mIsLnGame(imLnRowNo)) Then
                                If (Not tgClfCntr(imLnRowNo - 1).iOverride) Then
                                    ilBox = OVDAYSINDEX + 6
                                Else
                                    ilBox = OVENDTIMEINDEX
                                End If
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                '7/29/15: sWAPPED lENGTH AND oVERRIDE COLUMNS
                                'ilBox = DPINDEX
                                ilBox = LENINDEX
                            End If
                        End If
                    'End If
                Case Else
                    Contract.mOvSetShow imOVBoxNo
                    ilBox = imOVBoxNo - 1
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
            End Select
        Case OV2INDEX
            If imOVBoxNo = OV2SPLITINDEX Then
                If Contract.mSplitBranch() Then
                    Exit Sub
                End If
            End If
            Select Case imOVBoxNo
                Case -1
                    ilBox = imShowOV2
                    Contract.mOvEnableBox ilBox
                    Exit Sub
                Case OV2DAYSINDEX
                    Contract.mOvSetShow imOVBoxNo
                    imOVBoxNo = -1
                    ilBox = OV1INDEX
                Case OV2STARTTIMEINDEX
                    slStr = Contract.edcDGDropDown.Text
                    If slStr <> "" Then
                        If Not gValidTime(slStr) Then
                            Beep
                            Contract.edcDGDropDown.SetFocus
                            Exit Sub
                        End If
                    Else
                        smLnSave(26, imLnRowNo) = ""
                    End If
                    If imLnSave(8, imLnRowNo) = 1 Then
                        ilBox = OV2DAYSINDEX + 6
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = -1
                        ilBox = OV1INDEX
                    End If
                Case OV2SPLITINDEX
                    If ((Asc(tgSpf.sOverrideOptions) And SPPREFERREDDT) = SPPREFERREDDT) And (Not mIsLnGame(imLnRowNo)) And (Not tgClfCntr(imLnRowNo - 1).iInPast) Then
                        If Trim$(smLnSave(26, imLnRowNo)) = "" Then
                            ilBox = OV2STARTTIMEINDEX
                        Else
                            ilBox = OV2ENDTIMEINDEX
                        End If
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = -1
                        ilBox = OV1INDEX
                    End If
                Case OV2ALLOCATIONINDEX
                    If mRegionAllowed(imLnRowNo) Then  'Allocation
                        ilBox = OV2SPLITINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If ((Asc(tgSpf.sOverrideOptions) And SPPREFERREDDT) = SPPREFERREDDT) And (Not mIsLnGame(imLnRowNo)) And (Not tgClfCntr(imLnRowNo - 1).iInPast) Then
                            If Trim$(smLnSave(26, imLnRowNo)) = "" Then
                                ilBox = OV2STARTTIMEINDEX
                            Else
                                ilBox = OV2ENDTIMEINDEX
                            End If
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = -1
                            ilBox = OV1INDEX
                        End If
                    End If
                Case OV2ACQUISITIONINDEX
                    If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                        ilBox = OV2ALLOCATIONINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If mRegionAllowed(imLnRowNo) Then  'Allocation
                            ilBox = OV2SPLITINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If ((Asc(tgSpf.sOverrideOptions) And SPPREFERREDDT) = SPPREFERREDDT) And (Not mIsLnGame(imLnRowNo)) And (Not tgClfCntr(imLnRowNo - 1).iInPast) Then
                                If Trim$(smLnSave(26, imLnRowNo)) = "" Then
                                    ilBox = OV2STARTTIMEINDEX
                                Else
                                    ilBox = OV2ENDTIMEINDEX
                                End If
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                ilBox = OV1INDEX
                            End If
                        End If
                    End If
                Case OV2SOLOINDEX
                    'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                    If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                        ilBox = OV2ACQUISITIONINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                            ilBox = OV2ALLOCATIONINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If (tgClfCntr(imLnRowNo - 1).iInPast) Then
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = -1
                                ilBox = OV1INDEX
                            Else
                                If mRegionAllowed(imLnRowNo) Then  'Allocation
                                    ilBox = OV2SPLITINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                Else
                                    If ((Asc(tgSpf.sOverrideOptions) And SPPREFERREDDT) = SPPREFERREDDT) And (Not mIsLnGame(imLnRowNo)) Then
                                        If Trim$(smLnSave(26, imLnRowNo)) = "" Then
                                            ilBox = OV2STARTTIMEINDEX
                                        Else
                                            ilBox = OV2ENDTIMEINDEX
                                        End If
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = ilBox
                                        Contract.mOvEnableBox ilBox
                                        Exit Sub
                                    Else
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = -1
                                        ilBox = OV1INDEX
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case OV21STINDEX
                    If (Asc(tgSpf.sOverrideOptions) And SPSOLOAVAIL) = SPSOLOAVAIL Then
                        ilBox = OV2SOLOINDEX
                        Contract.mOvSetShow imOVBoxNo
                        imOVBoxNo = ilBox
                        Contract.mOvEnableBox ilBox
                        Exit Sub
                    Else
                        'If (mUsingAcqCost()) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (tgUrf(0).sChgAcq = "I") Then
                        If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") And (Contract.mAllowedToChgAcqCost(imLnRowNo)) Then
                            ilBox = OV2ACQUISITIONINDEX
                            Contract.mOvSetShow imOVBoxNo
                            imOVBoxNo = ilBox
                            Contract.mOvEnableBox ilBox
                            Exit Sub
                        Else
                            If ((Asc(tgSpf.sOverrideOptions) And SPALLOCATION) = SPALLOCATION) And (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                                ilBox = OV2ALLOCATIONINDEX
                                Contract.mOvSetShow imOVBoxNo
                                imOVBoxNo = ilBox
                                Contract.mOvEnableBox ilBox
                                Exit Sub
                            Else
                                If mRegionAllowed(imLnRowNo) Then  'Allocation
                                    ilBox = OV2SPLITINDEX
                                    Contract.mOvSetShow imOVBoxNo
                                    imOVBoxNo = ilBox
                                    Contract.mOvEnableBox ilBox
                                    Exit Sub
                                Else
                                    If ((Asc(tgSpf.sOverrideOptions) And SPPREFERREDDT) = SPPREFERREDDT) And (Not mIsLnGame(imLnRowNo)) And (Not tgClfCntr(imLnRowNo - 1).iInPast) Then
                                        If Trim$(smLnSave(26, imLnRowNo)) = "" Then
                                            ilBox = OV2STARTTIMEINDEX
                                        Else
                                            ilBox = OV2ENDTIMEINDEX
                                        End If
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = ilBox
                                        Contract.mOvEnableBox ilBox
                                        Exit Sub
                                    Else
                                        Contract.mOvSetShow imOVBoxNo
                                        imOVBoxNo = -1
                                        ilBox = OV1INDEX
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case Else
                    Contract.mOvSetShow imOVBoxNo
                    ilBox = imOVBoxNo - 1
                    imOVBoxNo = ilBox
                    Contract.mOvEnableBox ilBox
                    Exit Sub
            End Select
        Case WKSPOTSINDEX To WKSPOTSINDEX + 12, PRICEINDEX, LINECOMMINDEX  'LINECOMMINDEX
            'slStr = UCase$(Left$(Contract.edcLNDropDown.Text, 1))
            slCode = Trim$(Contract.edcLNDropDown.Text)
            ilIndex = InStr(1, slCode, "P", 1)
            If ilIndex = 0 Then
                ilIndex = InStr(1, slCode, "Q", 1)
            End If
            If ilIndex = 0 Then
                ilIndex = InStr(1, slCode, "C", 1)
            End If
            If (imLnBoxNo >= WKSPOTSINDEX) And (imLnBoxNo <= WKSPOTSINDEX + 12) And (ilIndex > 0) Then
                If ilIndex >= 1 Then
                    If imDGSave(2) = 0 Then
                        Contract.mLnSetShow imLnBoxNo
                    Else
                        Contract.edcLNDropDown.Text = Left$(slCode, ilIndex - 1)
                        If (Trim$(Contract.edcLNDropDown.Text) <> "") Or (lmWkSpots <> 0) Then
                            Contract.mLnSetShow imLnBoxNo
                        End If
                    End If
                End If
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
                If (InStr(1, slCode, "Q", 1) > 0) Or (InStr(1, slCode, "C", 1) > 0) Then
                    ilFound = LINECOMMINDEX
                Else
                    smLineType = mGetLineType(imLnRowNo)
                    If (smLineType <> "O") And (smLineType <> "A") Then
                        If smLineType <> "H" Then
                            ilFound = PRICEINDEX
                        Else
                            slStr = mGetPkLineType(imLnRowNo)
                            If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                                ilFound = LINECOMMINDEX
                            Else
                                ilFound = PRICEINDEX
                            End If
                        End If
                    Else
                        If (tgSpf.sCAudPkg = "Y") Then
                            ilFound = PRICEINDEX
                        Else
                            ilFound = LINECOMMINDEX
                        End If
                    End If
                End If
            ElseIf (imLnBoxNo = WKSPOTSINDEX) Then
                Contract.mLnSetShow imLnBoxNo
                ilBox = WKSPOTSINDEX + 12
                Contract.pbcLine_MouseUp 0, 0, CSng(lmLnHotSpot(2, 1) + 5), CSng(lmLnHotSpot(2, 2))
                imLnBoxNo = ilBox
                Contract.mLnEnableBox ilBox
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                'Skip to next allowed week
                Contract.mLnSetShow imLnBoxNo
                ilFound = -1
                ilLastBox = -1
                If imLnBoxNo > WKSPOTSINDEX Then
                    smLineType = mGetLineType(imLnRowNo)
                    If smLnSave(14, imLnRowNo) = "" Then    '14=Start Date
                        If imLnBoxNo = PRICEINDEX Then  'LINECOMMINDEX Then
                            ilFound = WKSPOTSINDEX + 12
                        ElseIf imLnBoxNo = LINECOMMINDEX Then
                            If (smLineType <> "O") And (smLineType <> "A") Then
                                If smLineType <> "H" Then
                                    'ilFound = PRICEINDEX
                                   '10843 remove dead code
'                                    If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                        ilBox = WKSPOTSINDEX + 12
'                                    Else
                                        ilFound = PRICEINDEX
                                   ' End If
                                Else
                                    slStr = mGetPkLineType(imLnRowNo)
                                    If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                                        ilBox = WKSPOTSINDEX + 12
                                    Else
                                        'ilFound = PRICEINDEX
                                       '10843 remove dead code
'                                        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                            ilBox = WKSPOTSINDEX + 12
'                                        Else
                                            ilFound = PRICEINDEX
                                      '  End If
                                    End If
                                End If
                            Else
                                If (tgSpf.sCAudPkg = "Y") Then
                                    'ilFound = PRICEINDEX
                                    '10843 remove dead code
'                                    If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                        ilBox = WKSPOTSINDEX + 12
'                                    Else
                                        ilFound = PRICEINDEX
                                   ' End If
                                Else
                                    If smLineType = "A" Then
                                        ilFound = LENINDEX
                                    Else
                                        ilBox = WKSPOTSINDEX + 12
                                    End If
                                End If
                            End If
                        Else
                            ilFound = imLnBoxNo - 1
                        End If
                    Else
                        If imLnBoxNo = PRICEINDEX Then  'LINECOMMINDEX Then
                            ilBox = WKSPOTSINDEX + 12
                        ElseIf imLnBoxNo = LINECOMMINDEX Then
                            If (smLineType <> "O") And (smLineType <> "A") Then
                                If smLineType <> "H" Then
                                    'ilFound = PRICEINDEX
                                   '10843 remove dead code
'                                    If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                        ilBox = WKSPOTSINDEX + 12
'                                    Else
                                        ilFound = PRICEINDEX
                                   ' End If
                                Else
                                    slStr = mGetPkLineType(imLnRowNo)
                                    If (slStr = "E") Or (((slStr = "O") Or (slStr = "A")) And (tgSpf.sCAudPkg = "Y")) Then
                                        ilBox = WKSPOTSINDEX + 12
                                    Else
                                        'ilFound = PRICEINDEX
                                        '10843 remove dead code
'                                        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                            ilBox = WKSPOTSINDEX + 12
'                                        Else
                                            ilFound = PRICEINDEX
                                       ' End If
                                    End If
                                End If
                            Else
                                If (tgSpf.sCAudPkg = "Y") Then
                                    'ilFound = PRICEINDEX
                                   '10843 remove dead code
'                                    If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType <= 1) Then
'                                        ilBox = WKSPOTSINDEX + 12
'                                    Else
                                        ilFound = PRICEINDEX
                                   ' End If
                                Else
                                    If smLineType = "A" Then
                                        ilFound = LENINDEX
                                    Else
                                        ilBox = WKSPOTSINDEX + 12
                                    End If
                                End If
                            End If
                        Else
                            ilBox = imLnBoxNo - 1
                        End If
                        'llLnStartDate = gDateValue(gObtainPrevMonday(smLnSave(14, imLnRowNo)))
                        llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 0)
                        If llAllowedChgWkStartDate <= tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
                            llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate + 1
                        End If
                        Do While (ilBox >= WKSPOTSINDEX) And (ilBox <= WKSPOTSINDEX + 12)
                            llDate = lmHSStartDate + 7 * (imHSCurWkNo + ilBox - WKSPOTSINDEX - 1)
                            'If (llAllowedChgWkStartDate <= llDate) And (llDate >= gDateValue(gObtainPrevMonday(smLnSave(14, imLnRowNo)))) And ((llDate <= gDateValue(gObtainNextSunday(smLnSave(15, imLnRowNo))))) Then
                            If (llAllowedChgWkStartDate <= llDate) Then
                                ilFound = ilBox
                                Exit Do
                            End If
                            ilBox = ilBox - 1
                        Loop
                    End If
                End If
                If ilFound <= 0 Then
                    ilFound = LENINDEX
                    If Contract.plcDemoBar.Visible Then
                        imDBBoxNo = -1
                    End If
                End If
                Contract.mCntrTotal 'Obtain total for new line
            End If
            If (ilBox = PRICEINDEX) And (Not Contract.mAllowedToChgSpotPrice(imLnRowNo)) Then
                mSetPriceInTab PRICEINDEX
                ilFound = LINECOMMINDEX
            End If
            ilBox = ilFound
            imLnBoxNo = ilBox
            Contract.mLnEnableBox ilBox
            Exit Sub
        Case Else
            ilBox = imLnBoxNo - 1
    End Select
    Contract.mLnSetShow imLnBoxNo
    imLnBoxNo = ilBox
    Contract.mLnEnableBox ilBox
End Sub

Public Sub mLnSetShowSub(ilBoxNo As Integer, LINECOMMINDEX As Integer, PRICEINDEX As Integer, OV1INDEX As Integer, LENINDEX As Integer, OV2INDEX As Integer, VEHINDEX As Integer, DPINDEX As Integer, WKSPOTSINDEX As Integer)
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilCode As Integer
    Dim ilBox As Integer
    Dim llSpotCount As Long
    Dim ilDatesBilled As Integer
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llActPrice As Long      'Actual price
    Dim ilChg As Integer
    Dim slName As String
    Dim ilCff As Integer
    Dim ilCgf As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim slPrice As String
    Dim ilOvChg As Integer
    Dim ilStdOvChg As Integer
    Dim ilRet As Integer
    Dim ilPaintLnRowNo As Integer
    Dim ilSvLnRowNo As Integer
    Dim ilIndex As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llWkSpots As Long
    Dim ilBypassLnTotal As Integer
    Dim llAllowedChgWkStartDate As Long
    Dim llEndPrevPeriod As Long
    ReDim ilOVDays(0 To 6) As Integer
    Dim slWDFlag As String
    ReDim ilDailySpots(0 To 6) As Integer
    Dim ilBilledCount As Integer
    '11/15/14: Handle changing Non-Sport to Sport vehicle, remove flights
    Dim blRemoveFlights As Boolean
    Dim ilRdfCode As Integer
    Dim ilRdf As Integer
    '10/20/15: Disallow a schedule line to be changed from stanadrd vehicle to package vehicle
    Dim slOrigType As String
    Dim blDaysChgd As Boolean
    Dim ilOldDnfCode As Integer
    Dim ilDnfCode As Integer
    Dim blVehicleChg As Boolean
    Dim lfBoxW As Single
    
    If (imUpdateAllowed) And (Not igRepNetChgAllowed) Then
        Contract.pbcLineArrow.Visible = False
    End If
    Contract.laclnFrame.Visible = False
    If ilBoxNo < imLBCTCtrls Or ilBoxNo > UBound(tmLnCtrls) Then
        Exit Sub
    End If
    If imLnRowNo <= 0 Then
        Exit Sub
    End If
    '12/9/15: If in view mode, this allow
    Contract.edcLNDropDown.ToolTipText = ""
    smLineType = mGetLineType(imLnRowNo)
    If smLineType = "H" Then
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                'If tgPackLn(ilLoop).iAlter <> True Then
                '10347 added 'R' and flipped logic to make it easier to read- if "C" or "R" and these box are chosen, don't block.
                If Not ((tgPackLn(ilLoop).sAlter = "C" And (ilBoxNo = LINECOMMINDEX Or ilBoxNo = OV1INDEX)) Or (tgPackLn(ilLoop).sAlter = "R" And (ilBoxNo = LINECOMMINDEX Or ilBoxNo = PRICEINDEX))) Then
              '  If (tgPackLn(ilLoop).sAlter <> "C") Or ((ilBoxNo <> LINECOMMINDEX) And (ilBoxNo <> OV1INDEX)) Then
                    If tgPackLn(ilLoop).sAlter <> "Y" Then
                        If Contract.pbcLineArrow.Visible Then
                            Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                        Else
                            'Contract.pbcClickFocus.SetFocus
                        End If
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next ilLoop
    End If
    tgSpf.sAllowDailyBuys = smSvAllowDailyWeekly
    If mIsLnGame(imLnRowNo) Then
        tgSpf.sAllowDailyBuys = "N"
    End If
    ilBox = ilBoxNo
    Select Case ilBox 'Branch on box type (control)
        Case VEHINDEX 'Vehicle
            Contract.lbcLnVehicle(igTabMapIndex).Visible = False
            Contract.lbcPkVehicle(igTabMapIndex).Visible = False
            Contract.pbcLbcVehicle.Visible = False
            Contract.edcLNDropDown.Visible = False
            Contract.cmcLnDropDown.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                '2/7/09: Added to handle case where focus can't be set
                On Error Resume Next
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                On Error GoTo 0
                Exit Sub
            End If
            blVehicleChg = False
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                If smLnSave(1, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                    '1/12/16: disallow hidden line vehicle to be changed to a package vehicle
                    'If tgClfCntr(imLnRowNo - 1).iLineSchd Then
                    If (tgClfCntr(imLnRowNo - 1).iLineSchd) Or (smLineType = "H") Then
                        'Check vehicle types.  Disallow Conventional or selling to be switch to Rep
                        gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
                        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                            slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            ilVef = gBinarySearchVef(CInt(slCode))
                            If ilVef <> -1 Then
                                slOrigType = tgMVef(ilVef).sType
                                If tgMVef(ilVef).sType <> "R" Then
                                    slStr = Contract.edcLNDropDown.Text
                                    gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                                    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                                        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                        ilVef = gBinarySearchVef(CInt(slCode))
                                        If ilVef <> -1 Then
                                            '10/20/15: Disallow a schedule line to be changed from standard vehicle to package vehicle
                                            'If (tgMVef(ilVef).sType = "R") Then
                                            If (tgMVef(ilVef).sType = "R") Or ((tgMVef(ilVef).sType = "P") And (slOrigType <> "P")) Then
                                                Beep
                                                If Contract.pbcLineArrow.Visible Then
                                                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                                Else
                                                    Contract.pbcClickFocus.SetFocus
                                                End If
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Else
                                    slStr = Contract.edcLNDropDown.Text
                                    gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                                    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                                        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                        ilVef = gBinarySearchVef(CInt(slCode))
                                        If ilVef <> -1 Then
                                            If tgMVef(ilVef).sType <> "R" Then
                                                Beep
                                                If Contract.pbcLineArrow.Visible Then
                                                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                                Else
                                                    Contract.pbcClickFocus.SetFocus
                                                End If
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                '10/20/15: Disallow a schedule line to be changed from package vehicle to standard vehicle
                If smLnSave(1, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                    If tgClfCntr(imLnRowNo - 1).iLineSchd Then
                        'Check vehicle types.  Disallow Conventional or selling to be switch to Rep
                        gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
                        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                            slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            ilVef = gBinarySearchVef(CInt(slCode))
                            If ilVef <> -1 Then
                                slOrigType = tgMVef(ilVef).sType
                                slStr = Contract.edcLNDropDown.Text
                                gFindMatch slStr, 0, Contract.lbcLnVehicle(igTabMapIndex)
                                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                                    slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    ilVef = gBinarySearchVef(CInt(slCode))
                                    If ilVef <> -1 Then
                                        If ((tgMVef(ilVef).sType <> "P") And (slOrigType = "P")) Then
                                            Beep
                                            If Contract.pbcLineArrow.Visible Then
                                                Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                            Else
                                                Contract.pbcClickFocus.SetFocus
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            slStr = Contract.edcLNDropDown.Text
            gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilBox)
            smLnShow(VEHINDEX, imLnRowNo) = tmLnCtrls(ilBox).sShow
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                If smLnSave(1, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                    blVehicleChg = True
                    bmLnChgRegenResearch = True
                    smLnSave(1, imLnRowNo) = Contract.lbcLnVehicle(igTabMapIndex).List(Contract.lbcLnVehicle(igTabMapIndex).ListIndex)
                    imComputeReq = True
                    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                        imLnChg = True
                    End If
                    'Remove daypart
                    imLnSave(1, imLnRowNo) = -1
                    smLnShow(DPINDEX, imLnRowNo) = ""
                    '10256
                    'smLnSave(29, imLnRowNo) = Contract.mObtainAct1LineupDefault(smLnSave(1, imLnRowNo))
                    '10325
                    mObtainAct1DefaultsForPackages imLnRowNo
                    tgClfCntr(imLnRowNo - 1).iPriChgd = True
                    tgClfCntr(imLnRowNo - 1).ClfRec.lghfcode = 0
                    tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                    mClearAvailAud imLnRowNo, True
                    smLnSave(24, imLnRowNo) = ""
                    blRemoveFlights = False
                    '4/16/15
                    If Contract.lbcLnVehicle(igTabMapIndex).ListIndex <> -1 Then
                        slNameCode = tmVehicleCode(Contract.lbcLnVehicle(igTabMapIndex).ListIndex).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVef = gBinarySearchVef(CInt(slCode))
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType = "G" Then
                                blRemoveFlights = True
                            End If
                        End If
                    '4/16/15
                    Else
                        blRemoveFlights = True
                    End If
                    If (tgClfCntr(imLnRowNo - 1).iFirstCgf <> -1) Or blRemoveFlights Then
                        'Clear flight information
                        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                        Do While ilCff <> -1
                            If tgCffCntr(ilCff).iStatus >= 1 Then
                                tgCffCntr(ilCff).iStatus = 2
                            Else
                                tgCffCntr(ilCff).iStatus = -1
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                        smLnSave(23, imLnRowNo) = ""
                    End If
                    ilCgf = tgClfCntr(imLnRowNo - 1).iFirstCgf
                    Do While ilCgf <> -1
                        tgCgfCntr(ilCgf).CgfRec.iNoSpots = 0
                        If tgCgfCntr(ilCgf).iStatus >= 1 Then
                            tgCgfCntr(ilCgf).iStatus = 2
                        Else
                            tgCgfCntr(ilCgf).iStatus = -1
                        End If
                        ilCgf = tgCgfCntr(ilCgf).iNextCgf
                    Loop
                    mClearAvails imLnRowNo
                End If
            Else
                If smLnSave(1, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                    blVehicleChg = True
                    bmLnChgRegenResearch = True
                    tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                    smLnSave(1, imLnRowNo) = Contract.edcLNDropDown.Text
                    For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                            tgPackLn(ilLoop).sVehName = smLnSave(1, imLnRowNo)
                            mPkVehPop Contract.lbcPkVehicle(igTabMapIndex), True, False
                            Exit For
                        End If
                    Next ilLoop

                    imComputeReq = True
                    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                        imLnChg = True
                    End If
                    'Remove daypart
                    smLnSave(20, imLnRowNo) = ""
                    smLnShow(DPINDEX, imLnRowNo) = ""
                    '10256
                    'smLnSave(29, imLnRowNo) = Contract.mObtainAct1LineupDefault(smLnSave(1, imLnRowNo))
                    '10325
                    mObtainAct1DefaultsForPackages imLnRowNo
                End If
            End If
            mGetOwnership imLnRowNo
            Contract.mLenPop
            If imTerminate Then
                Exit Sub
            End If
            If imCntrVefCode > 0 Then
                tgClfCntr(imLnRowNo - 1).ClfRec.iVefCode = imCntrVefCode
            End If
            If blVehicleChg = True Then
                Contract.mPrgPop imLnRowNo, "L"
                If imTerminate Then
                    Exit Sub
                End If
            End If
            'Set LLD for vehicle
            If imVpfIndex >= 0 Then
                gUnpackDateLong tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), tgClfCntr(imLnRowNo - 1).lLLD
                If (tgClfCntr(imLnRowNo - 1).lLLD <> 0) And (gWeekDayLong(tgClfCntr(imLnRowNo - 1).lLLD) <> 6) Then
                    tgClfCntr(imLnRowNo - 1).lLLD = tgClfCntr(imLnRowNo - 1).lLLD - gWeekDayLong(tgClfCntr(imLnRowNo - 1).lLLD) - 1
                End If
            Else
                tgClfCntr(imLnRowNo - 1).lLLD = 0
            End If
        Case DPINDEX 'Daypart
            Contract.lbcLnProgram(igTabMapIndex).Visible = False
            Contract.lbcPkProgram(igTabMapIndex).Visible = False
            Contract.edcLNDropDown.Visible = False
            Contract.cmcLnDropDown.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            'If Contract.tscLine.SelectedItem.Index = 10 Then
            If Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH) Then
                If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
                    If Contract.pbcLineArrow.Visible Then
                        Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                    Else
                        Contract.pbcClickFocus.SetFocus
                    End If
                    Exit Sub
                End If
                slStr = Contract.edcLNDropDown.Text
                If Contract.lbcLnProgram(igTabMapIndex).ListIndex < 0 Then
                    ilDnfCode = 0
                Else
                    ilDnfCode = Contract.lbcLnProgram(igTabMapIndex).ItemData(Contract.lbcLnProgram(igTabMapIndex).ListIndex)
                End If
                If (tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode <> ilDnfCode) And (ilDnfCode <> 0) Then
                    ilOldDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode
                    bmLnChgRegenResearch = True
                    tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                    imComputeReq = True
                    tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode = ilDnfCode
                    tgClfCntr(imLnRowNo - 1).lPop = 0
                    tgClfCntr(imLnRowNo - 1).iResearch = False
                    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                        imLnChg = True
                    End If
                    
                    ilRet = MsgBox("Apply New Book to All Applicable Lines", vbYesNo + vbDefaultButton2 + vbQuestion, "Books")
                    If ilRet = vbYes Then
                        mSetDnfCode ilOldDnfCode
                    End If
                    mGetMajorityPopulation
                End If
            Else
                ilRdfCode = -1
                slStr = Contract.edcLNDropDown.Text
                If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                    If Contract.lbcLnProgram(igTabMapIndex).ListIndex < 0 Then
                        ilCode = -1
                    Else
                        ilCode = Contract.lbcPrg(igTabMapIndex).ItemData(Contract.lbcLnProgram(igTabMapIndex).ListIndex)
                        If imLnSave(1, imLnRowNo) <> ilCode Then
                            bmLnChgRegenResearch = True
                            tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                            imComputeReq = True
                            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                                imLnChg = True
                                mDPAllowedDays ilCode, imAllowedDays()
                                mChgDaypart imLnRowNo, imAllowedDays()
                            End If
                            imLnSave(1, imLnRowNo) = ilCode
                            ilRdfCode = ilCode
                            mGetFlightDays imLnRowNo, imOvDays()
                            For ilDay = 0 To 6 Step 1
                                imLnSave(12 + ilDay, imLnRowNo) = imOvDays(ilDay)
                            Next ilDay
                            slName = mSetDPShowName(imLnRowNo, ilBoxNo, imOvDays())
                            tgClfCntr(imLnRowNo - 1).iPriChgd = True
                            mClearAvailAud imLnRowNo, True
                        End If
                    End If
                    imLnSave(1, imLnRowNo) = ilCode
                Else
                    If smLnSave(20, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                        bmLnChgRegenResearch = True
                        tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                        smLnSave(20, imLnRowNo) = Contract.edcLNDropDown.Text
                        imComputeReq = True
                        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                            imLnChg = True
                            For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                If StrComp(smLnSave(20, imLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                                    ilRdfCode = tgMRdf(ilLoop).iCode
                                    mDPAllowedDays tgMRdf(ilLoop).iCode, imAllowedDays()
                                    mChgDaypart imLnRowNo, imAllowedDays()
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                        mGetFlightDays imLnRowNo, imOvDays()
                        For ilDay = 0 To 6 Step 1
                            imLnSave(12 + ilDay, imLnRowNo) = imOvDays(ilDay)
                        Next ilDay
                        slName = mSetDPShowName(imLnRowNo, ilBoxNo, imOvDays())
    
                    End If
                End If
                
                If ilRdfCode >= 0 Then
                    ilRdf = gBinarySearchRdf(ilRdfCode)
                    If ilRdf <> -1 Then
                        'Dick's change by Dan  11/3/15
                        'If tgMRdf(ilRdfCode).sTimeOver = "Y" Then
                         If tgMRdf(ilRdf).sTimeOver = "Y" Then
                            tgClfCntr(imLnRowNo - 1).iOverride = True
                        Else
                            tgClfCntr(imLnRowNo - 1).iOverride = False
                        End If
                    End If
                End If
            End If
            mObtainLineFlags imLnRowNo, False, hmCHF, hmClf, hmSdf, hmPsf, hmSmf, imCntrVefCode, imVpfIndex, imTerminate
        Case OV1INDEX    'Part of TPINDEX
            Contract.mOvSetShow imOVBoxNo
            imOVBoxNo = -1
            Contract.pbcOv.Visible = False
            Contract.plcOv.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            If (imLnSave(8, imLnRowNo) <> imDGSave(2)) And (imLnSave(8, imLnRowNo) <> -1) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                    If imDGSave(2) = 0 Then
                        mCFFConvertDW imLnRowNo, "D", smDGSave(), imOvDays()
                    Else
                        mCFFConvertDW imLnRowNo, "W", smDGSave(), imOvDays()
                    End If
                End If
                imLnSave(8, imLnRowNo) = imDGSave(2)
                tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
            ElseIf imLnSave(8, imLnRowNo) = -1 Then
                imLnSave(8, imLnRowNo) = imDGSave(2)
            End If
            ilStdOvChg = 0
            If imLnSave(8, imLnRowNo) = 0 Then
                'Daily
                For ilDay = 0 To 6 Step 1
                    If Val(smLnSave(7 + ilDay, imLnRowNo)) <> Val(smDGSave(3 + ilDay)) Then
                        tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                            imLnChg = True
                        End If
                        ilStdOvChg = ilStdOvChg Or &H8
                    End If
                    smLnSave(7 + ilDay, imLnRowNo) = smDGSave(3 + ilDay)
                    '5/19/12: Save override daily spots
                    imOvDailySpots(ilDay) = Val(smDGSave(3 + ilDay))
                Next ilDay
            Else
                'Weekly
                'Not required as number of spots in a week not changed
            End If
            'Set Live or Recorded
            If imLnSave(31, imLnRowNo) <> imDGSave(7) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H40
                imLnSave(31, imLnRowNo) = imDGSave(7)
            End If
            'set preempt
            If imLnSave(6, imLnRowNo) <> imDGSave(3) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H1
                imLnSave(6, imLnRowNo) = imDGSave(3)
            End If
            'set rank
            If (imLnSave(19, imLnRowNo) <> imDGSave(4)) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H2
                imLnSave(19, imLnRowNo) = imDGSave(4)
            End If
            'set BB Open
            If (imLnSave(20, imLnRowNo) <> imDGSave(5)) Then
                imLnSave(20, imLnRowNo) = imDGSave(5)
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                    If imLnSave(20, imLnRowNo) <> -1 Then
                        mPackageInfoToHidden imLnRowNo, 3
                    End If
                End If
                ilStdOvChg = ilStdOvChg Or &H10
            End If
            'set BB Close
            If (imLnSave(21, imLnRowNo) <> imDGSave(6)) Then
                imLnSave(21, imLnRowNo) = imDGSave(6)
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                    If imLnSave(21, imLnRowNo) <> -1 Then
                        mPackageInfoToHidden imLnRowNo, 4
                    End If
                End If
                ilStdOvChg = ilStdOvChg Or &H20
            End If
            ilOvChg = False
            'Start/End Time
            If (smLnSave(2, imLnRowNo) <> smDGSave(13)) Or (smLnSave(3, imLnRowNo) <> smDGSave(14)) Then
                bmLnChgRegenResearch = True
                imComputeReq = True
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilOvChg = True
                ilStdOvChg = ilStdOvChg Or &H4
                tgClfCntr(imLnRowNo - 1).iPriChgd = True
                tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                mClearAvailAud imLnRowNo, True
            End If
            If (Trim$(smDGSave(13)) = "") And (Trim$(smDGSave(14)) = "") Then
                imLnSave(4, imLnRowNo) = 0
                smLnSave(2, imLnRowNo) = ""
                smLnSave(3, imLnRowNo) = ""
            ElseIf (Trim$(smDGSave(13)) = "") Then
                If Not gValidTime(smDGSave(14)) Then
                    imLnSave(4, imLnRowNo) = 0
                    smLnSave(2, imLnRowNo) = ""
                    smLnSave(3, imLnRowNo) = ""
                Else
                    mGetDPTimes imLnRowNo, smDGSave(13), slStr
                    If smDGSave(13) = "" Then
                        imLnSave(4, imLnRowNo) = 0
                        smLnSave(2, imLnRowNo) = ""
                        smLnSave(3, imLnRowNo) = ""
                    Else
                        imLnSave(4, imLnRowNo) = 1
                        smLnSave(2, imLnRowNo) = smDGSave(13)
                        smLnSave(3, imLnRowNo) = smDGSave(14)
                    End If
                End If
            ElseIf (Trim$(smDGSave(14)) = "") Then
                If Not gValidTime(smDGSave(13)) Then
                    imLnSave(4, imLnRowNo) = 0
                    smLnSave(2, imLnRowNo) = ""
                    smLnSave(3, imLnRowNo) = ""
                Else
                    mGetDPTimes imLnRowNo, slStr, smDGSave(14)
                    If smDGSave(14) = "" Then
                        imLnSave(4, imLnRowNo) = 0
                        smLnSave(2, imLnRowNo) = ""
                        smLnSave(3, imLnRowNo) = ""
                    Else
                        imLnSave(4, imLnRowNo) = 1
                        smLnSave(2, imLnRowNo) = smDGSave(13)
                        smLnSave(3, imLnRowNo) = smDGSave(14)
                    End If
                End If
            ElseIf (Not gValidTime(smDGSave(13))) Or (Not gValidTime(smDGSave(14))) Then
                imLnSave(4, imLnRowNo) = 0
                smLnSave(2, imLnRowNo) = ""
                smLnSave(3, imLnRowNo) = ""
            Else
                imLnSave(4, imLnRowNo) = 1
                smLnSave(2, imLnRowNo) = smDGSave(13)
                smLnSave(3, imLnRowNo) = smDGSave(14)
            End If
            
            'set act1 lineup
            If Trim$(smLnSave(29, imLnRowNo)) <> Trim$(smDGSave(15)) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H80
                smLnSave(29, imLnRowNo) = smDGSave(15)
            End If
            'set act1 stored times
            If imLnSave(32, imLnRowNo) <> imDGSave(8) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H100
                imLnSave(32, imLnRowNo) = imDGSave(8)
            End If
            'set act1 stored spots
            If imLnSave(33, imLnRowNo) <> imDGSave(9) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H200
                imLnSave(33, imLnRowNo) = imDGSave(9)
            End If            'set act1 stored clear %
            If imLnSave(34, imLnRowNo) <> imDGSave(10) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H400
                imLnSave(34, imLnRowNo) = imDGSave(10)
            End If            'set act1 daypart filter
            If imLnSave(35, imLnRowNo) <> imDGSave(11) Then
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                ilStdOvChg = ilStdOvChg Or &H800
                imLnSave(35, imLnRowNo) = imDGSave(11)
            End If
            
            mSetFlightDays imLnRowNo, imOvDays(), imLnRowNo, smDGSave(), True
            slStr = mSetDPShowName(imLnRowNo, DPINDEX, imOvDays())
            slStr = " "
            If Trim$(smLnSave(29, imLnRowNo)) <> "" And bmAct1PackageEnable Then
                slStr = "A"
            Else
                If imLnSave(8, imLnRowNo) = 0 Then
                    slStr = "D"
                Else
                    If Left$(tmLnCtrls(DPINDEX).sShow, 1) = "*" Then
                        slStr = "*"
                    ElseIf (imLnSave(20, imLnRowNo) > 0) Or (imLnSave(21, imLnRowNo) > 0) Then
                        slStr = "*"
                    End If
                End If
            End If
            gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilBox)
            smLnShow(OV1INDEX, imLnRowNo) = tmLnCtrls(ilBox).sShow
            ilChg = False
            If imLnSave(8, imLnRowNo) = 1 Then
                For ilDay = 0 To 6 Step 1
                    If (imLnSave(12 + ilDay, imLnRowNo) <> imOvDays(ilDay)) And (imLnSave(8, imLnRowNo) <> 0) Then
                        tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                        ilChg = True
                    End If
                    imLnSave(12 + ilDay, imLnRowNo) = imOvDays(ilDay)
                    If ((Asc(tgSpf.sOverrideOptions) And &H8) = &H8) Or (imLnSave(24, imLnRowNo) <> -1) Then
                        If imLnSave(12 + ilDay, imLnRowNo) = 1 Then
                            imLnSave(24 + ilDay, imLnRowNo) = 1
                        Else
                            If imLnSave(24 + ilDay, imLnRowNo) = -1 Then
                                imLnSave(24 + ilDay, imLnRowNo) = imLnSave(12 + ilDay, imLnRowNo)
                            End If
                        End If
                    End If
                Next ilDay
            End If
            slStr = Contract.mOV2Defined(imLnRowNo)
            gSetShow Contract.pbcLine, slStr, tmLnCtrls(OV2INDEX)
            smLnShow(OV2INDEX, imLnRowNo) = tmLnCtrls(OV2INDEX).sShow
            If ilChg Then
                ilStdOvChg = ilStdOvChg Or &H8
                ilOvChg = True
                imComputeReq = True
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    imLnChg = True
                End If
                tgClfCntr(imLnRowNo - 1).iPriChgd = True
                mClearAvailAud imLnRowNo, True
            End If
            mSetHiddenOverrides imLnRowNo, ilStdOvChg, DPINDEX, OV1INDEX, smDGSave(), imLnTotalType, PRICEINDEX
            If ilOvChg Then
                Contract.mResetDemoBar False
            End If
            If (imLnSave(8, imLnRowNo) = 0) And ((ilStdOvChg And &H8) = &H8) And (Not imInChkForStdPkg) Then   'Daily
                imPaintLnRowNo = -1
                Contract.pbcLine.Cls
                mLnTotals imLnRowNo, imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), "N", imNoWkDays()
                Contract.mCntrTotal
                Contract.pbcLine_Paint
            End If
        Case OV2INDEX    'Part of TPINDEX
            Contract.mOvSetShow imOVBoxNo
            imOVBoxNo = -1
            Contract.pbcOv.Visible = False
            Contract.plcOv.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                If imSvSolo <> imLnSave(23, imLnRowNo) Then
                    If imLnSave(23, imLnRowNo) <> -1 Then
                        mPackageInfoToHidden imLnRowNo, 1
                    End If
                End If
                If imSv1stPos <> imLnSave(22, imLnRowNo) Then
                    If imLnSave(22, imLnRowNo) <> -1 Then
                        mPackageInfoToHidden imLnRowNo, 2
                    End If
                End If
            End If
            slStr = Contract.mOV2Defined(imLnRowNo)
            gSetShow Contract.pbcLine, slStr, tmLnCtrls(OV2INDEX)
            smLnShow(OV2INDEX, imLnRowNo) = tmLnCtrls(OV2INDEX).sShow
        Case LENINDEX
            Contract.lbcLen.Visible = False
            Contract.edcLNDropDown.Visible = False
            Contract.cmcLnDropDown.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            If smLnSave(16, imLnRowNo) <> Contract.edcLNDropDown.Text Then
                igAnsCMC = -1
                bmLnChgRegenResearch = True
                If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                    '9/12/16: Add message if package line about changing the hidden lines
                    If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
                        'If Trim$(smLnSave(16, imLnRowNo)) <> "" Then
                        igEditBox = 0
                        igDefCMC = 0
                        ilRet = mPackageHiddenLenMatch(imLnRowNo, Val(Contract.edcLNDropDown.Text))
                        If ilRet = 3 Then
                            ilRet = MsgBox("Package Length change not permitted as the length " & Contract.edcLNDropDown.Text & " does not exist in all hidden lines", vbExclamation + vbOKOnly, "Change Disallowed")
                            If Contract.pbcLineArrow.Visible Then
                                Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                            Else
                                Contract.pbcClickFocus.SetFocus
                            End If
                            Exit Sub
                        ElseIf ilRet = 2 Then
                            If Trim$(smLnSave(16, imLnRowNo)) <> "" Then
                                'ilRet = MsgBox("Spot Length has been changed from " & smLnSave(16, imLnRowNo) & " to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Ok to copy spot length to hidden lines.  If the spot prices are incorrect, delete entire package and re-enter", vbOkCancel + vbExclamation, "Warning")
                                sgGenMsg = "Package Spot Length has been changed from " & smLnSave(16, imLnRowNo) & " to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press either Change Hidden or Leave Hidden.  If the spot prices are incorrect, then press Cancel and Delete entire package and re-enter"
                            Else
                                'ilRet = MsgBox("Spot Length set to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Ok to copy spot length to hidden lines.  If the spot prices are incorrect, delete entire package and re-enter", vbOkCancel + vbExclamation, "Warning")
                                sgGenMsg = "Spot Length set to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press either Change Hidden or Leave Hidden.  If the spot prices are incorrect, then press Cancel and delete entire package and re-enter"
                            End If
                            sgCMCTitle(0) = "Change Hidden"
                            sgCMCTitle(1) = "Leave Hidden"
                            sgCMCTitle(2) = "Cancel Package Change"
                            sgCMCTitle(3) = ""
                            GenMsg.Show vbModal
                            'If ilRet = vbCancel Then
                            If igAnsCMC = 2 Then
                                If Contract.pbcLineArrow.Visible Then
                                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                Else
                                    Contract.pbcClickFocus.SetFocus
                                End If
                                Exit Sub
                            End If
                        ElseIf ilRet = 1 Then
                            If Trim$(smLnSave(16, imLnRowNo)) <> "" Then
                                'ilRet = MsgBox("Spot Length has been changed from " & smLnSave(16, imLnRowNo) & " to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Ok to copy spot length to hidden lines.  If the spot prices are incorrect, delete entire package and re-enter", vbOkCancel + vbExclamation, "Warning")
                                sgGenMsg = "Package Spot Length has been changed from " & smLnSave(16, imLnRowNo) & " to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Change Hidden.  If the spot prices are incorrect, then press Cancel and Delete entire package and re-enter"
                            Else
                                'ilRet = MsgBox("Spot Length set to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Ok to copy spot length to hidden lines.  If the spot prices are incorrect, delete entire package and re-enter", vbOkCancel + vbExclamation, "Warning")
                                sgGenMsg = "Spot Length set to " & Contract.edcLNDropDown.Text & ".  If the spot prices are correct, then press Change Hidden.  If the spot prices are incorrect, then press Cancel and delete entire package and re-enter"
                            End If
                            sgCMCTitle(0) = "Change Hidden"
                            sgCMCTitle(1) = "Cancel Package Change"
                            sgCMCTitle(2) = ""
                            sgCMCTitle(3) = ""
                            GenMsg.Show vbModal
                            'If ilRet = vbCancel Then
                            If igAnsCMC = 1 Then
                                If Contract.pbcLineArrow.Visible Then
                                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                                Else
                                    Contract.pbcClickFocus.SetFocus
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                    'If length increase and start date in past- show warning message
                    If (imStep3Index >= 0) And (tgClfCntr(imLnRowNo - 1).iStatus = 1) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M") Or (tgClfCntr(imLnRowNo - 1).iLineSchd)) Then
                        If Val(Contract.edcLNDropDown.Text) > Val(smLnSave(16, imLnRowNo)) Then
                            slStr = smLnSave(14, imLnRowNo)
                            If gValidDate(slStr) Then
                                'If gDateValue(slStr) <= lmNowDate Then
                                If (gDateValue(slStr) <= lmNowDate) And (igAnsCMC = 0) Then
                                    'Warning message
                                    MsgBox "Spot Length increased, for avails on or prior to today, move spots first to insure room in each avail for new spot length or spots will be set to missed", vbOKOnly + vbExclamation, "Warning"
                                End If
                            End If
                        End If
                    End If
                    imLnChg = True
                    'imComputeReq = True
                End If
                slStr = Contract.edcLNDropDown.Text
                gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilBox)
                smLnShow(LENINDEX, imLnRowNo) = tmLnCtrls(ilBox).sShow
                smLnSave(16, imLnRowNo) = Contract.edcLNDropDown.Text
                ilIndex = imLnSave(10, imLnRowNo)
                Do While ilIndex > 0
                    For ilLoop = LBound(tmWkInfo(ilIndex).lSpotCount) To UBound(tmWkInfo(ilIndex).lSpotCount) Step 1
                        tmWkInfo(ilIndex).iAvail(ilLoop) = -29999   'Not defined
                    Next ilLoop
                    ilIndex = tmWkInfo(ilIndex).iNextIndex
                Loop
                If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                    mClearAvails imLnRowNo
                End If
                tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                '9/12/16: Set hidden line spot length
                'If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
                If ((smLineType = "O") Or (smLineType = "A") Or (smLineType = "E")) And (igAnsCMC = 0) Then
                    mPackageInfoToHidden imLnRowNo, 8, LENINDEX
                End If
                'Adjust acquisition cost as length changed
                If Contract.mUsingAcqCost() Then
                    smLnSave(24, imLnRowNo) = ""
                    Contract.mSetDefAcqCost Val(smLnSave(16, imLnRowNo))
                End If
                If (imToggleAvail <> -1) And (Contract.plcDemoBar.Visible) Then
                    Screen.MousePointer = vbHourglass
                    If Val(smLnSave(16, imLnRowNo)) = 10 Then
                        imToggleAvail = 2
                    Else
                        imToggleAvail = 1
                    End If
                    Contract.mBuildWkInfo False, True, True
                    Contract.pbcDemoBar(1).Cls
                    Contract.pbcDemoBar_Paint 1
                    Screen.MousePointer = vbDefault
                End If
            End If
        Case WKSPOTSINDEX To WKSPOTSINDEX + 12
            Screen.MousePointer = vbHourglass
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                imOVBoxNo = -1
                Contract.pbcOVWD(0).Visible = False
                imOVType = 0
                'If (smLineTypeOv <> "P") Then
                '    If tgSpf.sUsingBBs = "Y" Then
                '        Contract.pbcOvWD(0).Height = 1065
                '    Else
                '        Contract.pbcOvWD(0).Height = 720
                '    End If
                'Else
                '    Contract.pbcOvWD(0).Height = 720
                'End If
                If bmAct1PackageEnable Then
                    Contract.pbcOVWD(0).height = 1740    '1410
                Else
                    Contract.pbcOVWD(0).height = 1050    '1065
                End If
                
                Contract.pbcDBLnDaily.Visible = False
                On Error Resume Next
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            imPaintLnRowNo = imLnRowNo
            blDaysChgd = False
            If imLnSave(8, imLnRowNo) = 0 Then   'Daily
                Contract.mOvSetShow imOVBoxNo
                imOVBoxNo = -1
                Contract.pbcOVWD(0).Visible = False
                imOVType = 0
                'If (smLineTypeOv <> "P") Then
                '    If tgSpf.sUsingBBs = "Y" Then
                '        Contract.pbcOvWD(0).Height = 1065
                '    Else
                '        Contract.pbcOvWD(0).Height = 720
                '    End If
                'Else
                '    Contract.pbcOvWD(0).Height = 720
                'End If
                If bmAct1PackageEnable Then
                    Contract.pbcOVWD(0).height = 1740    '1410
                Else
                    Contract.pbcOVWD(0).height = 1050    '1065
                End If
                llSpotCount = 0
                If imLnDaily = 0 Then
                    For ilLoop = 0 To 6 Step 1
                        If imDailySpots(ilLoop) <> Val(smLnSave(7 + ilLoop, imLnRowNo)) Then
                            tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                            blDaysChgd = True
                        End If
                        ilDailySpots(ilLoop) = Val(smLnSave(7 + ilLoop, imLnRowNo))
                        llSpotCount = llSpotCount + Val(smLnSave(7 + ilLoop, imLnRowNo))
                    Next ilLoop
                End If
            Else
                slStr = Contract.edcLNDropDown.Text
                llSpotCount = Val(slStr)
                For ilLoop = 0 To 6 Step 1
                    If imLnSave(12 + ilLoop, imLnRowNo) = 0 Then
                        If Not imWkDays(ilLoop) Then
                            blDaysChgd = True
                        End If
                        imWkDays(ilLoop) = True
                    Else
                        If imWkDays(ilLoop) Then
                            blDaysChgd = True
                        End If
                        imWkDays(ilLoop) = False
                    End If
                Next ilLoop
                
            End If
            If llSpotCount > 0 Then
                lmLastWkDateAdded = lmHSStartDate + 7 * (imHSCurWkNo + ilBox - WKSPOTSINDEX - 1)
                imLastLnDateAdded = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
            End If
            If (lmWkSpots <> llSpotCount) Or blDaysChgd Then
                bmLnChgRegenResearch = True
                tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                If (tgClfCntr(imLnRowNo - 1).iFirstCff < 0) Or (tgClfCntr(imLnRowNo - 1).iCancel) Then
                    If lmHSStartDate > 0 Then
                        llDate = gDateValue(gObtainPrevMonday(Format$(lmHSStartDate, "m/d/yy"))) + 7 * (imHSCurWkNo + ilBox - WKSPOTSINDEX - 1)
                    Else
                        llDate = gDateValue(gObtainPrevMonday(Format$(gNow(), "m/d/yy"))) + 7 * (ilBox - WKSPOTSINDEX)
                    End If
                Else
                    llDate = lmHSStartDate + 7 * (imHSCurWkNo + ilBox - WKSPOTSINDEX - 1)
                End If
                'Test if in past and if message asked
                ilPaintLnRowNo = imPaintLnRowNo
                ilChg = True
                If (tgClfCntr(imLnRowNo - 1).lUnbilledDate > 0) And (llDate < tgClfCntr(imLnRowNo - 1).lUnbilledDate) Then
                    MsgBox "Week Previously Billed, Spots can't be Changed", vbOKOnly + vbExclamation, "Warning"
                    ilChg = False
                    imPaintLnRowNo = -1
                End If
                   
                If ilChg Then
                    '3/14/13:  Allow spots to be removed if billing by Order
                    If (llSpotCount = 0) And (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
                        ilChg = mReplaceWithFills(hmCHF, hmClf, hmSdf, hmSmf, llDate, lmMonNowDate, imPastMsgAsked, "W")
                        If Not ilChg Then
                            imPaintLnRowNo = -1
                        End If
                    Else
                        '3/14/13: Code below not changed
                        '8/21/06:  Disallow changes in current week
                        If (tgClfCntr(imLnRowNo - 1).ClfRec.sType <> "O") And (tgClfCntr(imLnRowNo - 1).ClfRec.sType <> "A") And (tgUrf(0).sChgCntr <> "I") Then
                            If (tgClfCntr(imLnRowNo - 1).iLineSchd) Then
                                If imLnSave(8, imLnRowNo) <> 0 Then '0= Daily
                                    If (llDate >= lmMonNowDate) And (llDate <= lmNowDate) Then
                                        llAllowedChgWkStartDate = mLnDate(imLnRowNo, smLnLLD, 0)
                                        If llAllowedChgWkStartDate < tgClfCntr(imLnRowNo - 1).lEndBilledDate Then
                                            llAllowedChgWkStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate
                                        End If
                                        gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndPrevPeriod
                                        If (llAllowedChgWkStartDate < llEndPrevPeriod) Then
                                            llAllowedChgWkStartDate = llEndPrevPeriod + 1
                                        End If
                                        If (llAllowedChgWkStartDate >= llDate) And (tgClfCntr(imLnRowNo - 1).iStatus > 0) And ((tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "F") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "I") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "A") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "N") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "P") Or (tgClfCntr(imLnRowNo - 1).ClfRec.sSchStatus = "M")) Then
                                            mGetCurrentWkNoSpots imLnRowNo, hmCHF, hmClf
                                            If tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots > 0 Then
                                                If llSpotCount < tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots Then
                                                    MsgBox "Cannot define spots less than " & Trim$(str$(tgClfCntr(imLnRowNo - 1).iMonThruToDaysSpots)) & " in current week as spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                                    ilChg = False
                                                    imPaintLnRowNo = -1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
        
                        '8/21/06:  Disallow prior weeks from being altered
                        ''Test if prior to todays date
                        If (llDate < lmMonNowDate) And ilChg And (tgUrf(0).sChgCntr <> "I") Then
                            imPaintLnRowNo = -1
                            If (tgClfCntr(imLnRowNo - 1).iStatus = 0) Or (Not tgClfCntr(imLnRowNo - 1).iLineSchd) Or (imLnRowNo = UBound(tgClfCntr) + 1) Or (lmWkSpots < llSpotCount) Then
                                If imPastMsgAsked = 0 Then
                                    ilRet = MsgBox("You are Changing a Past Week, Ok to Proceed", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                                    imPastMsgAsked = ilRet
                                    If ilRet = vbNo Then
                                        ilChg = False
                                    End If
                                Else
                                    If imPastMsgAsked = vbNo Then
                                        ilChg = False
                                    End If
                                End If
                            Else
                                MsgBox "Changing a Past Week not Permitted", vbOKOnly + vbExclamation, "Warning"
                                ilChg = False
                            End If
                        End If
                        
                        '6/9/12: Check MG in prior weeks
                        If (ilChg) And (lmWkSpots > llSpotCount) And (imLnSave(8, imLnRowNo) <> 0) Then
                            ilRet = mGetMGCount(imLnRowNo, hmCHF, hmClf, hmSdf, hmSmf, llDate, llDate + 6, ilBilledCount)
                            If (ilRet > 0) And (llSpotCount = 0) And (tgClfCntr(imLnRowNo - 1).ClfRec.sType = "O") Then
                                If (tgSpf.sVirtPkgCompute = "H") Then
                                    MsgBox "Change hidden lines as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                Else
                                    MsgBox "Cannot define spots as zero in week as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                End If
                                ilChg = False
                                imPaintLnRowNo = -1
                            ElseIf (ilRet > 0) And (tgClfCntr(imLnRowNo - 1).ClfRec.sType = "A") Then
                                MsgBox "Change hidden lines as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                ilChg = False
                                imPaintLnRowNo = -1
                            ElseIf (llSpotCount < ilRet) And (tgClfCntr(imLnRowNo - 1).ClfRec.sType <> "O") And (tgClfCntr(imLnRowNo - 1).ClfRec.sType <> "A") Then
                                MsgBox "Cannot define spots less than " & Trim$(str$(ilRet)) & " in week as MG/Outside spots already aired", vbOKOnly + vbExclamation, "Incomplete"
                                ilChg = False
                                imPaintLnRowNo = -1
                            End If
                        End If
                    
                    '3/14/13: end of change
                    End If
                End If
                
                If ilChg Then
                    imComputeReq = True
                    tgClfCntr(imLnRowNo - 1).iSPChg = True
                    igLnSPChg = True
                    If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                        imLnChg = True
                    End If
                    tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"
                    'mGetFlightDaysForDate imLnRowNo, llDate, imOvDays(), ilDays()
                    'mAddWeek imLnRowNo, llDate, llSpotCount, ilDays()
                    If tgClfCntr(imLnRowNo - 1).iCancel Then
                        'Remove cancel message
                        For ilDay = WKSPOTSINDEX To WKSPOTSINDEX + 12 Step 1
                            'gPaintArea Contract.pbcLine, tmLnCtrls(ilDay).fBoxX, tmLnCtrls(ilDay).fBoxY + (imLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) + 15, tmLnCtrls(ilDay).fBoxW - 15, tmLnCtrls(ilDay).fBoxH - 75, WHITE
                            gPaintArea Contract.pbcLine, tmLnCtrls(ilDay).fBoxX, tmLnCtrls(ilDay).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) + 15, tmLnCtrls(ilDay).fBoxW - 15, tmLnCtrls(ilDay).fBoxH - 75, WHITE
                        Next ilDay
                    End If
                    If imLnSave(8, imLnRowNo) = 0 Then   'Daily
                        If mAddStdPkgWk(hmPvf, imLnRowNo, llDate, "D", llSpotCount, ilDailySpots(), imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), hmCHF, hmClf, hmSdf, hmSmf, hmGhf, hmGsf) Then
                            If (igFreezeState <> 1) Or (smLineType <> "H") Then
                                imPaintLnRowNo = -1
                                Contract.pbcLine.Cls
                            End If
                        End If
                    Else
                        If mAddStdPkgWk(hmPvf, imLnRowNo, llDate, "W", llSpotCount, ilDailySpots(), imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), hmCHF, hmClf, hmSdf, hmSmf, hmGhf, hmGsf) Then
                            If (igFreezeState <> 1) Or (smLineType <> "H") Then
                                imPaintLnRowNo = -1
                                Contract.pbcLine.Cls
                            End If
                        End If
                    End If
                    'If mAddStdPkgWk(hmPvf, imLnRowNo, llDate, llSpotCount, imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf) Then
                    '    imPaintLnRowNo = -1
                    '    Contract.pbcLine.Cls
                    'End If
                    If imLnSave(8, imLnRowNo) = 0 Then   'Daily
                        mAddWeek imLastCffUsed, imLnRowNo, llDate, "D", llSpotCount, ilDailySpots()
                    Else
                        mAddWeek imLastCffUsed, imLnRowNo, llDate, "W", llSpotCount, imWkDays()
                    End If
                    If lmWkSpots = 0 Then
                        '4/8/11: Clear HideCBS flag as vehicle will not appear in copy if left as Y
                        tgClfCntr(imLnRowNo - 1).ClfRec.sHideCBS = "N"
                       'If (smLnSave(22, imLnRowNo) <> "") Or ((plcScreen.Caption <> "Order") And (Contract.plcDemoBar.Visible)) Then
                        If (smLnSave(22, imLnRowNo) <> "") Or (sgCntrScreen <> "Orders") Then
                            'Set Price for week
                            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                        'Set proposal always to rate card price- 1/29/98
                                        'Might want to change to test weeks around week being added, if they match rate card, then
                                        'use rate card.  If they don't, then use overall average
                                        'If (smLnSave(21, imLnRowNo) = "") And (plcScreen.Caption <> "Order") And (Contract.plcDemoBar.Visible) Then
                                        If (sgCntrScreen <> "Orders") Then  'And (Contract.plcDemoBar.Visible) Then
                                            'mSetFlightPriceForWk imLnRowNo, llDate, tgClfCntr(imLnRowNo - 1).sCurrentPrice, imOvDays(), tmRcf, imCntrVefCode, imDGSave()
                                            If tgClfCntr(imLnRowNo - 1).iPriceSet = 2 Then
                                               '10843 remove dead code
'                                                If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'                                                    mSetFlightPriceForWk imLnRowNo, llDate, "", imOvDays(), tmRcf, imCntrVefCode, imDGSave(), imLnTotalType
'                                                Else
                                                    mSetFlightPriceForWk imLnRowNo, llDate, tgClfCntr(imLnRowNo - 1).sCurrentPrice, imOvDays(), tmRcf, imCntrVefCode, imDGSave(), imLnTotalType
                                               ' End If
                                            Else
                                                'Get price for first week from rate card
                                                slStr = Contract.lbcPriceType(0).List(1)
                                                mSetFlightPriceForWk imLnRowNo, llDate, slStr, imOvDays(), tmRcf, imCntrVefCode, imDGSave(), imLnTotalType
                                                '12/1/04- retain by Rate Card
                                                If slStr = Contract.lbcPriceType(0).List(1) Then
                                                    tgClfCntr(imLnRowNo - 1).sCurrentPrice = slStr
                                                Else
                                                    tgClfCntr(imLnRowNo - 1).sCurrentPrice = ""
                                                End If
                                            End If
                                        Else
                                            tgCffCntr(ilCff).CffRec.sPriceType = smLnSave(22, imLnRowNo)
                                            tgCffCntr(ilCff).CffRec.lActPrice = gStrDecToLong(smLnSave(21, imLnRowNo), 2)
                                        End If
                                        Exit Do
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                        End If
                    ElseIf llSpotCount = 0 Then
                        If (smLnSave(22, imLnRowNo) <> "") Or ((sgCntrScreen <> "Orders") And (Contract.plcDemoBar.Visible)) Then
                            ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                            Do While ilCff <> -1
                                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                    If (llDate >= tgCffCntr(ilCff).lStartDate) And (llDate <= tgCffCntr(ilCff).lEndDate) Then
                                        tgCffCntr(ilCff).CffRec.sPriceType = ""
                                        tgCffCntr(ilCff).CffRec.lActPrice = 0
                                        Exit Do
                                    End If
                                End If
                                ilCff = tgCffCntr(ilCff).iNextCff
                            Loop
                        End If
                        If (tgClfCntr(imLnRowNo - 1).iOrigAnySpots > 0) Or (tgClfCntr(imLnRowNo - 1).iOrigAnySpots = -1) Then
                            If Not mFlightHaveSpots(imLnRowNo) Then
                                'Make CBS
                                tgClfCntr(imLnRowNo - 1).iCancel = True
                                ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate
                                llFlEndDate = llFlStartDate - 1
                                gPackDateLong llFlEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), tgCffCntr(ilCff).lStartDate    'Week Start date
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), tgCffCntr(ilCff).lEndDate    'Week Start date
                                smLnSave(14, imLnRowNo) = Format$(tgCffCntr(ilCff).lStartDate, "m/d/yy") 'slStartDate
                                smLnSave(15, imLnRowNo) = Format$(tgCffCntr(ilCff).lEndDate, "m/d/yy") 'slEndDate
                                If tgCffCntr(ilCff).iStatus = -1 Then
                                    tgCffCntr(ilCff).iStatus = 0    'Reset to new
                                ElseIf tgCffCntr(ilCff).iStatus = 2 Then
                                    tgCffCntr(ilCff).iStatus = 1    'Reset old
                                End If
                                Do While tgCffCntr(ilCff).iNextCff <> -1
                                    ilCff = tgCffCntr(ilCff).iNextCff
                                    If tgCffCntr(ilCff).iStatus = 0 Then
                                        tgCffCntr(ilCff).iStatus = -1   'New to not used
                                    ElseIf tgCffCntr(ilCff).iStatus = 1 Then
                                        tgCffCntr(ilCff).iStatus = 2    'Old to delete
                                    End If
                                Loop
                                'imLnSave(11, imLnRowNo) = 0
                                smLnSave(23, imLnRowNo) = "0"
                            End If
                        End If
                    End If
                    ilBypassLnTotal = False
                    If smLineType = "H" Then
                        If mGetPkLineType(imLnRowNo) = "E" Then
                            ilBypassLnTotal = True
                            ilSvLnRowNo = imLnRowNo
                            ilRet = 0
                            For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                                If (tgPackLn(ilLoop).sType = "E") Then
                                    ilIndex = tgPackLn(ilLoop).iIndex
                                    Do While ilIndex >= 0
                                        If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilSvLnRowNo - 1).ClfRec.iLine Then
                                            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                                If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilLoop).iLineNo Then
                                                    imLnRowNo = ilClf + 1
                                                    smLineType = mGetLineType(imLnRowNo)
                                                    'mLnTotals ilClf + 1, imLnTotalType, PRICEINDEX  'Set hidden line cost
                                                    For ilDay = 0 To 6 Step 1
                                                        ilOVDays(ilDay) = imLnSave(12 + ilDay, imLnRowNo)
                                                    Next ilDay
                                                    ilRet = mGetFlightSpotsDays(imLnRowNo, llDate, imOvDays(), llWkSpots, imWkDays(), slWDFlag, ilDailySpots())
                                                    If imLnSave(8, imLnRowNo) = 0 Then   'Daily
                                                        mAddWeek imLastCffUsed, imLnRowNo, llDate, "D", llSpotCount, ilDailySpots()
                                                    Else
                                                        mAddWeek imLastCffUsed, imLnRowNo, llDate, "W", llWkSpots + (llSpotCount - lmWkSpots), imWkDays()
                                                    End If
                                                    'imLnSave(11, imLnRowNo) = imLnSave(11, imLnRowNo) + (llSpotCount - lmWkSpots)
                                                    smLnSave(23, imLnRowNo) = Trim$(str$(Val(smLnSave(23, imLnRowNo)) + CLng(llSpotCount - lmWkSpots)))
                                                    'imLnSave(11, ilSvLnRowNo) = imLnSave(11, ilSvLnRowNo) + (llSpotCount - lmWkSpots)
                                                    smLnSave(23, ilSvLnRowNo) = Trim$(str$(Val(smLnSave(23, ilSvLnRowNo)) + CLng(llSpotCount - lmWkSpots)))
                                                    ilRet = 1
                                                    If (igFreezeState <> 1) Or (smLineType <> "H") Then
                                                        Contract.pbcLine.Cls
                                                    End If
                                                    imPaintLnRowNo = -1
                                                    Exit Do
                                                End If
                                            Next ilClf
                                        End If
                                        ilIndex = tgPackLn(ilIndex).iIndex
                                    Loop
                                    If ilRet > 0 Then
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            imLnRowNo = ilSvLnRowNo
                            smLineType = mGetLineType(imLnRowNo)
                            mCVehOptIndex imLnRowNo, imCntrVefCode, imVpfIndex, imTerminate
                        End If
                    Else
                        '10389 added visible demo bar
                        If ((smLineType = "O") Or (smLineType = "A")) And ((tgSpf.sVirtPkgCompute = "H") Or (Contract.plcDemoBar.Visible = True)) And (smEnabledLnPrice <> "") Then
                            'Dan 10389
                          ' If Contract.plcDemoBar.Visible Then
                            If gVisibleTabNeedsResearch() Then
                                ilSvLnRowNo = imLnRowNo
                                For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                                    If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilSvLnRowNo - 1).ClfRec.iLine Then
                                        ilIndex = tgPackLn(ilLoop).iIndex
                                        Do While ilIndex >= 0
                                            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                                imLnRowNo = ilClf + 1
                                                If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                                                    'mResetDemoBar False
                                                    tgClfCntr(imLnRowNo - 1).iResearch = False
                                                    If imDBTotals < 5 Then
                                                        'mBuildWkInfo True, True, False
                                                        mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
                                                    Else
                                                        mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilClf
                                            ilIndex = tgPackLn(ilIndex).iIndex
                                        Loop
                                        Exit For
                                    End If
                                Next ilLoop
                                imLnRowNo = ilSvLnRowNo
                            End If
                        End If
                    End If
                    Contract.mResetDemoBar ilBypassLnTotal
                    Contract.mCntrTotal
                End If
                Contract.pbcOVWD(0).Visible = False
                Contract.edcDGDropDown.Visible = False
                Contract.pbcDGPreempt.Visible = False
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
                If (igFreezeState <> 1) Or (smLineType <> "H") Then
                    Contract.pbcLine_Paint
                End If
                imPaintLnRowNo = ilPaintLnRowNo
            Else
                Contract.pbcOVWD(0).Visible = False
                Contract.edcDGDropDown.Visible = False
                Contract.pbcDGPreempt.Visible = False
                Contract.edcLNDropDown.Visible = False
                Contract.edcDBValue.Visible = False
                Contract.pbcLnDaily.Visible = False
                Contract.pbcDBLnDaily.Visible = False
            End If
            
            '5/19/12: Restore Override values
            If imLnSave(8, imLnRowNo) = 0 Then   'Daily
                For ilDay = 0 To 6 Step 1
                    'If (igFreezeState <> 1) Then
                        smLnSave(7 + ilDay, imLnRowNo) = imOvDailySpots(ilDay)
                    'Else
                    '    smLnSave(7 + ilBox - WKSPOTSINDEX, imLnRowNo) = Contract.edcDGDropDown.Text
                    'End If
                Next ilDay
            End If
            Screen.MousePointer = vbDefault
        Case PRICEINDEX 'Price
            Contract.lbcPriceType(0).Visible = False
            Contract.edcLNDropDown.Visible = False
            Contract.cmcLnDropDown.Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            ilChg = False
            If tgUrf(0).sChgPrices <> "V" Then
                'Price input is always to be Average Price- Jim 10/15/97
                'If imLnTotalType = 0 Then   'Show average
                'If imLnSave(11, imLnRowNo) > 0 Then
                If Val(smLnSave(23, imLnRowNo)) > 0 Then
                    'slStr = gDivStr(smLnSave(18, imLnRowNo), Trim$(Str$(imLnSave(11, imLnRowNo))))
                    'llAvg = gStrDecToLong(smLnSave(18, imLnRowNo), 2) \ imLnSave(11, imLnRowNo) '+ .5
                    If smLineType <> "E" Then
                       '10843 remove dead code
'                        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'                            slStr = smLnAvgPrice
'                        Else
                            'slStr = gLongToStrDec(gStrDecToLong(smLnSave(18, imLnRowNo), 2) \ imLnSave(11, imLnRowNo), 2)
                            slStr = gLongToStrDec(gStrDecToLong(smLnSave(18, imLnRowNo), 2) \ Val(smLnSave(23, imLnRowNo)), 2)
                        'End If
                    Else
                        slStr = smLnAvgPrice
                    End If
                Else
                    slStr = ""
                End If
                'If slStr <> Contract.edcLNDropDown.Text Then
                slStr = Trim$(slStr)
                slName = Trim$(Contract.edcLNDropDown.Text)
                If (slStr <> "") And (slName <> "") Then
                    If (Left$(slStr, 1) >= "0") And (Left$(slStr, 1) <= "9") And (Left$(slName, 1) >= "0") And (Left$(slName, 1) <= "9") Then
                        ilRet = gCompNumberStr(slStr, slName)
                    Else
                        If slStr <> slName Then
                            ilRet = 1
                        Else
                            ilRet = 0
                        End If
                    End If
                Else
                    If slStr <> slName Then
                        ilRet = 1
                    Else
                        ilRet = 0
                    End If
                End If
                If ilRet <> 0 Then
                    'Set all flights to this value
                    slPrice = Contract.edcLNDropDown.Text
                    ilChg = True
                Else
                    'Check if price not set- if so, set it
                    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            If tgCffCntr(ilCff).CffRec.sPriceType = "*" Then
                                slPrice = Contract.edcLNDropDown.Text
                                ilChg = True
                                Exit Do
                            End If
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                End If
            End If
            ilDatesBilled = False
            If (smEnabledLnPrice <> "") And (ilChg) Then
                gUnpackDateLong tgClfCntr(imLnRowNo - 1).ClfRec.iStartDate(0), tgClfCntr(imLnRowNo - 1).ClfRec.iStartDate(1), llDate
                If (tgClfCntr(imLnRowNo - 1).lEndBilledDate >= llDate) And (tgClfCntr(imLnRowNo - 1).lEndBilledDate > 0) And (llDate > 0) Then
                    ilDatesBilled = True
                End If
                'Jim- 6/21/98- don't ask question if price is zero
                If gStrDecToLong(smEnabledLnPrice, 2) <> 0 Then
                    'Jim- 10/31/03- don't ask price if the price is the same across all specified weeks
                    llActPrice = -1
                    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            If llActPrice = -1 Then
                                llActPrice = tgCffCntr(ilCff).CffRec.lActPrice
                            Else
                                If tgCffCntr(ilCff).CffRec.lActPrice <> llActPrice Then
                                    llActPrice = -2
                                    Exit Do
                                End If
                            End If
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                    If llActPrice = -2 Then
                       '10843 remove dead code
'                        If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'                            ilRet = MsgBox("Prices Varies, Change " & Contract.lbcQtPrice.List(imLnTotalType) & " to " & slPrice, vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
'                        Else
                            If Not ilDatesBilled Then
                                ilRet = MsgBox("Prices Varies, Change all Prices to " & slPrice, vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                            Else
                                'ilRet = MsgBox("Prices Varies across weeks, Price will only be changed in Unbilled Weeks (the average price will be average of billed and unbilled weeks), Set all unbilled weeks", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                                ilRet = Contract.mChgPriceQuestion(imLnRowNo, slPrice)
                            End If
                       ' End If
                        If ilRet = vbNo Then
                            ilChg = False
                        End If
                    Else
                        If ilDatesBilled Then
                            'ilRet = MsgBox("Price will only be changed in Unbilled Weeks (the average price will be average of billed and unbilled weeks), Set unbilled weeks", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                            ilRet = Contract.mChgPriceQuestion(imLnRowNo, slPrice)
                            If ilRet = vbNo Then
                                ilChg = False
                            End If
                        End If
                    End If
                Else
                    If ilDatesBilled Then
                        'ilRet = MsgBox("Price will only be changed in Unbilled Weeks (the average price will be average of billed and unbilled weeks), Set unbilled weeks", vbYesNo + vbDefaultButton2 + vbQuestion, "Warning")
                        ilRet = Contract.mChgPriceQuestion(imLnRowNo, slPrice)
                        If ilRet = vbNo Then
                            ilChg = False
                        End If
                    End If
                End If
            End If
            If ilChg Then
                If ilDatesBilled Then
                    mSplitFlightsIntoWeeks imLnRowNo
                End If
                slStr = slPrice
                'If (Left$(slStr, 1) >= "0") And (Left$(slStr, 1) <= "9") Then
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                'End If
                'gSetShow Contract.pbcLine, slStr, tmLnCtrls(PRICEINDEX)
                mSetPriceShow slStr, PRICEINDEX
                smLnShow(PRICEINDEX, imLnRowNo) = tmLnCtrls(ilBox).sShow
            End If
            If ilChg Then
                bmLnChgRegenResearch = True
                imComputeReq = True
                tgClfCntr(imLnRowNo - 1).iSPChg = True
                tgClfCntr(imLnRowNo - 1).sLineChgd = "Y"

                '11/30/04- sCurrentPrice set if lines created via Rate Bar but not just added into line area
                'the result was that the adding of weeks performed differently.
                'This price was retained so that the new weeks would have the same price as defined in the rate bar
                tgClfCntr(imLnRowNo - 1).sCurrentPrice = ""

                igLnSPChg = True
                'If (smLineType = "O") And (tgSpf.sVirtPkgCompute = "H") And (smEnabledLnPrice <> "") Then
                If ((smLineType = "O") Or (smLineType = "A")) And (tgSpf.sVirtPkgCompute = "H") And (smEnabledLnPrice <> "") Then
                    '11/30/04- sCurrentPrice set if lines created via Rate Bar but not just added into line area
                    'the result was that the adding of weeks performed differently.
                    'This price was retained so that the new weeks would have the same price as defined in the rate bar
                    ilSvLnRowNo = imLnRowNo
                    For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                        If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilSvLnRowNo - 1).ClfRec.iLine Then
                            ilIndex = tgPackLn(ilLoop).iIndex
                            Do While ilIndex >= 0
                                For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                    imLnRowNo = ilClf + 1
                                    If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                                        tgClfCntr(imLnRowNo - 1).sCurrentPrice = ""
                                        Exit For
                                    End If
                                Next ilClf
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                            Exit For
                        End If
                    Next ilLoop
                    imLnRowNo = ilSvLnRowNo
                   '10843 remove dead code
'                    If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'                        llFlStartDate = lmHSStartDate + 7 * (imHSCurWkNo - 1)
'                        slStr = Format$(llFlStartDate, "m/d/yy")
'                        gObtainMonthYear 0, slStr, ilCff, ilClf
'                        slStr = Trim$(str$(3 * ((ilCff - 1) \ 3) + 1)) & "/15/" & Trim$(str$(ilClf))
'                        llFlStartDate = gDateValue(gObtainStartStd(slStr))
'                        llFlEndDate = llFlStartDate + 7 * 12 - 1
'                        ilCff = (ilCff - 1) \ 3 + 1
'                        ilRet = ilCff
'                        Do While ilRet = ilCff
'                            llFlEndDate = llFlEndDate + 7
'                            slStr = Format(llFlEndDate, "m/d/yy")
'                            gObtainMonthYear 0, slStr, ilCff, ilClf
'                            ilCff = (ilCff - 1) \ 3 + 1
'                        Loop
'                        llFlEndDate = llFlEndDate - 7
'                    Else
                        '3/17/11: Disallow dates in price to be changed
                        'llFlStartDate = 0
                        'llFlEndDate = 0
                        If tgClfCntr(imLnRowNo - 1).iChgBilledPrice = True Then
                            gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llFlStartDate
                            llFlStartDate = llFlStartDate + 1
                            llFlEndDate = llFlStartDate + 2000
                        Else
                            If tgClfCntr(imLnRowNo - 1).lEndBilledDate <> 0 Then
                                llFlStartDate = tgClfCntr(imLnRowNo - 1).lEndBilledDate + 1
                                llFlEndDate = llFlStartDate + 2000
                            Else
                                llFlStartDate = 0
                                llFlEndDate = 0
                            End If
                        End If
                   ' End If
                    'If (Left$(slPrice, 1) >= "0") And (Left$(slPrice, 1) <= "9") Then
                    If gStrDecToLong(slPrice, 2) > 0 Then
                        'If (Left$(smEnabledLnPrice, 1) >= "0") And (Left$(smEnabledLnPrice, 1) <= "9") Then
                        If gStrDecToLong(smEnabledLnPrice, 2) > 0 Then
                            ilRet = mAdjStdPkPrice(imLnRowNo, llFlStartDate, llFlEndDate, gStrDecToLong(slPrice, 2), gStrDecToLong(smEnabledLnPrice, 2), imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode())
                        Else
                            ilRet = mAdjStdPkPrice(imLnRowNo, llFlStartDate, llFlEndDate, gStrDecToLong(slPrice, 2), 0, imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode())
                        End If
                    Else
                        'If (Left$(smEnabledLnPrice, 1) >= "0") And (Left$(smEnabledLnPrice, 1) <= "9") Then
                        If gStrDecToLong(smEnabledLnPrice, 2) > 0 Then
                            ilRet = mAdjStdPkPrice(imLnRowNo, llFlStartDate, llFlEndDate, 0, gStrDecToLong(smEnabledLnPrice, 2), imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode())
                        Else
                            ilRet = mAdjStdPkPrice(imLnRowNo, llFlStartDate, llFlEndDate, 0, 0, imLastCffUsed, tmRcf, imDGSave(), imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode())
                        End If
                    End If
                    If Contract.plcDemoBar.Visible Then
                        ilSvLnRowNo = imLnRowNo
                        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                            If tgPackLn(ilLoop).iLineNo = tgClfCntr(ilSvLnRowNo - 1).ClfRec.iLine Then
                                ilIndex = tgPackLn(ilLoop).iIndex
                                Do While ilIndex >= 0
                                    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                                        imLnRowNo = ilClf + 1
                                        If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilClf).ClfRec.iLine Then
                                            'mResetDemoBar False
                                            tgClfCntr(imLnRowNo - 1).iResearch = False
                                            If imDBTotals < 5 Then
                                                'mBuildWkInfo True, True, False
                                                mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
                                            Else
                                                mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
                                            End If
                                            Exit For
                                        End If
                                    Next ilClf
                                    ilIndex = tgPackLn(ilIndex).iIndex
                                Loop
                                Exit For
                            End If
                        Next ilLoop
                        imLnRowNo = ilSvLnRowNo
                    End If
                Else
                    Contract.mSetFlightPrice slPrice, False
                End If
                imLnChg = True
                If (igFreezeState = 1) And (smLineType = "H") Then
                    Contract.mResetDemoBar True
                Else
                    Contract.mResetDemoBar False
                End If
                
                mSetAcqOrSpotPrice "A", imLnRowNo, slPrice, PRICEINDEX
                
                Screen.MousePointer = vbDefault
            End If
            'Add line if this is the last line
            If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                bmLnChgRegenResearch = True
                If tgClfCntr(imLnRowNo - 1).iStatus = -1 Then
                    tgClfCntr(imLnRowNo - 1).iStatus = 0
                    imLnChg = True
                    imComputeReq = True
                End If
                ReDim Preserve tgClfCntr(0 To imLnRowNo) As CLFLIST
                ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                For ilLoop = LBound(imLnSave, FIELDBOUNDINDEX) To UBound(imLnSave, FIELDBOUNDINDEX) Step 1
                    imLnSave(ilLoop, UBound(tgClfCntr) + 1) = -1
                Next ilLoop
                Contract.mInitNewLine
                smLnSave(1, imLnRowNo + 1) = "" 'Vehicle list index
                Contract.mCntrTotal 'Obtain total for new line, placed back into service 7/15/97
                mSetSchStatus smOrigStatus
                ilChg = True    'Since setting price with each week, then this will force
                                'call to mSetPriDemoResearch
            Else
                If ilChg Then   'And tgUrf(0).sUseComputeCMC <> "I" Then
                    Contract.mCntrTotal 'Obtain total for new line, placed back into service 7/15/97
                    mSetSchStatus smOrigStatus
                End If
            End If
            If tgClfCntr(imLnRowNo - 1).bLnPriceSet = False Then 'new line
                bmLnChgRegenResearch = True
                tgClfCntr(imLnRowNo - 1).bLnPriceSet = True
            End If
            
            'Test if Standard Package- if so, determine it lines should be generated
            imInChkForStdPkg = True
            If Contract.mChkForStdPkg(imLnRowNo) = False Then
                imInChkForStdPkg = False
                Contract.mClearCtrlFields
                Contract.mTerminate
                Exit Sub
            End If
            imInChkForStdPkg = False
            If (sgCntrScreen = "Orders") Or ((Asc(tgSpf.sOptionFields) And OFRESEARCH) <> OFRESEARCH) Then
                Contract.lacAudTotals.Visible = False
            Else
                If (tgSpf.sGUsePropSys = "Y") Then
                    Contract.lacAudTotals.Visible = True
                Else
                    Contract.lacAudTotals.Visible = False
                End If
            End If
            If ilChg Then
                If tgSpf.sGUsePropSys = "Y" Then
                    'mSetPriDemoResearch hmDrf, hmMnf, imStep1Index
                End If
            End If
        Case LINECOMMINDEX
            imInLnComment = False
            If Contract.plcComment.Visible Then
                Contract.plcComment.Visible = False  'Set visibility
                If (imUpdateAllowed) And (Not igRepNetChgAllowed) Then
                    smLnComment(imLnRowNo) = gRemoveLeadingCrCl(Contract.edcComment.Text)
                End If
            End If
            imInCommentBranch = False
            Contract.edcLNDropDown.Visible = False
            Contract.cmcLnDropDown.Visible = False
            Contract.lbcComment(1).Visible = False
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                '12/28/18: Make work similar to header, click in field a second time caused the pop-up information to disappear
                'If Contract.pbcLineArrow.Visible Then
                '    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                'Else
                    Contract.pbcClickFocus.SetFocus
                'End If
                Exit Sub
            End If
            '10843 remove dead code
'            If tgSpf.sCLnStdQt = "Y" Then
'                smLnComment(imLnRowNo) = gRemoveLeadingCrCl(Contract.edcComment.Text)
'                slStr = Left$(Trim$(smLnComment(imLnRowNo)), 1) '20)
'            Else
                If Contract.lbcComment(1).ListIndex < 0 Then
                    smLnComment(imLnRowNo) = ""
                    slStr = ""
                ElseIf Contract.lbcComment(1).ListIndex = 0 Then
                    smLnComment(imLnRowNo) = ""
                    slStr = ""  '"[None]"
                Else
                    smLnComment(imLnRowNo) = gRemoveLeadingCrCl(Contract.edcComment.Text)
                    slStr = Left$(Trim$(smLnComment(imLnRowNo)), 1) '20)
                End If
            'End If
            If Trim$(slStr) <> "" Then
                For ilLoop = 0 To 4 Step 1
                    If Contract.ckcCommOn(ilLoop).Value = vbChecked Then
                        imLnCommShowOn(ilLoop, imLnRowNo) = 0
                    Else
                        imLnCommShowOn(ilLoop, imLnRowNo) = 1
                    End If
                Next ilLoop
            End If
            'gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilBoxNo)
            'smLnShow(LINECOMMINDEX, imLnRowNo) = tmLnCtrls(ilBoxNo).sShow
            smLnShow(LINECOMMINDEX, imLnRowNo) = slStr
            If smLnComment(imLnRowNo) <> smLnInitComment(imLnRowNo) Then
                imLnChg = True
            End If
    End Select
    'mPaintLnBkgd WHITE
    If (igFreezeState <> 1) Or (smLineType <> "H") Then
        imCurrHightLightLnRowNo = -1
        '9/15/18: Fixed dates above week not displayed
        imPaintLnRowNo = -1
        Contract.pbcLine_Paint
        imPrevHightLightLnRowNo = -1
    Else
        Contract.pbcLine_Paint
    End If
    ''12/5/19: clear value so that it does Not show-up in another field
    'Contract.edcLNDropDown.Text = ""
    'Contract.edcLNDropDown.Visible = False
    Contract.mSetCommands
End Sub

Public Sub mLnEnableBoxSub(ilBoxNo As Integer, LINECOMMINDEX As Integer, PRICEINDEX As Integer, OV1INDEX As Integer, LENINDEX As Integer, OV2INDEX As Integer, VEHINDEX As Integer, DPINDEX As Integer, WKSPOTSINDEX As Integer)

'
'   mLnEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    'Dim ilLen As Integer
    Dim slPrice As String
    Dim llDate As Long
    'Dim llSpotPrice As Long
    Dim ilLineNo As Integer
    Dim slWDFlag As String
    Dim ilDay As Integer
    ReDim ilDailySpots(0 To 6) As Integer
    Dim ilIgnoreRightMove As Integer
    
    If imTerminate Then
        Exit Sub
    End If
    If ilBoxNo < imLBCTCtrls Or ilBoxNo > UBound(tmLnCtrls) Then
        Exit Sub
    End If
    'Disallow new lines in on Research tab since columns replaced
    'If (tgClfCntr(imLnRowNo - 1).iFirstCff < 0) And (Contract.tscLine.SelectedItem.Index = 10) Then
    If (tgClfCntr(imLnRowNo - 1).iFirstCff < 0) And (Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH)) Then
        'Vehicle->smLnSave(1, imLnRowNo); Non-Package Daypart->imLnSave(1, imLnRowNo); Package Daypart->mLnSave(20, imLnRowNo); Len->smLnSave(16, imLnRowNo)
        imLnRowNo = -1
        Contract.pbcClickFocus.SetFocus
        Exit Sub
    End If
    'If (ilBoxNo = VEHINDEX) And (Contract.tscLine.SelectedItem.Index = 10) Then
    If (ilBoxNo = VEHINDEX) And (Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH)) Then
        imLnRowNo = -1
        Contract.pbcClickFocus.SetFocus
        Exit Sub
    End If
    '12/16/14: Add Hide/Show Hidden Lines
    mBuildMapLinePaint
    imMoveLnRowNo = mFindPaintLnRowNo(imLnRowNo)
    
    If (imMoveLnRowNo < Contract.vbcLines.Value + 1) Or (imMoveLnRowNo >= Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 2) Then
        Contract.mLnSetShow ilBoxNo
        Contract.pbcLineArrow.Visible = False
        Contract.laclnFrame.Visible = False
        Exit Sub
    End If
    Contract.mSetArrowFrame
    If (ilBoxNo < WKSPOTSINDEX) Or (ilBoxNo > WKSPOTSINDEX + 12) Then
        imDBBoxNo = -1
    End If
    smAlterHidden = "Y"
    smLineType = mGetLineType(imLnRowNo)
    If smLineType = "H" Then
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                'If tgPackLn(ilLoop).iAlter <> True Then
                smAlterHidden = tgPackLn(ilLoop).sAlter
                '10347 added 'R' and flipped logic to make it easier to read- if "C" or "R" and these box are chosen, don't block.
                If Not ((smAlterHidden = "C" And (ilBoxNo = LINECOMMINDEX Or ilBoxNo = OV1INDEX)) Or (smAlterHidden = "R" And (ilBoxNo = LINECOMMINDEX Or ilBoxNo = PRICEINDEX))) Then
                    If smAlterHidden <> "Y" Then
                        mBlockHiddenLineWriting
                        Exit Sub
                    End If
                End If
                Exit For
'                If (tgPackLn(ilLoop).sAlter <> "C") Or ((ilBoxNo <> LINECOMMINDEX) And (ilBoxNo <> OV1INDEX)) Then
'                    If tgPackLn(ilLoop).sAlter <> "Y" Then
'                        'If Contract.pbcLineArrow.Visible Then
'                        '    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
'                        'Else
'                            imCurrHightLightLnRowNo = -1
'                            Contract.pbcLine_Paint
'                            imPrevHightLightLnRowNo = -1
'                            MsgBox "Can't Alter Hidden Line", vbOKOnly + vbExclamation, "Warning"
'                            imLnRowNo = -1
'                            Contract.pbcClickFocus.SetFocus
'                        'End If
'                        ilBoxNo = -1
'                        Exit Sub
'                    End If
'                End If
'                Exit For
            End If
        Next ilLoop
    End If
    '2/21/17: Add test if allowed to alter package name
    If (ilBoxNo = VEHINDEX) And ((smLineType = "A") Or (smLineType = "O") Or (smLineType = "E")) Then
        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
            If tgPackLn(ilLoop).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                If tgPackLn(ilLoop).sAlterName <> "Y" Then
                    'If Contract.pbcLineArrow.Visible Then
                    '    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                    'Else
                        imCurrHightLightLnRowNo = -1
                        Contract.pbcLine_Paint
                        imPrevHightLightLnRowNo = -1
                        MsgBox "Can't Alter Package Name", vbOKOnly + vbExclamation, "Warning"
                        imLnRowNo = -1
                        Contract.pbcClickFocus.SetFocus
                    'End If
                    ilBoxNo = -1
                    Exit Sub
                End If
                Exit For
            End If
        Next ilLoop
    End If

    If tgClfCntr(imLnRowNo - 1).iFirstCff < 0 Then
        tgClfCntr(imLnRowNo - 1).iPriceSet = 1
    Else
        If tgClfCntr(imLnRowNo - 1).iPriceSet <> 1 Then
            tgClfCntr(imLnRowNo - 1).iPriceSet = 2
        End If
    End If
    tgSpf.sAllowDailyBuys = smSvAllowDailyWeekly
    If mIsLnGame(imLnRowNo) Then
        tgSpf.sAllowDailyBuys = "N"
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHINDEX 'Vehicle
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            If (tgClfCntr(imLnRowNo - 1).iStatus = -1) And (tgClfCntr(imLnRowNo - 1).ClfRec.iLine <= 0) Then
                ilLineNo = 0
                For ilLoop = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                    If (tgClfCntr(ilLoop).ClfRec.iLine > ilLineNo) And (tgClfCntr(ilLoop).iStatus >= 0) Then
                        ilLineNo = tgClfCntr(ilLoop).ClfRec.iLine
                    End If
                Next ilLoop
                For ilLoop = LBound(tgDelClf) To UBound(tgDelClf) - 1 Step 1
                    If (tgDelClf(ilLoop).ClfRec.iLine > ilLineNo) And (tgDelClf(ilLoop).iStatus >= 0) Then
                        ilLineNo = tgDelClf(ilLoop).ClfRec.iLine
                    End If
                Next ilLoop
                tgClfCntr(imLnRowNo - 1).ClfRec.iLine = ilLineNo + 1
                tgClfCntr(imLnRowNo - 1).bShowLine = True
                Contract.pbcLine_Paint
            End If
            'gShowHelpMess tmClfHelp(), CLFVEHICLE
            If (imLnRowNo = 1) And (Trim$(smLnSave(1, imLnRowNo)) = "") And imVefDormantAdd Then
                smVehicleCodeTag = ""
                imVefDormantAdd = False
            End If
            Contract.mVehPop False
            If imTerminate Then
                Exit Sub
            End If
            Contract.lbcLnVehicle(igTabMapIndex).height = gListBoxHeight(Contract.lbcLnVehicle(igTabMapIndex).ListCount, 5)
            Contract.lbcPkVehicle(igTabMapIndex).height = gListBoxHeight(Contract.lbcPkVehicle(igTabMapIndex).ListCount, 5)
            Contract.edcLNDropDown.Width = tmLnCtrls(VEHINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                Contract.edcLNDropDown.MaxLength = tgSpf.iVehLen
            Else
                Contract.edcLNDropDown.MaxLength = 20
            End If
            '12/16/14: Add Hide/Show Hidden Lines
            'gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(VEHINDEX).fBoxX, tmLnCtrls(VEHINDEX).fBoxY + (imLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(VEHINDEX).fBoxX, tmLnCtrls(VEHINDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            imChgMode = True
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                    Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                    Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(Contract.lbcLnVehicle(igTabMapIndex).ListIndex)
                Else
                    If imLnRowNo > 1 Then
                        gFindMatch smLnSave(1, imLnRowNo - 1), 0, Contract.lbcLnVehicle(igTabMapIndex)
                        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                            Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                            imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                            Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(Contract.lbcLnVehicle(igTabMapIndex).ListIndex)
                        Else
                            slNameCode = Contract.lbcRateCard.List(Contract.lbcRateCard.ListIndex)
                            ilRet = gParseItem(slNameCode, 2, "/", slName)
                            If ilRet <> CP_MSG_NONE Then
                                slName = sgUserDefVehicleName
                            End If
                            gFindMatch slName, 0, Contract.lbcLnVehicle(igTabMapIndex)
                            If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                                Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                                imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                                Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(Contract.lbcLnVehicle(igTabMapIndex).ListIndex)
                            Else
                                If Contract.lbcLnVehicle(igTabMapIndex).ListCount > 0 Then
                                    Contract.lbcLnVehicle(igTabMapIndex).ListIndex = 0
                                Else
                                    Contract.lbcLnVehicle(igTabMapIndex).ListIndex = -1
                                End If
                                imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                                Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(0)
                            End If
                        End If
                    Else
                        slNameCode = Contract.lbcRateCard.List(Contract.lbcRateCard.ListIndex)
                        ilRet = gParseItem(slNameCode, 2, "/", slName)
                        If ilRet <> CP_MSG_NONE Then
                            slName = sgUserDefVehicleName
                        End If
                        If (slName = "[All Vehicles]") And (sgUserDefVehicleName <> "") Then
                            slName = sgUserDefVehicleName
                        End If
                        gFindMatch slName, 0, Contract.lbcLnVehicle(igTabMapIndex)
                        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                            Contract.lbcLnVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnVehicle(igTabMapIndex))
                            imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                            Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(Contract.lbcLnVehicle(igTabMapIndex).ListIndex)
                        Else
                            If Contract.lbcLnVehicle(igTabMapIndex).ListCount > 0 Then
                                Contract.lbcLnVehicle(igTabMapIndex).ListIndex = 0
                            Else
                                Contract.lbcLnVehicle(igTabMapIndex).ListIndex = -1
                            End If
                            imComboBoxIndex = Contract.lbcLnVehicle(igTabMapIndex).ListIndex
                            Contract.edcLNDropDown.Text = Contract.lbcLnVehicle(igTabMapIndex).List(0)
                        End If
                    End If
                End If
            Else
                gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcPkVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcPkVehicle(igTabMapIndex)) >= 0 Then
                    Contract.lbcPkVehicle(igTabMapIndex).ListIndex = gLastFound(Contract.lbcPkVehicle(igTabMapIndex))
                    Contract.edcLNDropDown.Text = Contract.lbcPkVehicle(igTabMapIndex).List(Contract.lbcPkVehicle(igTabMapIndex).ListIndex)
                Else
                    Contract.lbcPkVehicle(igTabMapIndex).ListIndex = -1
                    Contract.edcLNDropDown.Text = smLnSave(1, imLnRowNo)
                End If
            End If
            imChgMode = False
            If imLnRowNo - Contract.vbcLines.Value <= Contract.vbcLines.LargeChange \ 2 Then
                Contract.lbcLnVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
                Contract.lbcPkVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcLnVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnVehicle(igTabMapIndex).height
                Contract.lbcPkVehicle(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcPkVehicle(igTabMapIndex).height
            End If
            If Contract.lbcLnVehicle(igTabMapIndex).Width < Contract.edcLNDropDown.Width Then
                Contract.lbcLnVehicle(igTabMapIndex).Width = Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width
                Contract.lbcPkVehicle(igTabMapIndex).Width = Contract.lbcLnVehicle(igTabMapIndex).Width
            End If
            'Width set in the paint event
            imINPBCPaint = False
            'Moved to paint event
            'pbcLbcVehicle.Move Contract.lbcLnVehicle(igTabMapIndex).Left + 15, Contract.lbcLnVehicle(igTabMapIndex).Top + 15, pbcLbcVehicle.Width - 30, Contract.lbcLnVehicle(igTabMapIndex).Height - 30 '2115, 1560
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case DPINDEX 'Daypart index
            ''If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
            'If (igRepNetChgAllowed) Or ((Not imUpdateAllowed) And (Contract.tscLine.SelectedItem.Index = 10)) Then
            If (igRepNetChgAllowed) Or ((Not imUpdateAllowed) And (Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH))) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            If (smLineType = "O") Or (smLineType = "A") Or (smLineType = "E") Then
                'If Contract.tscLine.SelectedItem.Index = 10 Then
                If Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH) Then
                    If Contract.pbcLineArrow.Visible Then
                        Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                    Else
                        Contract.pbcClickFocus.SetFocus
                    End If
                    Exit Sub
                End If
            End If
            'gShowHelpMess tmClfHelp(), CLFPROGRAM
            Contract.edcLNDropDown.MaxLength = 0
            Contract.mPrgPop imLnRowNo, "L"
            If imTerminate Then
                Exit Sub
            End If
            Contract.lbcLnProgram(igTabMapIndex).height = gListBoxHeight(Contract.lbcLnProgram(igTabMapIndex).ListCount, 5)
            Contract.lbcPkProgram(igTabMapIndex).height = gListBoxHeight(Contract.lbcPkProgram(igTabMapIndex).ListCount, 5)
            Contract.edcLNDropDown.Width = tmLnCtrls(DPINDEX).fBoxW
            '12/16/14: Add Hide/Show Hidden Lines
            'gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(DPINDEX).fBoxX, tmLnCtrls(DPINDEX).fBoxY + (imLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(DPINDEX).fBoxX, tmLnCtrls(DPINDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            imChgMode = True
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                'If Contract.tscLine.SelectedItem.Index <> 10 Then
                If Contract.tscLine.SelectedItem.Index <> imTabMap(TABRESEARCH) Then
                    Contract.lbcLnProgram(igTabMapIndex).ListIndex = -1
                    imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                    Contract.edcLNDropDown.Text = ""
                    For ilLoop = 0 To Contract.lbcPrg(igTabMapIndex).ListCount - 1 Step 1
                        If imLnSave(1, imLnRowNo) = Contract.lbcPrg(igTabMapIndex).ItemData(ilLoop) Then
                            Contract.lbcLnProgram(igTabMapIndex).ListIndex = ilLoop
                            imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                            Contract.edcLNDropDown.Text = Contract.lbcLnProgram(igTabMapIndex).List(ilLoop)
                            Exit For
                        End If
                    Next ilLoop
                    'If no matching program, select next program after one selected if same Vehicle
                    If (Contract.lbcLnProgram(igTabMapIndex).ListIndex = -1) And (imLnRowNo > 1) Then
                        If smLnSave(1, imLnRowNo) = smLnSave(1, imLnRowNo - 1) Then
                            For ilLoop = 0 To Contract.lbcPrg(igTabMapIndex).ListCount - 2 Step 1
                                If imLnSave(1, imLnRowNo - 1) = Contract.lbcPrg(igTabMapIndex).ItemData(ilLoop) Then
                                    Contract.lbcLnProgram(igTabMapIndex).ListIndex = ilLoop + 1
                                    imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                                    Contract.edcLNDropDown.Text = Contract.lbcLnProgram(igTabMapIndex).List(ilLoop + 1)
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    End If
                    If (Contract.lbcLnProgram(igTabMapIndex).ListIndex = -1) And (Contract.lbcLnProgram(igTabMapIndex).ListCount > 0) Then
                        Contract.lbcLnProgram(igTabMapIndex).ListIndex = 0    'default to first program
                        imComboBoxIndex = Contract.lbcLnProgram(igTabMapIndex).ListIndex
                        Contract.edcLNDropDown.Text = Contract.lbcLnProgram(igTabMapIndex).List(0)
                    End If
                End If
            Else
                gFindMatch smLnSave(20, imLnRowNo), 0, Contract.lbcPkProgram(igTabMapIndex)
                If gLastFound(Contract.lbcPkProgram(igTabMapIndex)) >= 0 Then
                    Contract.lbcPkProgram(igTabMapIndex).ListIndex = gLastFound(Contract.lbcPkProgram(igTabMapIndex))
                    imComboBoxIndex = Contract.lbcPkProgram(igTabMapIndex).ListIndex
                    Contract.edcLNDropDown.Text = Contract.lbcPkProgram(igTabMapIndex).List(Contract.lbcPkProgram(igTabMapIndex).ListIndex)
                Else
                    Contract.lbcPkProgram(igTabMapIndex).ListIndex = -1
                    imComboBoxIndex = Contract.lbcPkProgram(igTabMapIndex).ListIndex
                    Contract.edcLNDropDown.Text = smLnSave(20, imLnRowNo)
                End If
            End If
            imChgMode = False
            If imMoveLnRowNo - Contract.vbcLines.Value <= Contract.vbcLines.LargeChange \ 2 Then
                Contract.lbcLnProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
                Contract.lbcPkProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcLnProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLnProgram(igTabMapIndex).height
                Contract.lbcPkProgram(igTabMapIndex).Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcPkProgram(igTabMapIndex).height
            End If
            If (imUpdateAllowed) Then
                Contract.edcLNDropDown.SelStart = 0
                Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
                Contract.edcLNDropDown.Visible = True
                Contract.cmcLnDropDown.Visible = True
                If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                    Contract.lbcLnProgram(igTabMapIndex).Visible = True
                Else
                    Contract.lbcPkProgram(igTabMapIndex).Visible = True
                End If
                Contract.edcLNDropDown.SetFocus
            Else
                Contract.edcLNDropDown.Visible = True
            End If
        Case OV1INDEX 'BB index
            imOVType = 1
            Contract.mSetTmOVCtrl
            smLineTypeOv = mGetLineTypeForOV(imLnRowNo)
            If (smLineTypeOv <> "P") Then
                mDPAllowedDays imLnSave(1, imLnRowNo), imAllowedDays()
            Else
                For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If StrComp(smLnSave(20, imLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                        mDPAllowedDays tgMRdf(ilLoop).iCode, imAllowedDays()
                        Exit For
                    End If
                Next ilLoop
            End If
            If (imLnSave(8, imLnRowNo) = -1) Then
                imDGSave(2) = 1 'Weekly
                slStr = "Weekly"
            Else
                If imLnSave(8, imLnRowNo) = 0 Then
                    imDGSave(2) = 0 'Daily
                    slStr = "Daily"
                Else
                    imDGSave(2) = 1 'Weekly
                    slStr = "Weekly"
                End If
            End If
            If (tgSpf.sAllowDailyBuys = "Y") Then
                gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVWDINDEX)
            End If
            'set imOvDays- this is set in Contract.pbcLine and in pbcLnTab, pbcLnSTab
            For ilLoop = OVDAYSINDEX To OVDAYSINDEX + 6 Step 1
                If imDGSave(2) = 0 Then
                    smDGSave(3 + ilLoop - OVDAYSINDEX) = smLnSave(7 + ilLoop - OVDAYSINDEX, imLnRowNo)
                    slStr = smLnSave(7 + ilLoop - OVDAYSINDEX, imLnRowNo)
                    imOrigPkDailySpots(ilLoop - OVDAYSINDEX) = Val(slStr)
                    If imOvDays(ilLoop - OVDAYSINDEX) = 2 Then
                        slStr = "Mix"
                    End If
                Else
                    If imOvDays(ilLoop - OVDAYSINDEX) = 0 Then
                        slStr = "Yes"
                    ElseIf imOvDays(ilLoop - OVDAYSINDEX) = 1 Then
                        slStr = "No"
                    ElseIf imOvDays(ilLoop - OVDAYSINDEX) = 2 Then
                        slStr = "Mix"
                    Else
                        slStr = ""
                    End If
                End If
                gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ilLoop)
            Next ilLoop
            For ilLoop = LBound(imOvDays) To UBound(imOvDays) Step 1
                imSvOvDays(ilLoop) = imOvDays(ilLoop)
            Next ilLoop
            If imLnSave(4, imLnRowNo) = 1 Then  'Time override
                smDGSave(13) = smLnSave(2, imLnRowNo)
                smDGSave(14) = smLnSave(3, imLnRowNo)
            Else
                smDGSave(13) = ""
                smDGSave(14) = ""
            End If
            slStr = smDGSave(13)
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVSTARTTIMEINDEX)
            slStr = smDGSave(14)
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVENDTIMEINDEX)
            If imLnSave(6, imLnRowNo) < 0 Then  'Preemptible (later might want site question default preempt- default also set in mClearPB)
                imLnSave(6, imLnRowNo) = 1
                If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Then
                Else
                    imLnChg = True
                    imComputeReq = True
                End If
            End If
            '4/15/19: Default obtained from vehcile option
            'If (imLnSave(31, imLnRowNo) < 0) And (tmOV1Ctrls(OVCOPYINDEX).iReq <> True) Then  'Preemptible (later might want site question default preempt- default also set in mClearPB)
            '    imLnSave(31, imLnRowNo) = 1
            '    If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Then
            '    Else
            '        imLnChg = True
            '        imComputeReq = True
            '    End If
            'End If
            imDGSave(7) = imLnSave(31, imLnRowNo)
            If imDGSave(7) = 0 Then
                slStr = "LC"    '"Live Cmml"
            ElseIf imDGSave(7) = 2 Then
                slStr = "LP"    '"Live Promo"
            ElseIf imDGSave(7) = 3 Then
                slStr = "RP"    '"Rec Promo"
            ElseIf imDGSave(7) = 4 Then
                slStr = "PC"    '"Pre-recorded Live Cmml"
            ElseIf imDGSave(7) = 5 Then
                slStr = "PP"    '"Pre-recorded live promo"
            ElseIf imDGSave(7) = 1 Then
                slStr = "RC"    '"Rec Cmml"
            Else
                slStr = mObtainAudioTypeDefault(smLnSave(1, imLnRowNo))
                If slStr = "L" Then
                    slStr = "LC"
                    imDGSave(7) = 0
                ElseIf slStr = "M" Then
                    slStr = "LP"
                    imDGSave(7) = 2
                ElseIf slStr = "S" Then
                    slStr = "RP"
                    imDGSave(7) = 3
                ElseIf slStr = "P" Then
                    slStr = "PC"
                    imDGSave(7) = 4
                ElseIf slStr = "Q" Then
                    slStr = "PP"
                    imDGSave(7) = 5
                Else
                    slStr = "RC"
                    imDGSave(7) = 1
                End If
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVCOPYINDEX)
            imDGSave(3) = imLnSave(6, imLnRowNo)
            If imDGSave(3) = 0 Then
                slStr = "Yes"
            Else
                slStr = "No"
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVHIGHINDEX)
            'gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVPREEMPTINDEX)
            imDGSave(4) = imLnSave(19, imLnRowNo)
            If imDGSave(4) > 0 Then
                slStr = Trim$(str$(imDGSave(4)))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVRANKINDEX)
            imDGSave(5) = imLnSave(20, imLnRowNo)
            If imDGSave(5) > 0 Then
                slStr = Trim$(str$(imDGSave(5)))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVBBOPENINDEX)
            imDGSave(6) = imLnSave(21, imLnRowNo)
            If imDGSave(6) > 0 Then
                slStr = Trim$(str$(imDGSave(6)))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(OVBBCLOSEINDEX)
            
            smDGSave(15) = smLnSave(29, imLnRowNo)
            slStr = smDGSave(15)
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ACT1LINEUPINDEX)
            imDGSave(8) = imLnSave(32, imLnRowNo)
            If imDGSave(8) = 0 Then
                slStr = "Yes"
            ElseIf imDGSave(8) <> -1 Then
                slStr = "No"
            Else
                slStr = ""
                imDGSave(8) = 1 'No
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ACT1STOREDTIMESINDEX)
            imDGSave(9) = imLnSave(33, imLnRowNo)
            If imDGSave(9) = 0 Then
                slStr = "Yes"
            ElseIf imDGSave(9) <> -1 Then
                slStr = "No"
            Else
                slStr = ""
                imDGSave(9) = 1 'No
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ACT1STOREDSPOTSINDEX)
            imDGSave(10) = imLnSave(34, imLnRowNo)
            If imDGSave(10) = 0 Then
                slStr = "Yes"
            ElseIf imDGSave(10) <> -1 Then
                slStr = "No"
            Else
                slStr = ""
                imDGSave(10) = 1    'No
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ACT1STOREDCLEARPCTINDEX)
            imDGSave(11) = imLnSave(35, imLnRowNo)
            If imDGSave(11) = 0 Then
                slStr = "Yes"
            ElseIf imDGSave(11) <> -1 Then
                slStr = "No"
            Else
                slStr = ""
                imDGSave(11) = 1 'No
            End If
            gSetShow Contract.pbcOv, slStr, tmOV1Ctrls(ACT1DAYPARTFILTERINDEX)
            'If (smLineTypeOv <> "P") Then
            '    If tgSpf.sUsingBBs = "Y" Then
            '        Contract.pbcOvWD(0).Height = 1065
            '        Contract.pbcOvWD(1).Height = 1410
            '    Else
            '        Contract.pbcOvWD(0).Height = 720
            '        Contract.pbcOvWD(1).Height = 1065
            '    End If
            'Else
            '    Contract.pbcOvWD(0).Height = 720
            '    Contract.pbcOvWD(1).Height = 1065
            'End If
            '10256
            bmAct1PackageEnable = mIsPackageForActOne(smLineType, imLnRowNo)
            If bmAct1PackageEnable Then
                Contract.pbcOVWD(0).height = 1740    '1410
                Contract.pbcOVWD(1).height = 2085    '2120
            Else
                Contract.pbcOVWD(0).height = 1050    '1065
                Contract.pbcOVWD(1).height = 1395    '1425
            End If
            
            If (tgSpf.sAllowDailyBuys = "Y") Then
                Contract.pbcOv.height = Contract.pbcOVWD(1).height
                Contract.pbcOv.Width = Contract.pbcOVWD(1).Width
                Contract.pbcOv.Picture = Contract.pbcOVWD(1).Picture
            Else
                Contract.pbcOv.height = Contract.pbcOVWD(0).height
                Contract.pbcOv.Width = Contract.pbcOVWD(0).Width
                Contract.pbcOv.Picture = Contract.pbcOVWD(0).Picture
            End If
            Contract.plcOv.height = Contract.pbcOv.height + 345
            Contract.plcOv.Width = Contract.pbcOv.Width + 300
            Contract.plcOv.Move Contract.pbcLine.Left + tmLnCtrls(OV1INDEX).fBoxX, Contract.pbcLine.Top + tmLnCtrls(OV1INDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            If Contract.plcOv.Top + Contract.plcOv.height > Contract.cmcUpdate.Top + Contract.cmcUpdate.height Then
                Contract.plcOv.Move Contract.pbcLine.Left + tmLnCtrls(OV1INDEX).fBoxX, Contract.pbcLine.Top + tmLnCtrls(OV1INDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) - Contract.plcOv.height + (fgBoxGridH)
            End If
            Contract.pbcOv.Move Contract.plcOv.Left + 150, Contract.plcOv.Top + 225
            Contract.plcOv.Visible = True
            Contract.pbcOv.Visible = True
            'imOVBoxNo = OVDAYSINDEX
            'If (tgSpf.sUsingBBs = "Y") And (smLineTypeOv <> "P") Then
            '    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
            '        imOVBoxNo = OVBBOPENINDEX
            '    Else
            '        imOVBoxNo = OVCOPYINDEX 'OVBBCLOSEINDEX
            '    End If
            'Else
            '    If (tgUrf(0).iRemoteUserID <= 0) And (((tgUrf(0).iSlfCode <= 0) And (tgSpf.sGUsePropSys = "N")) Or (tgUrf(0).sWorkToHold = "Y") Or (tgUrf(0).sWorkToOrder = "Y") Or (tgUrf(0).sCompToHold = "Y") Or (tgUrf(0).sCompToOrder = "Y") Or (tgUrf(0).sIncompToHold = "Y") Or (tgUrf(0).sIncompToOrder = "Y")) Then
            '        If (tgSpf.sCPkEqual <> "Y") Then
            '            imOVBoxNo = OVHIGHINDEX
            '        Else
            '            imOVBoxNo = OVRANKINDEX
            '        End If
            '    Else
            '        imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX
            '    End If
            'End If
            If bmAct1PackageEnable And (smAlterHidden <> "C") Then
                If (Trim$(smLnSave(29, imLnRowNo)) = "") Or (imLnSave(32, imLnRowNo) < 0) Then
                    imOVBoxNo = ACT1LINEUPINDEX
                Else
                    imOVBoxNo = ACT1DAYPARTFILTERINDEX
                End If
            Else
                imOVBoxNo = OVCOPYINDEX
            End If
            If (Asc(tgSpf.sUsingFeatures8) And LRMANDATORY) = LRMANDATORY Then
                If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) = LIVECOPY) Or ((Asc(tgSpf.sUsingFeatures3) And PROMOCOPY) = PROMOCOPY) Then 'Using Live Copy
                    imOVBoxNo = OVCOPYINDEX 'OVENDTIMEINDEX  'OVPREEMPTINDEX
                End If
            End If
            Contract.mOvEnableBox imOVBoxNo
        Case OV2INDEX 'BB index
            imOVType = 2
            Contract.mSetTmOV2Ctrl
            mDPAllowedDays imLnSave(1, imLnRowNo), imAllowedDays()
            If (imLnSave(24, imLnRowNo) = -1) And (imLnSave(8, imLnRowNo) = 1) Then
                If (Asc(tgSpf.sOverrideOptions) And &H8) = &H8 Then
                    For ilDay = 0 To 6 Step 1
                        'If (imAllowedDays(ilDay + 1)) Then
                        If (imAllowedDays(ilDay + LBound(imAllowedDays))) Then
                            If imLnSave(12 + ilDay, imLnRowNo) = 0 Then
                                imLnSave(24 + ilDay, imLnRowNo) = 0
                            Else
                                imLnSave(24 + ilDay, imLnRowNo) = 1
                            End If
                        Else
                            imLnSave(24 + ilDay, imLnRowNo) = 1
                        End If
                    Next ilDay
                End If
            End If
            For ilLoop = OV2DAYSINDEX To OV2DAYSINDEX + 6 Step 1
                If imLnSave(8, imLnRowNo) = 1 Then
                    If imLnSave(24 + ilLoop - OV2DAYSINDEX, imLnRowNo) = 0 Then
                        slStr = "Yes"
                    Else
                        slStr = "No"
                    End If
                Else
                    slStr = ""
                End If
                gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(ilLoop)
                imSvOv2Days(ilLoop - OV2DAYSINDEX) = imLnSave(24 + ilLoop - OV2DAYSINDEX, imLnRowNo)
            Next ilLoop
            'Preferred Times
            slStr = smLnSave(25, imLnRowNo)
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2STARTTIMEINDEX)
            slStr = smLnSave(26, imLnRowNo)
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2ENDTIMEINDEX)

            If Trim$(smLnSave(28, imLnRowNo)) <> "" Then
                slStr = Trim$(smLnSave(28, imLnRowNo))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2SPLITINDEX)

            If Trim$(smLnSave(27, imLnRowNo)) <> "" Then
                slStr = Trim$(smLnSave(27, imLnRowNo))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2ALLOCATIONINDEX)

            If Contract.mUsingAcqCost() Then
                Contract.mSetDefAcqCost Val(smLnSave(16, imLnRowNo))
            End If
            If Trim$(smLnSave(24, imLnRowNo)) <> "" Then
                slStr = Trim$(smLnSave(24, imLnRowNo))
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2ACQUISITIONINDEX)

            If imLnSave(23, imLnRowNo) >= 0 Then
                If imLnSave(23, imLnRowNo) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
            Else
                slStr = ""
            End If
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV2SOLOINDEX)
            imSvSolo = imLnSave(23, imLnRowNo)

            If imLnSave(22, imLnRowNo) >= 0 Then
                If imLnSave(22, imLnRowNo) = 0 Then
                    slStr = "Yes"
                Else
                    slStr = "No"
                End If
            Else
                slStr = ""
            End If
            imSv1stPos = imLnSave(22, imLnRowNo)
            gSetShow Contract.pbcOv, slStr, tmOV2Ctrls(OV21STINDEX)

            Contract.pbcOv.height = Contract.pbcOVWD(2).height
            Contract.pbcOv.Width = Contract.pbcOVWD(2).Width
            Contract.pbcOv.Picture = Contract.pbcOVWD(2).Picture
            Contract.plcOv.height = Contract.pbcOv.height + 345
            Contract.plcOv.Width = Contract.pbcOv.Width + 300
            Contract.plcOv.Move Contract.pbcLine.Left + tmLnCtrls(OV2INDEX).fBoxX, Contract.pbcLine.Top + tmLnCtrls(OV2INDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            If Contract.plcOv.Top + Contract.plcOv.height > Contract.cmcUpdate.Top + Contract.cmcUpdate.height Then
                Contract.plcOv.Move Contract.pbcLine.Left + tmLnCtrls(OV1INDEX).fBoxX, Contract.pbcLine.Top + tmLnCtrls(OV2INDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) - Contract.plcOv.height + (fgBoxGridH)
            End If
            Contract.pbcOv.Move Contract.plcOv.Left + 150, Contract.plcOv.Top + 225
            Contract.plcOv.Visible = True
            Contract.pbcOv.Visible = True
            imOVBoxNo = imShowOV2
            Contract.mOvEnableBox imOVBoxNo
        Case LENINDEX 'Length index
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            'gShowHelpMess tmClfHelp(), CLFLENGTH
            Contract.mLenPop
            If imTerminate Then
                Exit Sub
            End If
            Contract.lbcLen.height = gListBoxHeight(Contract.lbcLen.ListCount, 5)
            Contract.edcLNDropDown.Width = tmLnCtrls(LENINDEX).fBoxW
            Contract.edcLNDropDown.MaxLength = 4
            gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(LENINDEX).fBoxX, tmLnCtrls(LENINDEX).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            imChgMode = True
            Contract.lbcLen.ListIndex = -1
            For ilLoop = 0 To Contract.lbcLen.ListCount - 1 Step 1
                If Contract.lbcLen.List(ilLoop) = smLnSave(16, imLnRowNo) Then
                    Contract.lbcLen.ListIndex = ilLoop
                    imComboBoxIndex = Contract.lbcLen.ListIndex
                    Contract.edcLNDropDown.Text = Contract.lbcLen.List(ilLoop)
                    Exit For
                End If
            Next ilLoop
            If Contract.lbcLen.ListIndex = -1 Then
                If imLnRowNo > 1 Then
                    gFindMatch smLnSave(16, imLnRowNo - 1), 0, Contract.lbcLen
                    If gLastFound(Contract.lbcLen) >= 0 Then
                        Contract.lbcLen.ListIndex = gLastFound(Contract.lbcLen)
                        imComboBoxIndex = Contract.lbcLen.ListIndex
                        Contract.edcLNDropDown.Text = Contract.lbcLen.List(Contract.lbcLen.ListIndex)
                    End If
                End If
                If Contract.lbcLen.ListIndex = -1 Then
                    If imVpfIndex >= 0 Then
                        gFindMatch str$(tgVpf(imVpfIndex).iSDLen), 0, Contract.lbcLen 'smBaseLen
                    Else
                        gFindMatch "30", 0, Contract.lbcLen 'smBaseLen
                    End If
                    If gLastFound(Contract.lbcLen) >= 0 Then
                        Contract.lbcLen.ListIndex = gLastFound(Contract.lbcLen)
                        imComboBoxIndex = Contract.lbcLen.ListIndex
                        Contract.edcLNDropDown.Text = Contract.lbcLen.List(Contract.lbcLen.ListIndex)
                    Else
                        If Contract.lbcLen.ListCount > 0 Then
                            Contract.lbcLen.ListIndex = 0
                            imComboBoxIndex = Contract.lbcLen.ListIndex
                            Contract.edcLNDropDown.Text = Contract.lbcLen.List(0)
                        Else
                            Contract.lbcLen.ListIndex = -1
                            imComboBoxIndex = Contract.lbcLen.ListIndex
                            Contract.edcLNDropDown.Text = ""
                        End If
                    End If
                End If
            End If
            imChgMode = False
            If imMoveLnRowNo - Contract.vbcLines.Value <= Contract.vbcLines.LargeChange \ 2 Then
                Contract.lbcLen.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcLen.Move Contract.edcLNDropDown.Left, Contract.edcLNDropDown.Top - Contract.lbcLen.height
            End If
            If (smLineType <> "O") And (smLineType <> "A") And (smLineType <> "E") Then
                'Check that vehicle defined, daypart and spot length to be able to compute price from Rate Card
                If (smLnSave(1, imLnRowNo) <> "") And (imLnSave(1, imLnRowNo) > 0) And (Contract.lbcLen.ListIndex >= 0) Then
                    Contract.cmcFlight.Enabled = True
                Else
                    Contract.cmcFlight.Enabled = False
                End If
            Else
                If (smLnSave(1, imLnRowNo) <> "") And (smLnSave(20, imLnRowNo) <> "") And (Contract.lbcLen.ListIndex >= 0) Then
                    Contract.cmcFlight.Enabled = True
                Else
                    Contract.cmcFlight.Enabled = False
                End If
            End If
            Contract.edcLNDropDown.SelStart = 0
            Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
            Contract.edcLNDropDown.Visible = True
            Contract.cmcLnDropDown.Visible = True
            Contract.edcLNDropDown.SetFocus
        Case WKSPOTSINDEX To WKSPOTSINDEX + 12
            If Contract.mEnableGame() Then
                Exit Sub
            End If
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            'If (tgClfCntr(imLnRowNo - 1).lFirstAvail = -1) And ((Contract.tscLine.SelectedItem.Index = 8) Or (Contract.tscLine.SelectedItem.Index = 9)) Then  'And (Trim$(smLnSave(15, ilRow)) <> "") Then
            '    Screen.MousePointer = vbHourglass
            '    bmInAvailTab = True
            '    mAvailTab imLnRowNo
            '    bmInAvailTab = False
            '    Screen.MousePointer = vbDefault
            'End If

            If imLnSave(8, imLnRowNo) = 0 Then
                Contract.pbcLnDaily.Width = tmLnCtrls(ilBoxNo).fBoxW
                gMoveTableCtrl Contract.pbcLine, Contract.pbcLnDaily, tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
                Contract.pbcOVWD(0).height = 360 '375
                gMoveTableCtrl Contract.pbcLine, Contract.pbcOVWD(0), tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) + Contract.pbcLnDaily.height
            Else
                Contract.edcLNDropDown.Width = tmLnCtrls(ilBoxNo).fBoxW
                Contract.edcLNDropDown.MaxLength = 4
                gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
                Contract.pbcOVWD(0).height = 360 '375
                gMoveTableCtrl Contract.pbcLine, Contract.pbcOVWD(0), tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15) + Contract.pbcLnDaily.height
            End If
            If lmHSStartDate < 0 Then
                Contract.mSetHdDate True, False
            End If
            llDate = lmHSStartDate + 7 * (imHSCurWkNo + ilBoxNo - WKSPOTSINDEX - 1)
            'mGetFlightDays imLnRowNo, imOvDays()
            imWkCff = mGetFlightSpotsDays(imLnRowNo, llDate, imOvDays(), lmWkSpots, imWkDays(), slWDFlag, ilDailySpots())
            If imLnSave(8, imLnRowNo) = 0 Then
                '5/19/12: Set spots per day for mSetShow can determine if changed
                If (lmWkSpots <= 0) And (tgClfCntr(imLnRowNo - 1).iFirstCff = -1) Then
                    For ilDay = 0 To 6 Step 1
                        smLnSave(7 + ilDay, imLnRowNo) = imOvDailySpots(ilDay)
                        imDailySpots(ilDay) = imOvDailySpots(ilDay)
                    Next ilDay
                Else
                    For ilDay = 0 To 6 Step 1
                        smLnSave(7 + ilDay, imLnRowNo) = ilDailySpots(ilDay)
                        imOvDailySpots(ilDay) = ilDailySpots(ilDay)
                        imDailySpots(ilDay) = ilDailySpots(ilDay)
                    Next ilDay
                End If
                For ilDay = 0 To 6 Step 1
                    slStr = smLnSave(7 + ilDay, imLnRowNo)
                    gSetShow Contract.pbcOVWD(0), slStr, tmDailyCtrls(ilDay + OV2DAYSINDEX)
                Next ilDay
                imOVType = 3
                Contract.edcLNDropDown.Text = ""
                If lmWkSpots = 0 Then
                    imLnDaily = 1
                    Contract.pbcOVWD(0).Visible = False
                Else
                    imLnDaily = 0
                    'If igFreezeState = 1 Then
                        If Not mIsLnGame(imLnRowNo) Then
                            If (tgUrf(0).sActFlightButton = "I") And ((Asc(tgSaf(0).sFeatures5) And SHOWDAYDROPDOWN) = SHOWDAYDROPDOWN) Then
                                Contract.pbcOVWD(0).Visible = True
                            End If
                        End If
                    'Else
                    '    Contract.pbcOvWD(0).Visible = False
                    'End If
                End If
                Contract.pbcLnDaily_Paint
                Contract.pbcLnDaily.Visible = True  'Set visibility
                Contract.pbcLnDaily.SetFocus
            Else
                smLineTypeOv = mGetLineTypeForOV(imLnRowNo)
                If (smLineTypeOv <> "P") Then
                    mDPAllowedDays imLnSave(1, imLnRowNo), imAllowedDays()
                Else
                    For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        If (StrComp(Trim$(smLnSave(20, imLnRowNo)), Trim$(tgMRdf(ilLoop).sName), 1) = 0) Or ((imLnSave(1, imLnRowNo) = tgMRdf(ilLoop).iCode) And (Trim$(smLnSave(20, imLnRowNo)) = "")) Then
                            mDPAllowedDays tgMRdf(ilLoop).iCode, imAllowedDays()
                            Exit For
                        End If
                    Next ilLoop
                End If
                If imWkCff <> -1 Then
                    For ilDay = 0 To 6 Step 1
                        If imWkDays(ilDay) Then
                            imLnSave(12 + ilDay, imLnRowNo) = 0
                        Else
                            imLnSave(12 + ilDay, imLnRowNo) = 1
                        End If
                    Next ilDay
                Else
                    'If want previous value, then eliminate this code as imLnSave(12+...) will have the laest.
                    For ilDay = 0 To 6 Step 1
                        If imOvDays(ilDay) = 1 Then
                            imLnSave(12 + ilDay, imLnRowNo) = 1
                        Else
                            imLnSave(12 + ilDay, imLnRowNo) = 0
                        End If
                    Next ilDay
                End If
                Contract.pbcOVWD(0).Cls
                For ilDay = 0 To 6 Step 1
                    If imLnSave(12 + ilDay, imLnRowNo) = 1 Then
                        slStr = "No"
                    ElseIf imLnSave(12 + ilDay, imLnRowNo) = 0 Then
                        slStr = "Yes"
                    Else
                        slStr = ""
                    End If
                    gSetShow Contract.pbcOVWD(0), slStr, tmDailyCtrls(ilDay + OV2DAYSINDEX)
                Next ilDay
                Contract.pbcOVWD_Paint 0
                imOVType = 3
                If lmWkSpots > 0 Then
                    Contract.edcLNDropDown.Text = Trim$(str$(lmWkSpots)) 'slStr'Trim$(smLnSave(6, imLnRowNo))
                    If Not mIsLnGame(imLnRowNo) Then
                        If (tgUrf(0).sActFlightButton = "I") And ((Asc(tgSaf(0).sFeatures5) And SHOWDAYDROPDOWN) = SHOWDAYDROPDOWN) Then
                            Contract.pbcOVWD(0).Visible = True
                        End If
                    End If
                Else
                    Contract.edcLNDropDown.Text = ""
                    Contract.pbcOVWD(0).Visible = False
                End If
                Contract.edcLNDropDown.SelStart = 0
                Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
                Contract.edcLNDropDown.Visible = True  'Set visibility
                Contract.edcLNDropDown.SetFocus
            End If
            'If (Contract.tscLine.SelectedItem.Index = 8) Or (Contract.tscLine.SelectedItem.Index = 9) Then
            If (Contract.tscLine.SelectedItem.Index = imTabMap(TABAVAIL)) Or (Contract.tscLine.SelectedItem.Index = imTabMap(TABPROTECTION)) Then
                Contract.edcLNDropDown.ToolTipText = mGetWeekToolTip(imLnRowNo, llDate, ilIgnoreRightMove, Contract.edcLNDropDown)
            End If

            If (Contract.plcDemoBar.Visible) And (imDBTotals < 5) Then
                imDBBoxNo = ilBoxNo - WKSPOTSINDEX + 1
                imDBRowNo = 1
                If imLnSave(8, imLnRowNo) = 0 Then
                    Contract.pbcDBLnDaily.Width = tmDBSCCtrls(imDBBoxNo).fBoxW
                    gMoveTableCtrl Contract.pbcDemoBar(1), Contract.pbcDBLnDaily, tmDBSCCtrls(imDBBoxNo).fBoxX, tmDBSCCtrls(imDBBoxNo).fBoxY
                    Contract.pbcDBLnDaily_Paint
                    Contract.pbcDBLnDaily.Visible = True  'Set visibility
                Else
                    Contract.edcDBValue.Width = tmDBSCCtrls(imDBBoxNo).fBoxW
                    Contract.edcDBValue.MaxLength = 4
                    gMoveTableCtrl Contract.pbcDemoBar(1), Contract.edcDBValue, tmDBSCCtrls(imDBBoxNo).fBoxX, tmDBSCCtrls(imDBBoxNo).fBoxY
                    Contract.edcDBValue.Text = Contract.edcLNDropDown.Text    'slStr'Trim$(smLnSave(6, imLnRowNo))
                    'Contract.edcDBValue.SelStart = 0
                    'Contract.edcDBValue.SelLength = Len(Contract.edcDBValue.Text)
                    Contract.edcDBValue.Visible = True  'Set visibility
                End If
            End If
        Case PRICEINDEX '(pop-up flight screen)
            tgClfCntr(imLnRowNo - 1).iPriceSet = 2
            If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                If Contract.pbcLineArrow.Visible Then
                    Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
                Else
                    Contract.pbcClickFocus.SetFocus
                End If
                Exit Sub
            End If
            'Set totals if coming from weeks
            Screen.MousePointer = vbHourglass
            'Add line if this is the last line
            If imLnRowNo >= UBound(tgClfCntr) + 1 Then
                If tgClfCntr(imLnRowNo - 1).iStatus = -1 Then
                    tgClfCntr(imLnRowNo - 1).iStatus = 0
                    imLnChg = True
                    imComputeReq = True
                End If
                ReDim Preserve tgClfCntr(0 To imLnRowNo) As CLFLIST
                ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
                ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                For ilLoop = LBound(imLnSave, FIELDBOUNDINDEX) To UBound(imLnSave, FIELDBOUNDINDEX) Step 1
                    imLnSave(ilLoop, UBound(tgClfCntr) + 1) = -1
                Next ilLoop
                Contract.mInitNewLine
                smLnSave(1, imLnRowNo + 1) = "" 'Vehicle list index
            End If
            'mSetTotalsFrmWks
            Screen.MousePointer = vbHourglass
            Contract.lbcPriceType(0).height = gListBoxHeight(Contract.lbcPriceType(0).ListCount, 5)
            Contract.edcLNDropDown.Width = tmLnCtrls(PRICEINDEX).fBoxW
            Contract.edcLNDropDown.MaxLength = 12
            gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
            Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
            imChgMode = True
            'Get Average Price always- Jim 10/15/97
            'Second call to reset smLnShow
           '10843 remove dead code
'            If (tgSpf.sCLnStdQt = "Y") And (imLnTotalType > 1) Then
'                mGetLnPriceInfo imLnRowNo, slPrice, imLnTotalType, PRICEINDEX
'                smLnAvgPrice = slPrice    'Used in Contract.mLnSetShow
'            Else
                mGetLnPriceInfo imLnRowNo, slPrice, 0, PRICEINDEX
                smLnAvgPrice = slPrice    'Used in Contract.mLnSetShow
                mGetLnPriceInfo imLnRowNo, slStr, imLnTotalType, PRICEINDEX
           ' End If

            gFindMatch slPrice, 1, Contract.lbcPriceType(0)
            If gLastFound(Contract.lbcPriceType(0)) > 0 Then
                Contract.lbcPriceType(0).ListIndex = gLastFound(Contract.lbcPriceType(0))
                Contract.edcLNDropDown.Text = Contract.lbcPriceType(0).List(Contract.lbcPriceType(0).ListIndex)
                smEnabledLnPrice = Contract.edcLNDropDown.Text
            Else
                If (slPrice = "") Or (tgClfCntr(imLnRowNo - 1).iOrigAnySpots < 0) Then
                    smEnabledLnPrice = ""
                    If sgCntrScreen = "Orders" Then
                        Contract.lbcPriceType(0).ListIndex = 0
                        Contract.edcLNDropDown.Text = ".00"
                    Else
                        Contract.lbcPriceType(0).ListIndex = 1
                        Contract.edcLNDropDown.Text = Contract.lbcPriceType(0).List(1)
                    End If
                    If tgClfCntr(imLnRowNo - 1).iOrigAnySpots < 0 Then
                        tgClfCntr(imLnRowNo - 1).iOrigAnySpots = 0
                    End If
                Else
                    Contract.lbcPriceType(0).ListIndex = 0
                    Contract.edcLNDropDown.Text = slPrice
                    smEnabledLnPrice = Contract.edcLNDropDown.Text
                End If
            End If
            imComboBoxIndex = Contract.lbcPriceType(0).ListIndex
            imChgMode = False
            If imMoveLnRowNo - Contract.vbcLines.Value <= Contract.vbcLines.LargeChange \ 2 Then
                Contract.lbcPriceType(0).Move Contract.cmcLnDropDown.Left + Contract.cmcLnDropDown.Width - Contract.lbcPriceType(0).Width, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
            Else
                Contract.lbcPriceType(0).Move Contract.cmcLnDropDown.Left + Contract.cmcLnDropDown.Width - Contract.lbcPriceType(0).Width, Contract.edcLNDropDown.Top - Contract.lbcPriceType(0).height
            End If
            If Not bmSettingPriceInTab Then
                Screen.MousePointer = vbDefault
                Contract.edcLNDropDown.SelStart = 0
                Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
                Contract.edcLNDropDown.Visible = True
                Contract.cmcLnDropDown.Visible = True
                Contract.edcLNDropDown.SetFocus
            End If
        Case LINECOMMINDEX
            'If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
            '    If Contract.pbcLineArrow.Visible Then
            '        Contract.pbcLineArrow.SetFocus   'Contract.pbcClickFocus.SetFocus
            '    Else
            '        Contract.pbcClickFocus.SetFocus
            '    End If
            '    Exit Sub
            'End If
           '10843 remove dead code
'            If tgSpf.sCLnStdQt = "Y" Then
'                imInLnComment = True
'                Contract.edcLNDropDown.Width = tmLnCtrls(ilBoxNo).fBoxW
'                Contract.edcLNDropDown.MaxLength = 1
'                gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(ilBoxNo).fBoxX, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
'                imChgMode = True
'                Contract.edcLNDropDown.Text = ""
'                Contract.edcComment.Text = smLnComment(imLnRowNo)
'                For ilLoop = 0 To 4 Step 1
'                    If imLnCommShowOn(ilLoop, imLnRowNo) = 0 Then
'                        Contract.ckcCommOn(ilLoop).Value = vbChecked
'                    Else
'                        Contract.ckcCommOn(ilLoop).Value = vbUnchecked
'                    End If
'                Next ilLoop
'                If Trim$(smLnComment(imLnRowNo)) = "" Then
'                    Contract.ckcCommOn(2).Value = vbChecked
'                End If
'                Contract.ckcCommOn(2).Enabled = True
'                Contract.ckcCommOn(4).Enabled = True
'                imChgMode = False
'                Screen.MousePointer = vbDefault
'                Contract.edcLNDropDown.SelStart = 0
'                Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
'                Contract.edcLNDropDown.Visible = True
'                Contract.edcLNDropDown.SetFocus
'            Else
                imInLnComment = True
                Contract.lbcComment(1).height = gListBoxHeight(Contract.lbcComment(1).ListCount, 2)
                Contract.edcLNDropDown.Width = 5 * tmLnCtrls(LINECOMMINDEX).fBoxW - Contract.cmcLnDropDown.Width
                Contract.edcLNDropDown.MaxLength = 20
                gMoveTableCtrl Contract.pbcLine, Contract.edcLNDropDown, tmLnCtrls(PRICEINDEX).fBoxX + Contract.cmcLnDropDown.Width, tmLnCtrls(ilBoxNo).fBoxY + (imMoveLnRowNo - Contract.vbcLines.Value - 1) * (fgBoxGridH + 15)
                Contract.cmcLnDropDown.Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width, Contract.edcLNDropDown.Top
                Contract.lbcComment(1).Move Contract.edcLNDropDown.Left + Contract.edcLNDropDown.Width + Contract.cmcLnDropDown.Width - Contract.lbcComment(1).Width, Contract.edcLNDropDown.Top + Contract.edcLNDropDown.height
                imChgMode = True
                If smLnComment(imLnRowNo) <> "" Then
                    Contract.lbcComment(1).Clear
                    Contract.lbcComment(1).AddItem Trim$(Left$(smLnComment(imLnRowNo), 20)), 0
                    Contract.lbcComment(1).AddItem "[None]", 0
                    Contract.edcLNDropDown.Text = Contract.lbcComment(1).List(1)
                Else
                    Contract.lbcComment(1).Clear
                    Contract.lbcComment(1).AddItem "[Comment]", 0
                    Contract.lbcComment(1).AddItem "[None]", 0
                    Contract.lbcComment(1).ListIndex = 0 '[None]
                    Contract.edcLNDropDown.Text = Contract.lbcComment(1).List(0)
                End If
                Contract.edcComment.Text = smLnComment(imLnRowNo)
                imComboBoxIndex = Contract.lbcComment(1).ListIndex
                For ilLoop = 0 To 4 Step 1
                    If imLnCommShowOn(ilLoop, imLnRowNo) = 0 Then
                        Contract.ckcCommOn(ilLoop).Value = vbChecked
                    Else
                        Contract.ckcCommOn(ilLoop).Value = vbUnchecked
                    End If
                Next ilLoop
                If Trim$(smLnComment(imLnRowNo)) = "" Then
                    Contract.ckcCommOn(2).Value = vbChecked
                End If
                Contract.ckcCommOn(2).Enabled = True
                Contract.ckcCommOn(4).Enabled = True
                imChgMode = False
                Screen.MousePointer = vbDefault
                If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
                    If Trim$(smLnComment(imLnRowNo)) <> "" Then
                        imDoubleClickName = True
                        ilRet = Contract.mLnCommentBranch()
                    Else
                        Contract.pbcClickFocus.SetFocus
                    End If
                    Exit Sub
                End If
                
                Contract.edcLNDropDown.SelStart = 0
                Contract.edcLNDropDown.SelLength = Len(Contract.edcLNDropDown.Text)
                Contract.edcLNDropDown.Visible = False   'True
                Contract.cmcLnDropDown.Visible = True
                Contract.cmcLnDropDown.SetFocus
            'End If
    End Select
    'mPaintLnBkgd imLnRowNo, GRAY
    Contract.mSetCommands
End Sub

Public Function mChkForStdPkgSub(ilLnRowNo As Integer, LINECOMMINDEX As Integer, VEHINDEX As Integer, DPINDEX As Integer, LENINDEX As Integer, PRICEINDEX As Integer, OV1INDEX As Integer) As Boolean

    Dim slLineType As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilCode As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim ilVpfIndex As Integer
    Dim ilLen As Integer
    Dim ilFound As Integer
    Dim ilSvLnRowNo As Integer
    Dim ilCff As Integer
    Dim ilNewCff As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim llTstDate As Long
    Dim llSvDate1 As Long
    Dim llSvDate2 As Long
    ReDim ilDays(0 To 6) As Integer
    Dim llSpotCount As Long
    Dim ilSpots As Integer
    Dim ilDay As Integer
    Dim ilUpper As Integer
    Dim ilPrevIndex As Integer
    Dim slDate As String
    Dim llPPrice As Long
    Dim llAPrice As Long
    Dim ilPvfIndex As Integer
    Dim ilMnfDemo As Integer
    Dim llAvgAud As Long
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llTime As Long
    Dim ilMnfSocEco As Integer
    Dim ilPvfSpotAdj As Integer
    Dim slPkLineType As String
    Dim ilAnyRated As Integer
    Dim llReallDate As Long
    Dim llPopEst As Long
    Dim ilPassDnfCode As Integer
    ReDim ilAudDays(0 To 6) As Integer
    Dim slWDFlag As String
    Dim slStr As String
    Dim llRafCode As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilGVef As Integer
    '12/18/15: Adjust only lines with proposal prices or unless no proposal prices
    Dim blPropPriceExist As Boolean
    Dim ilFreezeState As Integer

    mChkForStdPkgSub = True
    If imInFlightScreen Then
        Exit Function
    End If
    On Error GoTo TotalErr:
    gUnpackDateLong tgSpf.iReallDate(0), tgSpf.iReallDate(1), llReallDate
    llOvStartTime = 0
    llOvEndTime = 0
    ilMnfSocEco = 0
    If Contract.lbcDemo(0).ListIndex >= 1 Then
        slNameCode = tgDemoCode(Contract.lbcDemo(0).ListIndex - 1).sKey  'Traffic!Contract.lbcDemoCode.List(Contract.lbcDemo(ilLoop).ListIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfDemo = CInt(slCode)
    Else
        ilMnfDemo = 0
    End If
    slLineType = mGetLineType(ilLnRowNo)
    If (slLineType = "A") Or (slLineType = "O") Or (slLineType = "E") Then
        Exit Function
    End If
    'If (ilLnRowNo < LBound(smLnSave, 2)) Or (ilLnRowNo > UBound(smLnSave, 2)) Then
    If (ilLnRowNo < imLB1Or2) Or (ilLnRowNo > UBound(smLnSave, LINEBOUNDINDEX)) Then
        Exit Function
    End If
    gFindMatch smLnSave(1, ilLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) < 0 Then
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    ilFreezeState = igFreezeState
    igFreezeState = 0
    '12/18/15: Adjust only lines with proposal prices or unless no proposal prices
    blPropPriceExist = False
    ilSvLnRowNo = ilLnRowNo
    ilLen = Val(smLnSave(16, ilLnRowNo))
    imPvfRecLen = Len(tmPvf)
    If igTabMapIndex = 1 Then
        slNameCode = tmVehicleCode1(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
    ElseIf igTabMapIndex = 2 Then
        slNameCode = tmVehicleCode2(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
    Else
        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
    End If
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilCode = CInt(slCode)

    ReDim tmPBDPGen(0 To 0) As PBDPGEN
    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
    '    If tgMVef(ilVef).iCode = ilCode Then
        ilVef = gBinarySearchVef(ilCode)
        If ilVef <> -1 Then
            If (tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0) Then
                'If plcDemoBar.Visible Then  'Hide demo bar
                '    cmcDemoBar_Click
                'End If
                'Read Pvf and Build lines
                imExpandingPkg = True
                tmPvfSrchKey.lCode = tgMVef(ilVef).lPvfCode
                ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    Contract.mMoveCtrlToRec True 'Set tgChg.iRcfCode
                    ilUpper = UBound(tgPackLn)
                    tgPackLn(ilUpper).sType = tgMVef(ilVef).sStdInvTime
                    slPkLineType = tgPackLn(ilUpper).sType
                    'If (imLnSave(4, ilSvLnRowNo) = 1) And (slPkLineType = "E") Then
                    'Removed E test to match change code down below (2/1/01)
                    If (imLnSave(4, ilSvLnRowNo) = 1) Then
                        llOvStartTime = gTimeToLong(smLnSave(2, ilSvLnRowNo), False)
                        llOvEndTime = gTimeToLong(smLnSave(3, ilSvLnRowNo), True)
                    End If
                    'If tgMVef(ilVef).sStdAlter = "Y" Then
                    '    tgPackLn(ilUpper).iAlter = True
                    'Else
                    '    tgPackLn(ilUpper).iAlter = False
                    'End If
                    tgPackLn(ilUpper).sAlter = tgMVef(ilVef).sStdAlter
                    tgPackLn(ilUpper).sAlterName = tgMVef(ilVef).sStdAlterName
                    tgPackLn(ilUpper).iStdPkg = True
                    tgPackLn(ilUpper).sSource = "A"
                    tgPackLn(ilUpper).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
                    tgPackLn(ilUpper).iIndex = -1
                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilIndex).iCode = ilCode Then
                        ilIndex = gBinarySearchVef(ilCode)
                        If ilIndex <> -1 Then
                            If tgMVef(ilIndex).sState = "A" Then
                    '            Exit For
                    '        End If
                                ilFound = True
                                tgPackLn(ilUpper).sVehName = tgMVef(ilIndex).sName
                                tgPackLn(ilUpper).iMnfVehGp3Mkt = tgMVef(ilIndex).iMnfVehGp3Mkt
                    '        Exit For
                            End If
                        End If
                    'Next ilIndex
                    If ilFound Then
                        ilFound = False
                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        '    If tgMRdf(ilRdf).iCode = imLnSave(1, imLnRowNo) Then
                            ilRdf = gBinarySearchRdf(imLnSave(1, imLnRowNo))
                            If ilRdf <> -1 Then
                                If tgMRdf(ilRdf).sState = "A" Then
                            '            Exit For
                            '        End If
                                    ilFound = True
                                    tgPackLn(ilUpper).sDPName = Trim$(tgMRdf(ilRdf).sName)
                                    smLnSave(20, imLnRowNo) = Trim$(tgMRdf(ilRdf).sName)
                                End If
                        '        Exit For
                            End If
                        'Next ilRdf
                    End If
                    If Not ilFound Then
                        ilRet = -1
                    Else
                        ReDim Preserve tgPackLn(0 To ilUpper + 1) As PACKLN
                        ilPrevIndex = ilUpper
                        ilUpper = ilUpper + 1
                    End If
                End If
                ReDim lmPvfCode(0 To 0) As Long
                ReDim imPvfSpotAdj(0 To 0) As Integer
                ilPvfIndex = 0
                ilPvfSpotAdj = 1
                Do While ilRet = BTRV_ERR_NONE
                    'Create Lines
                    For ilLoop = LBound(tmPvf.iVefCode) To UBound(tmPvf.iVefCode) Step 1
                        If (tmPvf.iVefCode(ilLoop) > 0) And (tmPvf.iRdfCode(ilLoop) > 0) Then
                            tmPBDPGen(UBound(tmPBDPGen)).iRdfCode = tmPvf.iRdfCode(ilLoop)
                            tmPBDPGen(UBound(tmPBDPGen)).iVefCode = tmPvf.iVefCode(ilLoop)
                            tmPBDPGen(UBound(tmPBDPGen)).iAvgRating = ilPvfSpotAdj * tmPvf.iNoSpot(ilLoop)
                            tmPBDPGen(UBound(tmPBDPGen)).lAvgPrice = tmPvf.iPctRate(ilLoop)
                            ilFound = False
                            'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '    If tgMVef(ilIndex).iCode = tmPvf.iVefCode(ilLoop) Then
                                ilIndex = gBinarySearchVef(tmPvf.iVefCode(ilLoop))
                                If ilIndex <> -1 Then
                                    If tgMVef(ilIndex).sState = "A" Then
                            '            Exit For
                            '        End If
                                        If (tgMVef(ilIndex).sType = "P") And (tgMVef(ilIndex).lPvfCode > 0) Then
                                            ilFound = False
                                            lmPvfCode(UBound(lmPvfCode)) = tgMVef(ilIndex).lPvfCode
                                            imPvfSpotAdj(UBound(imPvfSpotAdj)) = tmPvf.iNoSpot(ilLoop)
                                            ReDim Preserve lmPvfCode(0 To UBound(lmPvfCode) + 1) As Long
                                            ReDim Preserve imPvfSpotAdj(0 To UBound(imPvfSpotAdj) + 1) As Integer
                                        ElseIf (tgMVef(ilIndex).sType = "P") Then
                                            ilFound = False
                                        Else
                                            ilFound = True
                                            tmPBDPGen(UBound(tmPBDPGen)).sVehName = tgMVef(ilIndex).sName
                                        End If
                                    End If
                            '        Exit For
                                End If
                            'Next ilIndex
                            If ilFound Then
                                ilFound = False
                                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                    If (tgMRif(llRif).iRcfCode = tmRcf.iCode) Then
                                        If (tgMRif(llRif).iVefCode = tmPvf.iVefCode(ilLoop)) And (tgMRif(llRif).iRdfCode = tmPvf.iRdfCode(ilLoop)) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                Next llRif
                            End If
                            If ilFound Then
                                ilFound = False
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                '    If tgMRdf(ilRdf).iCode = tmPvf.iRdfcode(ilLoop) Then
                                    ilRdf = gBinarySearchRdf(tmPvf.iRdfCode(ilLoop))
                                    If ilRdf <> -1 Then
                                        If tgMRdf(ilRdf).sState = "A" Then
                                    '            Exit For
                                    '        End If
                                            ilFound = True
                                            tmPBDPGen(UBound(tmPBDPGen)).sDPName = tgMRdf(ilRdf).sName
                                        End If
                                '        Exit For
                                    End If
                                'Next ilRdf
                            End If
                            If ilFound Then
                                ReDim Preserve tmPBDPGen(0 To UBound(tmPBDPGen) + 1) As PBDPGEN
                            End If
                        End If
                    Next ilLoop
                    If tmPvf.lLkPvfCode <= 0 Then
                        If ilPvfIndex >= UBound(lmPvfCode) Then
                            Exit Do
                        Else
                            tmPvf.lLkPvfCode = lmPvfCode(ilPvfIndex)
                            ilPvfSpotAdj = imPvfSpotAdj(ilPvfIndex)
                            ilPvfIndex = ilPvfIndex + 1
                        End If
                    End If
                    tmPvfSrchKey.lCode = tmPvf.lLkPvfCode
                    ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                Loop
                'Split Package line into separate weeks in cff because of pricing
                tgClfCntr(ilSvLnRowNo - 1).bPkExpanded = Not bgHideHiddenLines
                ilCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llEndDate
                        If llStartDate + 6 < llEndDate Then
                            'Save image as it is used to reset into new flight created, then dates changed
                            tmCff = tgCffCntr(ilCff).CffRec
                            If tmCff.sDyWk = "D" Then
                                slWDFlag = "D"
                                For ilIndex = 0 To 6 Step 1
                                    ilDays(ilIndex) = tmCff.iDay(ilIndex)
                                Next ilIndex
                            Else
                                slWDFlag = "W"
                            End If
                            For llDate = llStartDate To llEndDate Step 7
                                mAddWeek imLastCffUsed, ilSvLnRowNo, llDate, slWDFlag, llSpotCount, ilDays()
                                ilNewCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                                Do While ilNewCff <> -1
                                    gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llSvDate1
                                    If llDate = llSvDate1 Then
                                        gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iEndDate(0), tgCffCntr(ilNewCff).CffRec.iEndDate(1), llSvDate2
                                        tgCffCntr(ilNewCff).CffRec = tmCff
                                        gPackDateLong llSvDate1, tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1)
                                        gPackDateLong llSvDate2, tgCffCntr(ilNewCff).CffRec.iEndDate(0), tgCffCntr(ilNewCff).CffRec.iEndDate(1)
                                        Exit Do
                                    End If
                                    ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                Loop
                            Next llDate
                            ilCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                        End If
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
                ilAnyRated = False
                llAPrice = 0
                For ilLoop = LBound(tmPBDPGen) To UBound(tmPBDPGen) - 1 Step 1
                    If bgCntrFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Contract, tmPBDPGen(ilLoop).iVefCode)
                        bgCntrFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(tmPBDPGen(ilLoop).iVefCode)
                    End If
                    ilFound = False
                    For ilIndex = LBound(tgVpf(ilVpfIndex).iSLen) To UBound(tgVpf(ilVpfIndex).iSLen) Step 1
                        If tgVpf(ilVpfIndex).iSLen(ilIndex) <> 0 Then
                            If tgVpf(ilVpfIndex).iSLen(ilIndex) = ilLen Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilIndex
                    If ilFound Then
                        'Test if any days to book into
                        'If slPkLineType = "E" Then 'Removed when adding Rate Bar-6/7/99 Ok'd by Jim
                            mDPAllowedDays tmPBDPGen(ilLoop).iRdfCode, imAllowedDays()
                            ilFound = False
                            For ilDay = 0 To 6 Step 1
                                If (imAllowedDays(ilDay + LBound(imAllowedDays))) Then
                                    If imLnSave(12 + ilDay, ilSvLnRowNo) <> 1 Then
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                            Next ilDay
                            If (imLnSave(4, ilSvLnRowNo) = 1) And (ilFound) Then
                                'Test override times
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) Step 1
                                '    If tmPBDPGen(ilLoop).iRdfcode = tgMRdf(ilRdf).iCode Then
                                    ilRdf = gBinarySearchRdf(tmPBDPGen(ilLoop).iRdfCode)
                                    If ilRdf <> -1 Then
                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llTime
                                                If llOvEndTime <= llTime Then
                                                    ilFound = False
                                                    Exit For
                                                End If
                                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llTime
                                                If llOvStartTime >= llTime Then
                                                    ilFound = False
                                                    Exit For
                                                End If
                                            End If
                                        Next ilIndex
                                '        Exit For
                                    End If
                                'Next ilRdf
                            End If
                        'End If
                    End If
                    If ilFound Then
                        ReDim Preserve tgClfCntr(0 To UBound(tgClfCntr) + 1) As CLFLIST
                        ReDim Preserve smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
                        ReDim Preserve imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
                        ReDim Preserve smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
                        ReDim Preserve smLnComment(0 To UBound(tgClfCntr) + 1) As String
                        ReDim Preserve smLnInitComment(0 To UBound(tgClfCntr) + 1) As String
                        ReDim Preserve imLnCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                        ReDim Preserve imLnInitCommShowOn(0 To 4, 0 To UBound(tgClfCntr) + 1) As Integer
                        imLnRowNo = UBound(tgClfCntr) 'ilLineNo - 1    'Init line area to be filled
                        Contract.mInitNewLine
                        Screen.MousePointer = vbHourglass
                        tgClfCntr(imLnRowNo).ClfRec.iLine = 0
                        imLnRowNo = UBound(tgClfCntr) - 1 'ilLineNo    'Init line area not used first
                        Contract.mInitNewLine
                        Screen.MousePointer = vbHourglass
                        tgClfCntr(imLnRowNo).iStatus = 0
                        imLnChg = True
                        imComputeReq = True
                        'slStr = Trim$(Str$(tgClfCntr(imLnRowNo).ClfRec.iLine))
                        'gSetShow Contract.pbcLine, slStr, tmLnCtrls(LINEINDEX)
                        'smLnShow(LINEINDEX, imLnRowNo + 1) = tmLnCtrls(LINEINDEX).sShow

                        imLnRowNo = UBound(tgClfCntr)
                        tgClfCntr(imLnRowNo - 1).iSPChg = True
                        igLnSPChg = True
                        'Set values into line area
                        imLnBoxNo = VEHINDEX
                        Contract.edcLNDropDown.MaxLength = 0
                        Contract.edcLNDropDown.Text = Trim$(tmPBDPGen(ilLoop).sVehName)  'lbcAllVehicle.List(ilVeh)
                        Contract.edcLnDropDown_Change
                        Contract.mLnSetShow imLnBoxNo
                        Screen.MousePointer = vbHourglass
                        imLnBoxNo = DPINDEX
                        Contract.edcLNDropDown.MaxLength = 0
                        Contract.edcLNDropDown.Text = mMakePrgName(tmPBDPGen(ilLoop).iRdfCode)   'Trim$(tmPBDPGen(ilLoop).sDPName)    'lbcDaypart.List(ilDP)
                        Contract.edcLnDropDown_Change
                        Contract.mLnSetShow imLnBoxNo
                        Screen.MousePointer = vbHourglass
                        imLnSave(2, imLnRowNo) = -1 'Buy type 0=Program; 1=Time
                        imLnSave(3, imLnRowNo) = 0  'BB: N/A
                        imLnSave(4, imLnRowNo) = 0  'Overtime flag
                        smLnSave(2, imLnRowNo) = "" 'Override start time
                        smLnSave(3, imLnRowNo) = "" 'Override end time
                        imLnSave(31, imLnRowNo) = imLnSave(31, ilSvLnRowNo)
                        imLnSave(6, imLnRowNo) = imLnSave(6, ilSvLnRowNo)
                        imDGSave(3) = imLnSave(6, imLnRowNo)
                        imLnSave(19, imLnRowNo) = imLnSave(19, ilSvLnRowNo)
                        imDGSave(4) = imLnSave(19, ilSvLnRowNo)
                        imDGSave(5) = 0     'Open BB
                        imDGSave(6) = 0     'Close BB
                        '11/19/20: Clear line BB equivalent value so that override question is not generated becuase overrides not matching
                        imLnSave(20, imLnRowNo) = imDGSave(5)
                        imLnSave(21, imLnRowNo) = imDGSave(6)
                        'end of chg
                        
                        imDGSave(7) = imLnSave(31, imLnRowNo)   'Live/Recorded
                        imDGSave(8) = -1    'Stored Times
                        imDGSave(9) = -1    'Stored Spots
                        imDGSave(10) = -1   'Stored Clear %
                        imDGSave(11) = -1   'Daypart Filter
                        smDGSave(13) = ""
                        smDGSave(14) = ""
                        smDGSave(15) = ""   'ACT1 Liseup
                        If (tgClfCntr(imLnRowNo - 1).sOwnership = "B") Or (tgClfCntr(imLnRowNo - 1).sOwnership = "D") Then
                            imLnSave(4, imLnRowNo) = imLnSave(4, ilSvLnRowNo)
                            smLnSave(2, imLnRowNo) = smLnSave(2, ilSvLnRowNo)
                            smLnSave(3, imLnRowNo) = smLnSave(3, ilSvLnRowNo)
                            smDGSave(13) = smLnSave(2, imLnRowNo)
                        Else
                            If (imLnSave(4, ilSvLnRowNo) = 1) Then  'Remove when adding Rate Bar code- 6/7/99 Ok'd by JimAnd (slPkLineType = "E") Then
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) Step 1
                                '    If tmPBDPGen(ilLoop).iRdfcode = tgMRdf(ilRdf).iCode Then
                                    ilRdf = gBinarySearchRdf(tmPBDPGen(ilLoop).iRdfCode)
                                    If ilRdf <> -1 Then
                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llTime
                                                If llOvStartTime > llTime Then
                                                    imLnSave(4, imLnRowNo) = imLnSave(4, ilSvLnRowNo)  'Time override
                                                    smLnSave(2, imLnRowNo) = smLnSave(2, ilSvLnRowNo)
                                                    smDGSave(13) = smLnSave(2, imLnRowNo)
                                                End If
                                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llTime
                                                If llOvEndTime < llTime Then
                                                    If imLnSave(4, imLnRowNo) = 0 Then
                                                        gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), "A", "1", smLnSave(2, imLnRowNo)
                                                        smDGSave(13) = smLnSave(2, imLnRowNo)
                                                    End If
                                                    imLnSave(4, imLnRowNo) = imLnSave(4, ilSvLnRowNo)  'Time override
                                                    smLnSave(3, imLnRowNo) = smLnSave(3, ilSvLnRowNo)
                                                    smDGSave(14) = smLnSave(3, imLnRowNo)
                                                Else
                                                    If imLnSave(4, imLnRowNo) = 1 Then
                                                        smLnSave(3, imLnRowNo) = gFormatTimeLong(llTime, "A", "1")
                                                        smDGSave(14) = smLnSave(3, imLnRowNo)
                                                    End If
                                                End If
                                            End If
                                        Next ilIndex
                                '        Exit For
                                    End If
                                'Next ilRdf
                            End If
                        End If
                        imLnSave(5, imLnRowNo) = 0  'Unique Spot Types: N/A
                        imLnBoxNo = LENINDEX
                        Contract.edcLNDropDown.Text = Trim$(str$(ilLen))   'Trim$(Str$(lbcLength.List(ilLen)))
                        Contract.edcLnDropDown_Change
                        Contract.mLnSetShow imLnBoxNo
                        Screen.MousePointer = vbHourglass
                        imLnBoxNo = PRICEINDEX
                        'If rbcPBPreempt(1).Value Then
                        '    imLnSave(6, imLnRowNo) = 1
                        'Else
                        '    imLnSave(6, imLnRowNo) = 0
                        'End If
                        mDPAllowedDays imLnSave(1, imLnRowNo), imAllowedDays()
                        ilFound = False
                        For ilDay = 0 To 6 Step 1
                            If (imAllowedDays(ilDay + LBound(imAllowedDays))) Then
                                'If slPkLineType <> "E" Then    'Removed as part of Rate Bar 6/7/99- Jim Ok'd
                                '    ilDays(ilDay) = 1
                                '    ilAudDays(ilDay) = True
                                '    ilFound = True
                                'Else
                                    If imLnSave(12 + ilDay, ilSvLnRowNo) <> 1 Then
                                        ilDays(ilDay) = 1
                                        ilAudDays(ilDay) = True
                                        ilFound = True
                                    Else
                                        ilDays(ilDay) = 0
                                        ilAudDays(ilDay) = False
                                    End If
                                'End If
                            Else
                                ilDays(ilDay) = 0
                                ilAudDays(ilDay) = False
                            End If
                        Next ilDay
                        If Not ilFound Then 'det days on and spots to zero
                            For ilDay = 0 To 6 Step 1
                                If (imAllowedDays(ilDay + LBound(imAllowedDays))) Then
                                    ilDays(ilDay) = 1
                                    ilAudDays(ilDay) = True
                                Else
                                    ilDays(ilDay) = 0
                                    ilAudDays(ilDay) = False
                                End If
                            Next ilDay
                        End If
                        ilCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                        Do While ilCff <> -1
                            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llEndDate
                                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                    slWDFlag = "D"
                                    For ilDay = 0 To 6 Step 1
                                        ilDays(ilDay) = tmPBDPGen(ilLoop).iAvgRating * tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                        smDGSave(3 + ilDay) = Trim$(str$(ilDays(ilDay)))
                                    Next ilDay
                                    imDGSave(2) = 0
                                Else
                                    slWDFlag = "W"
                                    imDGSave(2) = 1
                                End If
                                imLnSave(8, imLnRowNo) = imDGSave(2)
                                For llDate = llStartDate To llEndDate Step 7
                                    slDate = Format$(llDate, "m/d/yy")
                                    If slWDFlag = "D" Then
                                        llSpotCount = 0
                                        For ilIndex = 0 To 6 Step 1
                                            llSpotCount = llSpotCount + ilDays(ilIndex)
                                        Next ilIndex
                                    Else
                                        llSpotCount = tmPBDPGen(ilLoop).iAvgRating * tgCffCntr(ilCff).CffRec.iSpotsWk
                                    End If
                                    '6/25/12: Check if game vehicle
                                    ilGVef = gBinarySearchVef(tgClfCntr(imLnRowNo - 1).ClfRec.iVefCode)
                                    If ilGVef <> -1 Then
                                        If tgMVef(ilGVef).sType = "G" Then
                                            'Add CGF
                                            llSpotCount = mAddGames(hmGhf, hmGsf, tgClfCntr(imLnRowNo - 1).ClfRec.iVefCode, llSpotCount, imLnRowNo, llDate, imAllowedDays())
                                            slWDFlag = "W"
                                            imDGSave(2) = 1
                                            ilFound = True
                                        End If
                                    End If
                                    If ilFound Then
                                        mAddWeek imLastCffUsed, imLnRowNo, llDate, slWDFlag, llSpotCount, ilDays()
                                    Else
                                        mAddWeek imLastCffUsed, imLnRowNo, llDate, slWDFlag, 0, ilDays()
                                    End If
                                    'Set Price
                                    If slPkLineType = "E" Then
                                        ilNewCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                                        Do While ilNewCff <> -1
                                            gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llTstDate
                                            If llTstDate = llDate Then
                                                tgCffCntr(ilNewCff).CffRec.sPriceType = "T"
                                                tgCffCntr(ilNewCff).CffRec.lPropPrice = 0
                                                Exit Do
                                            End If
                                            ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                        Loop
                                    Else
                                        ilNewCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                                        Do While ilNewCff <> -1
                                            gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llTstDate
                                            If llTstDate = llDate Then
                                                '7/9/16: Set line flag that price is set.  With this setting new weeks will be added with current average instead of the rate card
                                                tgClfCntr(imLnRowNo - 1).iPriceSet = 2
                                                tgCffCntr(ilNewCff).CffRec.sPriceType = "T"
                                                tgCffCntr(ilNewCff).CffRec.lPropPrice = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slDate, imCntrVefCode, imLnRowNo)
                                                '12/18/15: Adjust only lines with proposal price unless none exist
                                                If tgCffCntr(ilNewCff).CffRec.lPropPrice > 0 Then
                                                    blPropPriceExist = True
                                                End If
                                                If tgMVef(ilVef).sStdPrice = "P" Then
                                                    'tgCffCntr(ilNewCff).CffRec.lActPrice = (tmPBDPGen(ilLoop).lAvgPrice * tgCffCntr(ilCff).CffRec.lActPrice) \ 10000
                                                    tgCffCntr(ilNewCff).CffRec.lActPrice = tmPBDPGen(ilLoop).lAvgPrice
                                                ElseIf tgMVef(ilVef).sStdPrice = "A" Then   'Audience
                                                    llAvgAud = 0
                                                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                    '    If tgMVef(ilIndex).iCode = tmPBDPGen(ilLoop).iVefCode Then
                                                        ilIndex = gBinarySearchVef(tmPBDPGen(ilLoop).iVefCode)
                                                        If ilIndex <> -1 Then
                                                            'If (tgMVef(ilIndex).iDnfCode > 0) And (ilMnfDemo > 0) Then
                                                            '    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, tgMVef(ilIndex).iDnfCode, tmPBDPGen(ilLoop).iVefCode, ilMnfSocEco, ilMnfDemo, tmPBDPGen(ilLoop).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), llAvgAud)
                                                            'End If
                                                            '4/4/06-Hidden line property not set until later, don't use GetLineType
                                                            slStr = "H" 'mGetLineType(imLnRowNo)
                                                            llRafCode = 0
                                                            If tgSpf.iReallMnfDemo <= 0 Then
                                                                If (tgMVef(ilIndex).iReallDnfCode > 0) And (ilMnfDemo > 0) Then
                                                                    ilPassDnfCode = tgMVef(ilIndex).iReallDnfCode
                                                                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmPBDPGen(ilLoop).iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmPBDPGen(ilLoop).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), slStr, llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                                                    '10256
                                                                    mActOneLineInfoSet imLnRowNo
                                                                ElseIf (tgMVef(ilIndex).iDnfCode > 0) And (ilMnfDemo > 0) And (llReallDate <= 0) Then
                                                                    ilPassDnfCode = tgMVef(ilIndex).iDnfCode
                                                                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmPBDPGen(ilLoop).iVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmPBDPGen(ilLoop).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), slStr, llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                                                    '10256
                                                                    mActOneLineInfoSet imLnRowNo
                                                                End If
                                                            Else
                                                                If (tgMVef(ilIndex).iReallDnfCode > 0) And (tgSpf.iReallMnfDemo > 0) Then
                                                                    ilPassDnfCode = tgMVef(ilIndex).iReallDnfCode
                                                                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilPassDnfCode, tmPBDPGen(ilLoop).iVefCode, ilMnfSocEco, tgSpf.iReallMnfDemo, llDate, llDate, tmPBDPGen(ilLoop).iRdfCode, llOvStartTime, llOvEndTime, ilAudDays(), slStr, llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                                                    '10256
                                                                    mActOneLineInfoSet imLnRowNo
                                                                End If
                                                            End If
                                                    '        Exit For
                                                        End If
                                                    'Next ilIndex
                                                    tgCffCntr(ilNewCff).CffRec.lActPrice = llAvgAud
                                                ElseIf tgMVef(ilVef).sStdPrice = "S" Then   'Spot Count
                                                    tgCffCntr(ilNewCff).CffRec.lActPrice = llSpotCount
                                                Else
                                                    'Set price temporary
                                                    tgCffCntr(ilNewCff).CffRec.lActPrice = tgCffCntr(ilNewCff).CffRec.lPropPrice 'tgCffCntr(ilCff).CffRec.lActPrice
                                                    llAPrice = llAPrice + tgCffCntr(ilNewCff).CffRec.lPropPrice
                                                End If
                                                'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                '    If tgMVef(ilIndex).iCode = tmPBDPGen(ilLoop).iVefCode Then
                                                    ilIndex = gBinarySearchVef(tmPBDPGen(ilLoop).iVefCode)
                                                    If ilIndex <> -1 Then
                                                        If tgMVef(ilIndex).iMnfVehGp5Rsch > 0 Then
                                                            For ilLp = LBound(tmRatedGp) To UBound(tmRatedGp) - 1 Step 1
                                                                If tgMVef(ilIndex).iMnfVehGp5Rsch = tmRatedGp(ilLp).iMnfVehGp5Rsch Then
                                                                    If tmRatedGp(ilLp).sRated <> "Y" Then
                                                                        tgCffCntr(ilNewCff).CffRec.lActPrice = -tgCffCntr(ilNewCff).CffRec.lActPrice
                                                                    Else
                                                                        ilAnyRated = True
                                                                    End If
                                                                    Exit For
                                                                End If
                                                            Next ilLp
                                                        End If
                                                 '       Exit For
                                                    End If
                                                'Next ilIndex
                                                Exit Do
                                            End If
                                            ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                        Loop
                                    End If
                                Next llDate
                                mGetFlightDays imLnRowNo, imOvDays()
                                imLnBoxNo = OV1INDEX
                                Contract.mLnSetShow imLnBoxNo
                                Screen.MousePointer = vbHourglass
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                        mInitLnTotals imLnRowNo, imLnTotalType, PRICEINDEX
                        tgPackLn(ilPrevIndex).iIndex = ilUpper
                        tgPackLn(ilUpper).sType = "H"
                        tgClfCntr(imLnRowNo - 1).bShowLine = Not bgHideHiddenLines
                        'If tgMVef(ilVef).sStdAlter = "Y" Then
                        '    tgPackLn(ilUpper).iAlter = True
                        'Else
                        '    tgPackLn(ilUpper).iAlter = False
                        'End If
                        tgPackLn(ilUpper).sAlter = tgMVef(ilVef).sStdAlter
                        tgPackLn(ilUpper).sAlterName = tgMVef(ilVef).sStdAlterName
                        tgPackLn(ilUpper).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine
                        tgPackLn(ilUpper).iIndex = -1
                        tgPackLn(ilUpper).sVehName = tmPBDPGen(ilLoop).sVehName
                        tgPackLn(ilUpper).sDPName = tmPBDPGen(ilLoop).sDPName
                        ReDim Preserve tgPackLn(0 To ilUpper + 1) As PACKLN
                        ilPrevIndex = ilUpper
                        ilUpper = ilUpper + 1
                    End If
                Next ilLoop
                If (llAPrice = 0) And (tgMVef(ilVef).sStdPrice = "R") And (Not ilAnyRated) Then
                    imLnRowNo = ilSvLnRowNo + 1 'First Line
                    ilCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                    Do While ilCff <> -1
                        If (tgMVef(ilVef).sStdPrice <> "S") Or ((slPkLineType = "A") And (tgMVef(ilVef).sStdPrice = "S")) Then
                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                ilSpots = 0
                                For ilDay = 0 To 6 Step 1
                                    ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                Next ilDay
                            Else
                                ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk
                            End If
                            llPPrice = ilSpots * tgCffCntr(ilCff).CffRec.lActPrice
                        Else
                            llPPrice = tgCffCntr(ilCff).CffRec.lActPrice
                        End If
                        If llPPrice > 0 Then
                            gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                            For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                                ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                                Do While ilNewCff <> -1
                                    gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        tgCffCntr(ilNewCff).CffRec.lActPrice = 1
                                        Exit Do
                                    End If
                                    ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                Loop
                            Next ilLoop
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                End If
                If slPkLineType <> "E" Then
                    imLnRowNo = ilSvLnRowNo + 1 'First Line
                    ilCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                    Do While ilCff <> -1
                        llAPrice = 0    'tgCffCntr(ilCff).CffRec.lActPrice
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                        For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                            ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                            Do While ilNewCff <> -1
                                gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                If llStartDate = llDate Then
                                    If ilAnyRated Then
                                        If tgCffCntr(ilNewCff).CffRec.lActPrice < 0 Then
                                            tgCffCntr(ilNewCff).CffRec.lActPrice = 0
                                        End If
                                    Else
                                        If tgCffCntr(ilNewCff).CffRec.lActPrice < 0 Then
                                            tgCffCntr(ilNewCff).CffRec.lActPrice = -tgCffCntr(ilNewCff).CffRec.lActPrice
                                        End If
                                    End If
                                    If (tgMVef(ilVef).sStdPrice = "P") Or (tgMVef(ilVef).sStdPrice = "S") Then
                                        llAPrice = llAPrice + tgCffCntr(ilNewCff).CffRec.lActPrice
                                    Else
                                        If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                            ilSpots = 0
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                        End If
                                        llAPrice = llAPrice + ilSpots * tgCffCntr(ilNewCff).CffRec.lActPrice
                                    End If
                                    Exit Do
                                End If
                                ilNewCff = tgCffCntr(ilNewCff).iNextCff
                            Loop
                        Next ilLoop
                        If (tgMVef(ilVef).sStdPrice <> "S") Or ((slPkLineType = "A") And (tgMVef(ilVef).sStdPrice = "S")) Or ((slPkLineType = "O") And (tgMVef(ilVef).sStdPrice = "S")) Then
                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                llPPrice = 0
                                For ilDay = 0 To 6 Step 1
                                    llPPrice = llPPrice + tgCffCntr(ilCff).CffRec.iDay(ilDay) * tgCffCntr(ilCff).CffRec.lActPrice
                                Next ilDay
                            Else
                                llPPrice = tgCffCntr(ilCff).CffRec.iSpotsWk * tgCffCntr(ilCff).CffRec.lActPrice
                            End If
                        Else
                            llPPrice = tgCffCntr(ilCff).CffRec.lActPrice
                        End If
                        'ilNewCff = tgClfCntr(ilSvLnRowNo - 1).iFirstCff
                        'Do While ilNewCff <> -1
                        '    gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                        '    If llStartDate = llDate Then
                        '        llPPrice = tgCffCntr(ilCff).CffRec.iSpotsWk * tgCffCntr(ilNewCff).CffRec.lActPrice
                        '        Exit Do
                        '    End If
                        '    ilNewCff = tgCffCntr(ilNewCff).iNextCff
                        'Loop
                        'Distribute dollars
                        If llAPrice > 0 Then
                            'tgCffCntr(ilCff).CffRec.lActPrice = (CSng(llPPrice) * tgCffCntr(ilCff).CffRec.lActPrice) / llAPrice
                            'If tgCffCntr(ilCff).CffRec.iSpotsWk > 0 Then
                            '    tgCffCntr(ilCff).CffRec.lActPrice = (tgCffCntr(ilCff).CffRec.lActPrice / tgCffCntr(ilCff).CffRec.iSpotsWk)
                            'End If
                            llSpotCount = 0
                            For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                                ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                                Do While ilNewCff <> -1
                                    gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        If tgMVef(ilVef).sStdPrice <> "S" Then
                                            tgCffCntr(ilNewCff).CffRec.lActPrice = (CSng(llPPrice) * tgCffCntr(ilNewCff).CffRec.lActPrice) / llAPrice
                                            If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                                ilSpots = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilSpots = ilSpots + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilSpots = tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                            End If
                                            If (ilSpots > 0) And (tgMVef(ilVef).sStdPrice = "P") Then
                                                tgCffCntr(ilNewCff).CffRec.lActPrice = tgCffCntr(ilNewCff).CffRec.lActPrice / ilSpots
                                            End If
                                        Else
                                            If tgCffCntr(ilNewCff).CffRec.lActPrice > 0 Then
                                                tgCffCntr(ilNewCff).CffRec.lActPrice = CSng(llPPrice) / llAPrice
                                            Else
                                                tgCffCntr(ilNewCff).CffRec.lActPrice = 0
                                            End If
                                            If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                                For ilDay = 0 To 6 Step 1
                                                    llSpotCount = llSpotCount + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                                Next ilDay
                                            Else
                                                llSpotCount = llSpotCount + tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                            End If
                                        End If
                                        Exit Do
                                    End If
                                    ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                Loop
                            Next ilLoop
                            If tgMVef(ilVef).sStdInvTime = "E" Then
                                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                    ilSpots = 0
                                    For ilDay = 0 To 6 Step 1
                                        ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                    Next ilDay
                                Else
                                    ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk
                                End If

                                If ilSpots > 0 Then
                                    tgCffCntr(ilCff).CffRec.lActPrice = llPPrice
                                Else
                                    tgCffCntr(ilCff).CffRec.lActPrice = 0
                                End If
                            End If
                            'llPPrice = llPPrice - tgCffCntr(ilCff).CffRec.lActPrice
                            gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                            For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                                ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                                Do While ilNewCff <> -1
                                    gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                    If llStartDate = llDate Then
                                        If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                            ilSpots = 0
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                        End If
                                        llPPrice = llPPrice - ilSpots * tgCffCntr(ilNewCff).CffRec.lActPrice
                                        Exit Do
                                    End If
                                    ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                Loop
                            Next ilLoop
                            If llPPrice <> 0 Then
                                'If tgCffCntr(ilCff).CffRec.iSpotsWk = 1 Then
                                '    tgCffCntr(ilCff).CffRec.lActPrice = tgCffCntr(ilCff).CffRec.lActPrice + llPPrice
                                '    llPPrice = 0
                                'End If
                                gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                                For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                                    If llPPrice = 0 Then
                                        Exit For
                                    End If
                                    ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                                    Do While ilNewCff <> -1
                                        gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                        If llStartDate = llDate Then
                                            If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                                ilSpots = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilSpots = ilSpots + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilSpots = tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                            End If
                                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                                ilRet = 0
                                                For ilDay = 0 To 6 Step 1
                                                    ilRet = ilRet + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                                Next ilDay
                                            Else
                                                ilRet = tgCffCntr(ilCff).CffRec.iSpotsWk
                                            End If
                                            '12/18/15: disallow negative result
                                            'If ilSpots = ilRet Then
                                            '    tgCffCntr(ilNewCff).CffRec.lActPrice = tgCffCntr(ilNewCff).CffRec.lActPrice + (llPPrice / ilSpots)
                                            '    llPPrice = 0
                                            'End If
                                            If (ilSpots = ilRet) And (tgCffCntr(ilNewCff).CffRec.lActPrice + (llPPrice / ilSpots) >= 0) Then
                                                If (tgCffCntr(ilNewCff).CffRec.lPropPrice > 0) Or (Not blPropPriceExist) Then
                                                    tgCffCntr(ilNewCff).CffRec.lActPrice = tgCffCntr(ilNewCff).CffRec.lActPrice + (llPPrice / ilSpots)
                                                    llPPrice = 0
                                                End If
                                            End If
                                            Exit Do
                                        End If
                                        ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                    Loop
                                Next ilLoop
                                '12/18/15: Test for 1 spot
                                If llPPrice <> 0 Then
                                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llStartDate
                                    For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                                        If llPPrice = 0 Then
                                            Exit For
                                        End If
                                        ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                                        Do While ilNewCff <> -1
                                            gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                            If llStartDate = llDate Then
                                                If tgCffCntr(ilNewCff).CffRec.sDyWk = "D" Then
                                                    ilSpots = 0
                                                    For ilDay = 0 To 6 Step 1
                                                        ilSpots = ilSpots + tgCffCntr(ilNewCff).CffRec.iDay(ilDay)
                                                    Next ilDay
                                                Else
                                                    ilSpots = tgCffCntr(ilNewCff).CffRec.iSpotsWk
                                                End If
                                                If (ilSpots = 1) And (tgCffCntr(ilNewCff).CffRec.lActPrice + (llPPrice / ilSpots) >= 0) Then
                                                    If (tgCffCntr(ilNewCff).CffRec.lPropPrice > 0) Or (Not blPropPriceExist) Then
                                                        tgCffCntr(ilNewCff).CffRec.lActPrice = tgCffCntr(ilNewCff).CffRec.lActPrice + (llPPrice / ilSpots)
                                                        llPPrice = 0
                                                    End If
                                                End If
                                                Exit Do
                                            End If
                                            ilNewCff = tgCffCntr(ilNewCff).iNextCff
                                        Loop
                                    Next ilLoop
                                End If
                            End If
                        Else
                            tgCffCntr(ilCff).CffRec.lActPrice = 0
                            ilNewCff = tgClfCntr(ilLoop - 1).iFirstCff
                            Do While ilNewCff <> -1
                                gUnpackDateLong tgCffCntr(ilNewCff).CffRec.iStartDate(0), tgCffCntr(ilNewCff).CffRec.iStartDate(1), llDate
                                If llStartDate = llDate Then
                                    tgCffCntr(ilNewCff).CffRec.lActPrice = 0
                                    Exit Do
                                End If
                                ilNewCff = tgCffCntr(ilNewCff).iNextCff
                            Loop
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                        For ilLoop = ilSvLnRowNo + 1 To UBound(tgClfCntr) Step 1
                            mInitLnTotals ilLoop, imLnTotalType, PRICEINDEX
                        Next ilLoop
                    Loop
                End If
                '12/17/19: Handle TTP 9658
                mResortLines hmMnf
                imExpandingPkg = False

                mLnTotals ilSvLnRowNo, imLnTotalType, PRICEINDEX, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), "N", imNoWkDays()
                
                mAvailTab
                Contract.mLnSetShow -1, True
                Contract.pbcLine.Cls
                'mResortLines hmMnf
                Screen.MousePointer = vbHourglass
                imLnRowNo = -1
                Contract.pbcLnDates.Cls
                'Contract.mCntrTotal    'This calls Contract.pbcLnDates_Paint
                'Demo bar is made invisible so the code below is not required as imLnRowNo set to -1 in GotFocus and mSetCommands called
                'If tgUrf(0).sUseComputeCMC <> "I" Then
                '    mResetDemoBar True, True
                'End If
                mSetSchStatus smOrigStatus
                'mSetPriDemoResearch hmDrf, hmMnf, imStep1Index
                If (sgCntrScreen = "Orders") Or ((Asc(tgSpf.sOptionFields) And OFRESEARCH) <> OFRESEARCH) Then
                    Contract.lacAudTotals.Visible = False
                Else
                    If (tgSpf.sGUsePropSys = "Y") Then
                        Contract.lacAudTotals.Visible = True
                    Else
                        Contract.lacAudTotals.Visible = False
                    End If
                End If
                imBypassLnPaint = False
                imLnRowNo = ilSvLnRowNo
                If imLnSave(23, imLnRowNo) <> -1 Then
                    mPackageInfoToHidden imLnRowNo, 1
                End If
                If imLnSave(22, imLnRowNo) <> -1 Then
                    mPackageInfoToHidden imLnRowNo, 2
                End If
                If imLnSave(20, imLnRowNo) <> -1 Then
                    mPackageInfoToHidden imLnRowNo, 3
                End If
                If imLnSave(21, imLnRowNo) <> -1 Then
                    mPackageInfoToHidden imLnRowNo, 4
                End If
                If Trim$(smLnSave(28, imLnRowNo)) <> "" Then
                    mPackageInfoToHidden imLnRowNo, 7
                End If
                mGetFlightDays imLnRowNo, imOvDays()
                Contract.mSetOv2Flag imLnRowNo
                mSetDemoBarBooks imLnRowNo, ilFound, imCppCpm, 2, imCntrVefCode, imVpfIndex, imTerminate, imChgMode 'Using ilfound for ilDnfCode (mBuildWkInfo sets tgClf to DnfCode)
                'mGetResearchTotals Was commented out 9/26/97
                'mGetResearchTotals False
                '12/11/19
                If (Contract.plcDemoBar.Visible Or ((Asc(tgSaf(0).sFeatures6) And CALCULATESEARCHONLINECHG) = CALCULATESEARCHONLINECHG)) Then
                    mGetResearchTotals imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, imLnTotalType
                End If
                Contract.mBuildWkInfo True, True, False
                Contract.mCntrTotal    'This calls Contract.pbcLnDates_Paint
                Contract.pbcDemoBar(0).Cls
                Contract.pbcDemoBar(1).Cls
                Contract.pbcDemoBar_Paint 0
                Contract.pbcDemoBar_Paint 1
                'If UBound(tgClfCntr) <= vbcLines.LargeChange Then 'was <=
                '    vbcLines.Max = LBound(tgClfCntr)
                'Else
                '    vbcLines.Max = UBound(tgClfCntr) - vbcLines.LargeChange ' - 1
                'End If
                imExpandingPkg = False
                'mBuildMapLinePaint
                Contract.mSetLineScroll True, False
                Contract.pbcLine_Paint
                Contract.pbcLnTotal_Paint
            End If
    '        Exit For
        End If
    'Next ilVef
    igFreezeState = ilFreezeState
    imExpandingPkg = False
    ilLnRowNo = ilSvLnRowNo
    Screen.MousePointer = vbDefault
    Exit Function
TotalErr:
    mChkForStdPkgSub = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mLnTestFields                   *
'*                                                     *
'*             Created:8/10/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Public Function mLnTestFieldsFunction(VEHINDEX As Integer, PRICEINDEX As Integer, DPINDEX As Integer, OV2INDEX As Integer, WKSPOTSINDEX As Integer, LENINDEX As Integer) As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    'Dim ilIndex As Integer
    Dim ilAnySpots As Integer
    Dim ilClf As Integer
    Dim ilDayDefined As Integer
    Dim ilLegalDays As Integer
    Dim ilCff As Integer
    Dim slStr As String
    Dim ilSpotWarning As Integer
    Dim ilSvLnRowNo As Integer
    Dim slVehName As String
    Dim slDPName As String
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim ilAnyDays As Integer
    Dim ilDay As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim llPrefSTime As Long
    Dim llPrefETime As Long
    Dim ilMktCode As Integer    'Market Code from vehicle to match
    Dim slVefType As String    'Market Code from vehicle to match
    Dim ilIndex As Integer
    Dim ilRdf As Integer
    Dim ilOk As Integer
    Dim llNowDate As Long
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim llAllowedChgWkStartDate As Long
    Dim llEndPrevPeriod As Long
    '10/20/15: Check if vehicle type can be switched
    Dim slLineType As String
    '1/11/16: allow conventional vehicle name used as package vehicles
    Dim blFound As Boolean
    Dim ilPkg As Integer
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long

    ilSvLnRowNo = imLnRowNo
    ilSpotWarning = -1
    ilMktCode = -1
    imCntrCluster = False
    slVefType = ""
    imCntrRep = False
    imCntrAir = False

    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))

    If (Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS Then  'Allocation
        Contract.mPopRegion
    End If

    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        slVehName = ""
        'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgClfCntr(ilClf).ClfRec.iVefCode = tgMVef(ilLoop).iCode Then
            ilLoop = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilLoop <> -1 Then
                slVehName = " " & Trim$(tgMVef(ilLoop).sName)
                If (sgMktClusterDef = "Y") Or (sgRepDef = "Y") Then
                    If (tgClfCntr(ilClf).ClfRec.sType = "S") Or (tgClfCntr(ilClf).ClfRec.sType = "H") Then
                        'For ilCff = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
                        '    If tgMktMnf(ilCff).iCode = tgMVef(ilLoop).iMnfVehGp3Mkt Then
                                If slVefType = "" Then
                                    slVefType = tgMVef(ilLoop).sType
                         '           slRep = Trim$(tgMktMnf(ilCff).sRPU)
                         '           If (slRep <> "Y") And (slVefType = "R") Then
                                    If slVefType = "R" Then
                                        imCntrRep = True
                                    'ElseIf (slRep = "Y") And ((slVefType = "C") Or (slVefType = "S") Or (slVefType = "R")) Then
                         '           ElseIf (slRep = "Y") And (slVefType = "R") Then
                         '               imCntrCluster = True
                                    Else
                                        imCntrAir = True
                                    End If
                                Else
                         '           If (Trim$(tgMktMnf(ilCff).sRPU) <> "Y") And (tgMVef(ilLoop).sType = "R") Then
                                    If tgMVef(ilLoop).sType = "R" Then
                                        '8/25/20: ignore site option
                                        'If (Asc(tgSpf.sUsingFeatures2) And MIXAIRTIMEANDREP) <> MIXAIRTIMEANDREP Then
                                            'If (Not imCntrCluster) And (Not imCntrRep) Then
                                            If imCntrAir Then
                                                Screen.MousePointer = vbDefault    'Default
                                                Contract.mActiveLnTab
                                                ilRes = MsgBox("All Vehicle must be same type, Line #" & str$(tgClfCntr(ilClf).ClfRec.iLine) & ", " & slVehName, vbOKOnly + vbExclamation, "Incomplete")
                                                imLnRowNo = ilClf + 1
                                                imLnBoxNo = VEHINDEX
                                                mLnTestFieldsFunction = NO
                                                Exit Function
                                            End If
                                        'End If
                                        imCntrRep = True
                                    ElseIf tgMVef(ilLoop).sType <> "R" Then
                                        '8/25/20: ignore site option
                                        'If (Asc(tgSpf.sUsingFeatures2) And MIXAIRTIMEANDREP) <> MIXAIRTIMEANDREP Then
                                            If (imCntrCluster) Or (imCntrRep) Then
                                                Screen.MousePointer = vbDefault    'Default
                                                Contract.mActiveLnTab
                                                ilRes = MsgBox("All Vehicle must be same type, Line #" & str$(tgClfCntr(ilClf).ClfRec.iLine) & ", " & slVehName, vbOKOnly + vbExclamation, "Incomplete")
                                                imLnRowNo = ilClf + 1
                                                imLnBoxNo = VEHINDEX
                                                mLnTestFieldsFunction = NO
                                                Exit Function
                                            End If
                                        'End If
                                        imCntrAir = True
                                    End If
                                End If
                        '        Exit For
                        '    End If
                        'Next ilCff
                    End If
                End If
                '10/20/15: Check if vehicle types can be switched.  If line schedule, the switch is checked in mLnSetShow
                slLineType = mGetLineType(ilClf + 1)
                If tgMVef(ilLoop).sType = "P" Then
                    If (slLineType <> "A") And (slLineType <> "O") And (slLineType <> "E") Then
                        Screen.MousePointer = vbDefault    'Default
                        Contract.mActiveLnTab
                        ilRes = MsgBox("Line #" & str$(tgClfCntr(ilClf).ClfRec.iLine) & ", " & " Not converted to Package Line. Exit Price field to convert line", vbOKOnly + vbExclamation, "Incomplete")
                        imLnRowNo = ilClf + 1
                        imLnBoxNo = PRICEINDEX
                        mLnTestFieldsFunction = NO
                        Exit Function
                    End If
                Else
                    If (slLineType = "A") Or (slLineType = "O") Or (slLineType = "E") Then
                        '1/11/16: all vehicle if used as package name
                        blFound = False
                        For ilPkg = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                            If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilPkg).iLineNo Then
                                blFound = True
                                Exit For
                            End If
                        Next ilPkg
                        If Not blFound Then
                            Screen.MousePointer = vbDefault    'Default
                            Contract.mActiveLnTab
                            ilRes = MsgBox("Line #" & str$(tgClfCntr(ilClf).ClfRec.iLine) & ", " & " can not be converted to Standard Line. Trash package line or switch back to Package vehicle", vbOKOnly + vbExclamation, "Incomplete")
                            imLnRowNo = ilClf + 1
                            imLnBoxNo = VEHINDEX
                            mLnTestFieldsFunction = NO
                            Exit Function
                        End If
                    End If
                End If
        '        Exit For
            End If
        'Next ilLoop
        If tgClfCntr(ilClf).ClfRec.iRdfCode <= 0 Then
            Screen.MousePointer = vbDefault    'Default
            Contract.mActiveLnTab
            ilRes = MsgBox("Daypart Name must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName, vbOKOnly + vbExclamation, "Incomplete")
            imLnRowNo = ilClf + 1
            imLnBoxNo = DPINDEX
            mLnTestFieldsFunction = NO
            Exit Function
        End If
        slDPName = " " & mMakePrgName(tgClfCntr(ilClf).ClfRec.iRdfCode)
        If Not tgClfCntr(ilClf).iCancel Then
            If (tgClfCntr(ilClf).ClfRec.iStartTime(0) = 1) And (tgClfCntr(ilClf).ClfRec.iStartTime(1) = 0) Then
                If (tgClfCntr(ilClf).ClfRec.iEndTime(0) <> 1) Or (tgClfCntr(ilClf).ClfRec.iEndTime(1) <> 0) Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Start Time must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = DPINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            Else
                gUnpackTime tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), "A", "1", slStr
                If Not gValidTime(slStr) Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Invalid Start Time for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = DPINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            End If
            '12m = 0, 0; Blank = 1, 0
            If (tgClfCntr(ilClf).ClfRec.iEndTime(0) = 1) And (tgClfCntr(ilClf).ClfRec.iEndTime(1) = 0) Then
                If (tgClfCntr(ilClf).ClfRec.iStartTime(0) <> 1) Or (tgClfCntr(ilClf).ClfRec.iStartTime(1) <> 0) Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("End Time must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = DPINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            Else
                gUnpackTime tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), "A", "1", slStr
                If Not gValidTime(slStr) Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Invalid End Time for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = DPINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), False, llSTime
                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), True, llETime
            End If
            'Check Preferred Times
            gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iPrefStartTime(0), tgClfCntr(ilClf).ClfRec.iPrefStartTime(1), False, llPrefSTime
            gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iPrefEndTime(0), tgClfCntr(ilClf).ClfRec.iPrefEndTime(1), True, llPrefETime
            If llPrefSTime <> -1 Then
                If llPrefETime = -1 Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Preferred End Time must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = OV2INDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            Else
                If llPrefETime <> -1 Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Preferred Start Time must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = OV2INDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            End If
            If (llPrefSTime <> -1) And (llPrefETime <> -1) Then
                If llPrefETime < llPrefSTime Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Preferred End Time must be after Preferred Start Time for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = OV2INDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), False, llSTime
                gUnpackTimeLong tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), True, llETime
                If (llSTime <> -1) And (llETime <> -1) Then
                    If (llSTime > llPrefSTime) Or (llPrefETime > llETime) Then
                        Screen.MousePointer = vbDefault    'Default
                        Contract.mActiveLnTab
                        ilRes = MsgBox("Preferred Times must be a subset of the Override Times" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                        imLnRowNo = ilClf + 1
                        imLnBoxNo = OV2INDEX
                        mLnTestFieldsFunction = NO
                        Exit Function
                    End If
                Else
                    ilRdf = gBinarySearchRdf(tgClfCntr(ilClf).ClfRec.iRdfCode)
                    If ilRdf <> -1 Then
                        ilOk = False
                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llSTime
                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llETime
                                If (llSTime <= llPrefSTime) And (llPrefETime <= llETime) Then
                                    ilOk = True
                                    Exit For
                                End If
                            End If
                        Next ilIndex
                        If Not ilOk Then
                            Screen.MousePointer = vbDefault    'Default
                            Contract.mActiveLnTab
                            ilRes = MsgBox("Preferred Times must be a subset of the Daypart Allowed Times" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                            imLnRowNo = ilClf + 1
                            imLnBoxNo = OV2INDEX
                            mLnTestFieldsFunction = NO
                            Exit Function
                        End If
                    End If
                End If
            End If

            If tgClfCntr(ilClf).iFirstCff < 0 Then
                Screen.MousePointer = vbDefault    'Default
                Contract.mActiveLnTab
                ilRes = MsgBox("Dates must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                imLnRowNo = ilClf + 1
                If (tgClfCntr(ilClf).ClfRec.sType = "O") Or (tgClfCntr(ilClf).ClfRec.sType = "A") Then
                    imLnBoxNo = WKSPOTSINDEX
                Else
                    imLnBoxNo = PRICEINDEX
                End If
                mLnTestFieldsFunction = NO
                Exit Function
            End If
            'Check that at least one day defined (except in current week)
            ilDayDefined = False
            ilAnySpots = 0
            ilCff = tgClfCntr(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                    ilDayDefined = False
                    For ilLoop = 0 To 6 Step 1
                        If (tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilLoop) = "1") Then
                            ilDayDefined = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilDayDefined Then    'Allow weeks with no days defined as long as no spots requested
                        If tgCffCntr(ilCff).CffRec.iSpotsWk <> 0 Then
                            ilAnySpots = 1
                            Exit Do
                        Else
                            ilDayDefined = True
                        End If
                    End If
                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                        For ilLoop = 0 To 6 Step 1
                            ilAnySpots = ilAnySpots + tgCffCntr(ilCff).CffRec.iDay(ilLoop)
                        Next ilLoop
                    Else
                        ilAnySpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                    End If

                    If (tgClfCntr(ilClf).ClfRec.sType <> "O") And (tgClfCntr(ilClf).ClfRec.sType <> "A") Then
                        If (tgClfCntr(ilClf).iLineSchd) Then
                            If tgCffCntr(ilCff).CffRec.sDyWk <> "D" Then
                                If ((tgCffCntr(ilCff).lStartDate <= llNowDate) And (tgCffCntr(ilCff).lEndDate >= llNowDate)) Or (tgCffCntr(ilCff).lEndDate < tgCffCntr(ilCff).lStartDate) Then
                                    llAllowedChgWkStartDate = mLnDate(ilClf + 1, smLnLLD, 0)
                                    If llAllowedChgWkStartDate < tgClfCntr(ilClf).lEndBilledDate Then
                                        llAllowedChgWkStartDate = tgClfCntr(ilClf).lEndBilledDate
                                    End If
                                    gUnpackDateLong tgSpf.iRPRP(0), tgSpf.iRPRP(1), llEndPrevPeriod
                                    If (llAllowedChgWkStartDate < llEndPrevPeriod) Then
                                        llAllowedChgWkStartDate = llEndPrevPeriod + 1
                                    End If
                                    If (llAllowedChgWkStartDate >= tgCffCntr(ilCff).lStartDate) And (tgClfCntr(ilClf).iStatus > 0) And ((tgClfCntr(ilClf).ClfRec.sSchStatus = "F") Or (tgClfCntr(ilClf).ClfRec.sSchStatus = "I") Or (tgClfCntr(ilClf).ClfRec.sSchStatus = "A") Or (tgClfCntr(ilClf).ClfRec.sSchStatus = "N") Or (tgClfCntr(ilClf).ClfRec.sSchStatus = "P") Or (tgClfCntr(ilClf).ClfRec.sSchStatus = "M")) Then
                                        mGetCurrentWkNoSpots ilClf + 1, hmCHF, hmClf
                                        If tgClfCntr(ilClf).iMonThruToDaysSpots > 0 Then
                                            If ilAnySpots < tgClfCntr(ilClf).iMonThruToDaysSpots Then
                                                Screen.MousePointer = vbDefault    'Default
                                                Contract.mActiveLnTab
                                                ilRes = MsgBox("Cannot define spots less than " & Trim$(str$(tgClfCntr(ilClf).iMonThruToDaysSpots)) & " in current week as spots already aired", vbOKOnly + vbExclamation, "Incomplete")
                                                imLnRowNo = ilClf + 1
                                                imLnBoxNo = WKSPOTSINDEX
                                                mLnTestFieldsFunction = NO
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                End If
                'If ilDayDefined Then
                '    Exit Do
                'End If
                ilCff = tgCffCntr(ilCff).iNextCff
            Loop
            If (Not ilDayDefined) And (ilAnySpots > 0) Then
                Screen.MousePointer = vbDefault    'Default
                Contract.mActiveLnTab
                ilRes = MsgBox("Days must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                imLnRowNo = ilClf + 1
                If (tgClfCntr(ilClf).ClfRec.sType = "O") Or (tgClfCntr(ilClf).ClfRec.sType = "A") Then
                    imLnBoxNo = WKSPOTSINDEX
                Else
                    imLnBoxNo = PRICEINDEX
                End If
                mLnTestFieldsFunction = NO
                Exit Function
            End If
            If ilAnySpots <= 0 Then
                If (tgClfCntr(ilClf).ClfRec.sType = "O") Or (tgClfCntr(ilClf).ClfRec.sType = "A") Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Spots must be specified on Package Lines for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = WKSPOTSINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                Else
                    If ilSpotWarning = -1 Then
                        ilSpotWarning = ilClf + 1   'Save first line without spots
                    End If
                End If
            End If
            'Check that only legal days defined
            If (ilAnySpots > 0) And ((tgClfCntr(ilClf).sOwnership <> "B") And (tgClfCntr(ilClf).sOwnership <> "D")) Then
                mDPAllowedDays tgClfCntr(ilClf).ClfRec.iRdfCode, imAllowedDays()
                ilLegalDays = True
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llSDate
                        'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llEDate
                        llSDate = tgCffCntr(ilCff).lStartDate
                        llEDate = tgCffCntr(ilCff).lEndDate
                        If llEDate >= llSDate Then
                            For ilLoop = 0 To 6 Step 1
                                If (tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilLoop) = "1") Then
                                    If Not imAllowedDays(ilLoop + LBound(imAllowedDays)) Then
                                        ilLegalDays = False
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If (ilLegalDays) And (llSDate + 6 > llEDate) Then
                                ilAnyDays = False
                                For llDate = llSDate To llSDate + 6 Step 1
                                    ilDay = gWeekDayLong(llDate)
                                    If llDate <= llEDate Then
                                        If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                            ilAnyDays = True
                                        End If
                                    Else
                                        If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                            ilLegalDays = False
                                            Exit For
                                        End If
                                    End If
                                Next llDate
                                If Not ilAnyDays Then
                                    ilLegalDays = False
                                End If
                            End If
                        End If
                    End If
                    If Not ilLegalDays Then
                        Exit Do
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
                If Not ilLegalDays Then
                    Screen.MousePointer = vbDefault    'Default
                    Contract.mActiveLnTab
                    ilRes = MsgBox("Invalid days specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                    imLnRowNo = ilClf + 1
                    imLnBoxNo = WKSPOTSINDEX
                    mLnTestFieldsFunction = NO
                    Exit Function
                End If
            End If
        End If

        If tgClfCntr(ilClf).ClfRec.iLen <= 0 Then
            Screen.MousePointer = vbDefault    'Default
            Contract.mActiveLnTab
            ilRes = MsgBox("Length must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
            imLnRowNo = ilClf + 1
            imLnBoxNo = LENINDEX
            mLnTestFieldsFunction = NO
            Exit Function
        End If
        If tmOV1Ctrls(OVCOPYINDEX).iReq = True Then
            If Trim$(tgClfCntr(ilClf).ClfRec.sLiveCopy) = "" Then
                Screen.MousePointer = vbDefault    'Default
                Contract.mActiveLnTab
                ilRes = MsgBox("Copy Override must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                imLnRowNo = ilClf + 1
                imLnBoxNo = LENINDEX
                mLnTestFieldsFunction = NO
                Exit Function
            End If
        End If
        '"V" = Reservation
        If tgChfCntr.sType <> "V" Then
            If Not tgClfCntr(ilClf).iCancel Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        If (Trim$(tgCffCntr(ilCff).CffRec.sPriceType) = "") Then
                            Screen.MousePointer = vbDefault    'Default
                            Contract.mActiveLnTab
                            ilRes = MsgBox("Spot Price Type must be specified for Line" & str$(tgClfCntr(ilClf).ClfRec.iLine) & slVehName & slDPName, vbOKOnly + vbExclamation, "Incomplete")
                            imLnRowNo = ilClf + 1
                            If (tgClfCntr(ilClf).ClfRec.sType = "O") Or (tgClfCntr(ilClf).ClfRec.sType = "A") Then
                                imLnBoxNo = WKSPOTSINDEX
                            Else
                                imLnBoxNo = PRICEINDEX
                            End If
                            mLnTestFieldsFunction = NO
                            Exit Function
                        End If
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
            End If
        End If
        'Check that Regions are Ok
        If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) And (tgClfCntr(ilClf).ClfRec.lRafCode > 0) Then  'Allocation
            ilFound = False
            For ilLoop = 0 To UBound(tmRegionCode) - 1 Step 1
                slNameCode = tmRegionCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If tgClfCntr(ilClf).ClfRec.lRafCode = Val(slCode) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                Screen.MousePointer = vbDefault    'Default
                Contract.mActiveLnTab
                ilRes = MsgBox("Region definition missing for Advertiser", vbOKOnly + vbExclamation, "Incomplete")
                imLnRowNo = ilClf + 1
                imLnBoxNo = OV2INDEX
                mLnTestFieldsFunction = NO
                Exit Function
            End If
        End If
    Next ilClf
    For llSbf = 0 To UBound(tgIBSbf) - 1 Step 1
        If (tgIBSbf(llSbf).iStatus >= 0) Then
            If tgIBSbf(llSbf).SbfRec.iBillVefCode > 0 Then
                ilLoop = gBinarySearchVef(tgIBSbf(llSbf).SbfRec.iBillVefCode)
                If ilLoop <> -1 Then
                    slVehName = " " & Trim$(tgMVef(ilLoop).sName)
                    If slVefType = "" Then
                       slVefType = tgMVef(ilLoop).sType
                       If slVefType = "R" Then
                           imCntrRep = True
                       Else
                           imCntrAir = True
                       End If
                    Else
                        If tgMVef(ilLoop).sType = "R" Then
                            '8/25/20: ignore site option
                            'If (Asc(tgSpf.sUsingFeatures2) And MIXAIRTIMEANDREP) <> MIXAIRTIMEANDREP Then
                               If imCntrAir Then
                                   Screen.MousePointer = vbDefault    'Default
                                   Contract.mActiveLnTab
                                   ilRes = MsgBox("All Vehicle must be same type, NTR " & slVehName, vbOKOnly + vbExclamation, "Incomplete")
                                   imLnRowNo = ilClf + 1
                                   imLnBoxNo = VEHINDEX
                                   mLnTestFieldsFunction = NO
                                   Exit Function
                               End If
                           'End If
                           imCntrRep = True
                        ElseIf tgMVef(ilLoop).sType <> "R" Then
                            '8/25/20: ignore site option
                            'If (Asc(tgSpf.sUsingFeatures2) And MIXAIRTIMEANDREP) <> MIXAIRTIMEANDREP Then
                               If (imCntrCluster) Or (imCntrRep) Then
                                   Screen.MousePointer = vbDefault    'Default
                                   Contract.mActiveLnTab
                                   ilRes = MsgBox("All Vehicle must be same type, NTR " & slVehName, vbOKOnly + vbExclamation, "Incomplete")
                                   imLnRowNo = ilClf + 1
                                   imLnBoxNo = VEHINDEX
                                   mLnTestFieldsFunction = NO
                                   Exit Function
                               End If
                           'End If
                           imCntrAir = True
                       End If
                   End If
                End If
            End If
        End If
    Next llSbf
    If ilSpotWarning > 0 Then
        Screen.MousePointer = vbDefault    'Default
        Contract.mActiveLnTab
        ilRes = MsgBox("Lines without Spots, Continue with Save", vbOKCancel + vbQuestion, "Warning")
        If ilRes = vbCancel Then
            imLnRowNo = ilSpotWarning
            If (tgClfCntr(ilClf).ClfRec.sType = "O") Or (tgClfCntr(ilClf).ClfRec.sType = "A") Then
                imLnBoxNo = WKSPOTSINDEX
            Else
                imLnBoxNo = PRICEINDEX
            End If
            mLnTestFieldsFunction = NO
            Exit Function
        End If
    End If
    Screen.MousePointer = vbHourglass    'Default
    imLnRowNo = ilSvLnRowNo
    mLnTestFieldsFunction = YES
End Function

Public Sub mInitLoadSub(LINECOMMINDEX As Integer, OV1INDEX As Integer, SPOTSINDEX As Integer)
'
'   mInitLoad
'   Where:
'
    Dim ilLoop As Integer
    
    '5/18/12: Only include dormant salesperson if referenced in contract
    Dim llSlf As Long
    Dim ilVff As Integer
    Dim ilRet As Integer    'Return Status
    Dim slCode As String
    Dim tlCff As CFF    'Only required to obtain length
    Dim tlCgf As CGF
    Dim tlMsf As MSF
    Dim tlMgf As MGF
    ReDim tgClfCntr(0 To 0) As CLFLIST
    ReDim tgDelClf(0 To 0) As CLFLIST
    
    bmLnChgRegenResearch = False
    bmMixturePopulationMsgShown = False
    imLB1Or2 = 1
    tgClfCntr(0).iFirstCff = -1
    tgClfCntr(0).iFirstCgf = -1
    tgClfCntr(0).lFirstAvail = -1
    tgClfCntr(0).iStatus = -1
    ReDim tgCffCntr(0 To 0) As CFFLIST
    tgCffCntr(0).iStatus = -1 'Not Used
    tgCffCntr(0).lRecPos = 0
    tgCffCntr(0).iNextCff = -1
    imLastCffUsed = 0
    ReDim tgCgfCntr(0 To 0) As CGFLIST
    tgCgfCntr(0).iStatus = -1 'Not Used
    tgCgfCntr(0).iNextCgf = -1
    ReDim tgMsfCntr(0 To 0) As MSFLIST
    tgMsfCntr(0).iStatus = -1 'Not Used
    tgMsfCntr(0).iFirstMgf = -1
    tgMsfCntr(0).iCxfIndex = -1
    ReDim tgMgfCntr(0 To 0) As MGFLIST
    tgMgfCntr(0).iStatus = -1
    ReDim tgLnAvailInfo(0 To 0) As LNAVAILINFO
    tgLnAvailInfo(0).lNextIndex = -1
    ReDim smLnComment(0 To 1) As String
    ReDim smLnInitComment(0 To 1) As String
    ReDim imLnCommShowOn(0 To 4, 0 To 1) As Integer
    ReDim imLnInitCommShowOn(0 To 4, 0 To 1) As Integer
    ReDim tmWkInfo(0 To 1) As WKINFO
    ReDim lmDiffCntrCode(0 To 1) As Long
    ReDim tgPackVeh(0 To 0) As PACKVEH
    ReDim tgPackDP(0 To 0) As PACKDP
    ReDim tgPackLn(0 To 0) As PACKLN
    ReDim igLnMakingPack(0 To 0) As Integer
    ReDim tgCShtTitle(0 To 0) As CSHTTITLE
    ReDim tmRvfMPRec(0 To 0) As RVFMPREC
    ReDim tgFBSbf(0 To 0) As SBFLIST
    ReDim tmRegionCode(0 To 0) As SORTCODE
    smRegionCodeTag = ""
    ReDim smLnSave(0 To 29, 0 To UBound(tgClfCntr) + 1) As String
    ReDim imLnSave(0 To 35, 0 To UBound(tgClfCntr) + 1) As Integer
    ReDim smLnShow(0 To LINECOMMINDEX, 0 To UBound(tgClfCntr) + 1) As String
    ReDim smAirShow(0 To 14, 0 To 1) As String
    ReDim smAirSave(0 To 13, 0 To 1) As String
    ReDim imAirSave(0 To 2, 0 To 1) As Integer

    ReDim tmVehicleCode(0 To 0) As SORTCODE

    tgFBSbf(0).iStatus = -1 'Not Used
    tgFBSbf(0).lRecPos = 0
    ReDim tgIBSbf(0 To 0) As SBFLIST
    tgIBSbf(0).iStatus = -1 'Not Used
    tgIBSbf(0).lRecPos = 0
    ReDim tgMBSbf(0 To 0) As SBFLIST
    tgMBSbf(0).iStatus = -1 'Not Used
    tgMBSbf(0).lRecPos = 0
    ReDim tgPBSbf(0 To 0) As SBFLIST
    tgPBSbf(0).iStatus = -1 'Not Used
    tgPBSbf(0).lRecPos = 0
    ReDim lgDeletedPcfCode(0 To 0) As Long
    imLBCTCtrls = 1
    
    If tgSpf.sSubCompany = "Y" Then
        imSPNoCol = 4
        ReDim tmSPCtrls(0 To 40) As FIELDAREA
        Contract.pbcSPerson.Move Contract.pbcSPerson.Left, Contract.pbcSPerson.Left, 6750
    ElseIf tgSpf.sCommByCntr = "Y" Then
        imSPNoCol = 3
        ReDim tmSPCtrls(0 To 30) As FIELDAREA
        Contract.pbcSPerson.Move Contract.pbcSPerson.Left, Contract.pbcSPerson.Left, 4800
    Else
        imSPNoCol = 2
        ReDim tmSPCtrls(0 To 20) As FIELDAREA
        Contract.pbcSPerson.Move Contract.pbcSPerson.Left, Contract.pbcSPerson.Left, 3825
    End If
    ReDim tgMapLinePaint(0 To 0) As MAPLINEPAINT
    Screen.MousePointer = vbHourglass

    Contract.plcScreen.Caption = sgCntrScreen

    igJobShowing(CONTRACTSJOB) = True
    smSvAllowDailyWeekly = tgSpf.sAllowDailyBuys
    imDeactivated = False
    igDPCallSource = 0
    Contract.pbcHeader.Picture = Contract.pbcHdImage(0).Picture
    Contract.pbcLogo.Picture = Traffic!pbcLogo.Picture
    Contract.imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    Contract.imcKey.Picture = IconTraf!imcKey.Picture
    Contract.imcShare(0).Picture = IconTraf!imcPrinter.Picture
    Contract.imcShare(1).Picture = IconTraf!imcExport.Picture
    'imcLnArrow.Picture = IconTraf!imcArrowUp.Picture
'    Contract.pbcWkArrow.Picture = IconTraf!imcArrow.Picture
'    Contract.pbcWkArrow.Width = 90
'    Contract.pbcWkArrow.Height = 165
    Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
    Contract.pbcLineArrow.Width = 90
    Contract.pbcLineArrow.height = 165
    Contract.mSetTabColors True
    Contract.mInitBox
    'Dan M 11/28/11 add email button to proposal screen
    Contract.mCreateEmailButton
    imListField(1) = 15
    imListField(2) = 9 * igAlignCharWidth '6
    imListField(3) = 20 * igAlignCharWidth '20
    imAllClicked = False
    imSetAll = True
    Contract.cmcMkt.Caption = "&Market"
    If (sgCntrScreen <> "Orders") And ((igAlertCntrStatus = 1) Or (igAlertCntrStatus = 2) Or (igAlertCntrStatus = 3)) Then
        Contract.cmcMkt.Caption = "&Alert"
        Contract.cmcMkt.Visible = True
    ElseIf (sgCntrScreen = "Orders") And ((igAlertCntrStatus = 4) Or (igAlertCntrStatus = 5)) Then
        Contract.cmcMkt.Caption = "&Alert"
        Contract.cmcMkt.Visible = True
    ElseIf (sgCntrScreen = "Orders") And ((igAlertCntrStatus = 6) Or (igAlertCntrStatus = 7)) Then
        Contract.cmcMkt.Caption = "&Alert"
        Contract.cmcMkt.Visible = False
    Else
        If tgSpf.sMktBase = "Y" Then
            Contract.cmcMkt.Visible = True
        End If
    End If
    'TTP 10937 - Contract Dashboard: new screen to show contract header information
    If (sgCntrScreen <> "Orders") And ((igDashboardCntrStatus = 1) Or (igDashboardCntrStatus = 2) Or (igDashboardCntrStatus = 3)) Then
        Contract.cmcMkt.Caption = "&Dash"
        'Contract.cmcMkt.Visible = True
    ElseIf (sgCntrScreen = "Orders") And ((igDashboardCntrStatus = 4) Or (igDashboardCntrStatus = 5)) Then
        Contract.cmcMkt.Caption = "&Dash"
        'Contract.cmcMkt.Visible = True
    ElseIf (sgCntrScreen = "Orders") And ((igDashboardCntrStatus = 6) Or (igDashboardCntrStatus = 7)) Then
        Contract.cmcMkt.Caption = "&Dash"
        Contract.cmcMkt.Visible = False
    Else
        If tgSpf.sMktBase = "Y" Then
            Contract.cmcMkt.Visible = True
        End If
    End If
    
    bgCntrFirstCallToVpfFind = True
    igPackByWkAvg = True
    imInChkForStdPkg = False
    imBypassLnPaint = False
    bmInAvailTab = False
    imTerminate = False
    imPopReqd = False
    imFirstActivate = True
    imSelectDelay = False
    imSelectUpdate = False
    imStartMode = True
    imSettingCD = False
    imStepBoxNo = -1
    imStep1Index = -1
    imStep2Index = -1
    imStep3Index = -1
    imInSelect = False
    imButtonIndex = -1
    imFirstTimeStep = True
    lmLockRecCode = -1
    imSumType = 0
    imSaveStay = False
    imCHFRecLen = Len(tgChfCntr)     'Get and save CHF record length
    imHdBoxNo = -1 'Initialize current Box to N/A
    imLnBoxNo = -1
    imNoBoxNo = -1
    imSPBoxNo = -1
    imAirBoxNo = -1
    imDGBoxNo = -1
    imCEBoxNo = -1
    imDMBoxNo = -1
    imPTBoxNo = -1
    imRSBoxNo = -1
    imSBBoxNo = -1
    imDBBoxNo = -1
    imDBRowNo = -1
    imDB0BoxNo = -1
    igGameSchdVefCode = 0
    imExpandingPkg = False
    imInTseLine = False
    imInQtPrice = False
    imComputeReq = False
    igShowHelpMsg = True
    imInLnComment = False
    imLnTotalType = 0
    imBypassResetDemoBar = False
    imToggleCPPCPM = 0
    imToggleAvail = -1 'Avail = None
    gSetCtrl tmAvToggleCtrls(1), 30, 405, 1380, 180
    imOVDGenerated = False
    Contract.lacPrice(0).Caption = "Average Price"  ' & sgLF & sgCR & "Price"
    Contract.pbcSumType_Paint (1)
    If (Asc(tgSaf(0).sFeatures3) And FREEZEDEFAULT) = FREEZEDEFAULT Then
        igFreezeState = 1
        Contract.lacPrice(1).Caption = "Freeze Calculation"
    Else
        igFreezeState = 0
        Contract.lacPrice(1).Caption = "Unfreeze Calculation"
    End If
    bgSortByLineNo = False
    igPackLnNo = -1
    imDirProcess = -1
    imRBDetailAnchorRow = -1
    imUndoSelection = -1
    imReturnToAlerts = False
    imMultiMediaVefCode = -1
    smMultiMediaTypeItem = ""
    imCurrHightLightLnRowNo = -1
    imPrevHightLightLnRowNo = -1
    imPaintLnRowNo = -1
'    cbcRBView.AddItem "Specifications", 0
'    cbcRBView.AddItem "Vehicle Summary", 1
'    cbcRBView.AddItem "Daypart Summary", 2
    imAskBookQ = False
    imInCommentBranch = False
    imInBuildWkInfo = False
    imDiffMode = False
    imAllClicked = False
    imTabDirection = 0  'Left to right movement
    imInSave = False
    imLbcMouseDown = False
    imChgMode = False
    imVefDormantAdd = True
    bmCheckSlfDormant = True
    imDGChgMode = False
    imRBView = 0
    imPBChgMode = False
    imPBDateIndex = -1
    imDBRatio = 0   'Actual/Rate Card
    Contract.pbcRatio.Cls
    Contract.pbcRatio.CurrentX = -15
    Contract.pbcRatio.CurrentY = 0
    Contract.pbcRatio.Print "Actual/Rate Card"
    lmTabColor(0) = -1
    imDBTotals = 0
    'lacDBTotal.Caption = "Line Totals"
    imBSMode = False
    imAirCancel = False
    imLnChg = False
    imGameInvChg = False
    imNTRChg = False
    '10910
    bgContractAttachments_Changed = False
    imCPMChg = False
    imAllAnsw = False
    imAnyFixedBilled = False
    imAnyLnSchd = False
    imCreditRestr = -1  'No restrictions
    imBypassFocus = False
    smPhoneImage = Contract.mkcPhone.Text
    smAdfStamp = ""
    smAgfStamp = ""
    smPctTrade = ""
    lmChfRecPos = 0
    lmCntrChfCode = 0
    For ilLoop = 0 To 5
        lmCxfRecPos(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(imDetailSortIndex) To UBound(imDetailSortIndex) Step 1
        imDetailSortIndex(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(imSumSortIndex) To UBound(imSumSortIndex) Step 1
        imSumSortIndex(ilLoop) = 0
    Next ilLoop
    lmCmfRecPos = 0
    smRateCardPrice = ""
    smHighestPrice = ""
    smAvgPrice = ""
    smLowestPrice = ""
    imTotColor = 2
    imCreditColor = 2
    imStatsColor = 2
    imVpfIndex = -1
    imCntrVefCode = 0
    imDelaySrce = -1
    imSvLnRowNo = -1
    bmInCancel = False
    imInFlightScreen = False
    bmSettingPriceInTab = False
    imInResearch = False
    imInSummary = False
    lmHSStartDate = -1
    imHSStartWkNo = 0
    imHSEndWkNo = 0
    imHSCurWkNo = 0
    imInHotSpot = False
    imBypassEPriceComp = True
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    lmMonNowDate = gDateValue(gObtainPrevMonday(Format$(gNow(), "m/d/yy")))
    lmFirstSchMonDate = gDateValue(Format$(gNow(), "m/d/yy"))
    If gWeekDayLong(lmFirstSchMonDate) <> 6 Then
        lmFirstSchMonDate = lmFirstSchMonDate - gWeekDayLong(lmFirstSchMonDate)
    Else
        lmFirstSchMonDate = lmFirstSchMonDate + 1
    End If
    'imTestTypeIndexSetting = True
    imExtraRcfAdded = False
    imLbcArrowSetting = False
    imDoubleClickName = False
    imPBCDoubleClick = False
    hmCHF = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", Contract
    On Error GoTo 0
    hmPassChf = hmCHF
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", Contract
    On Error GoTo 0
    hmClf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", Contract
    On Error GoTo 0
    imClfRecLen = Len(tgClfCntr(0).ClfRec) 'btrRecordLength(hmClf)     'Get Clf size
    hmCff = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", Contract
    On Error GoTo 0
    imCffRecLen = Len(tlCff) 'btrRecordLength(hmCff)    'Get Cff size
    
    hmGhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ghf.Btr)", Contract
    On Error GoTo 0
    hmGsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Gsf.Btr)", Contract
    On Error GoTo 0
    hmCgf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", Contract
    On Error GoTo 0
    imCgfRecLen = Len(tlCgf) 'btrRecordLength(hmCff)    'Get Cff size
    hmEff = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmEff, "", sgDBPath & "Eff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Eff.Btr)", Contract
    On Error GoTo 0
    imEffRecLen = Len(tmEff) 'btrRecordLength(hmCff)    'Get Cff size
    hmMsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmMsf, "", sgDBPath & "Msf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Msf.Btr)", Contract
    On Error GoTo 0
    imMsfRecLen = Len(tlMsf) 'btrRecordLength(hmCff)    'Get Cff size
    hmMgf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmMgf, "", sgDBPath & "Mgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mgf.Btr)", Contract
    On Error GoTo 0
    imMgfRecLen = Len(tlMgf) 'btrRecordLength(hmCff)    'Get Cff size
    hmCxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cxf.Btr)", Contract
    On Error GoTo 0
    hmCmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCmf, "", sgDBPath & "Cmf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cmf.Btr)", Contract
    On Error GoTo 0
    hmVLF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", Contract
    On Error GoTo 0
    hmCef = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cef.Btr)", Contract
    On Error GoTo 0
    imCefRecLen = Len(tmCef) 'btrRecordLength(hmCff)    'Get Cff size
    hmAdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", Contract
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf) 'btrRecordLength(hmAdf)    'Get Cff size
    hmAgf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", Contract
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf) 'btrRecordLength(hmAgf)    'Get Cff size
    hmPrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Prf.Btr)", Contract
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf) 'btrRecordLength(hmAdf)    'Get Cff size
    hmSif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", Contract
    On Error GoTo 0
    imSifRecLen = Len(tmSif) 'btrRecordLength(hmAdf)    'Get Cff size
    hmMnf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mnf.Btr)", Contract
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf) 'btrRecordLength(hmMnf)    'Get Cff size
    hmDrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Drf.Btr)", Contract
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf) 'btrRecordLength(hmDrf)    'Get Cff size
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", Contract
    On Error GoTo 0
    'imDpfRecLen = Len(tmDpf) 'btrRecordLength(hmDrf)    'Get Cff size
    lgDpfNoRecs = btrRecords(hmDpf)
    hmDef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", Contract
    On Error GoTo 0
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", Contract
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)
    hmPnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pnf.Btr)", Contract
    On Error GoTo 0
    imPnfRecLen = Len(tmPnf) 'btrRecordLength(hmPnf)    'Get Cff size
    hmPvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pvf.Btr)", Contract
    On Error GoTo 0
    imPvfRecLen = Len(tmPvf) 'btrRecordLength(hmAdf)    'Get Cff size

    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Or ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        hmGif = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmGif, "", sgDBPath & "Gif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Gif.Btr)", Contract
        On Error GoTo 0
        imGifRecLen = Len(tmGif) 'btrRecordLength(hmAdf)    'Get Cff size
    End If

    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", Contract
    On Error GoTo 0

    If tgUrf(0).iRemoteUserID <= 0 Then
        hmSdf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", Contract
        On Error GoTo 0
        imSdfRecLen = Len(tmSdf) 'btrRecordLength(hmSdf)    'Get Cff size
        hmPsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmPsf, "", sgDBPath & "Psf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Psf.Btr)", Contract
        On Error GoTo 0
        hmSmf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", Contract
        On Error GoTo 0
        imSmfRecLen = Len(tmSmf) 'btrRecordLength(hmSmf)    'Get Cff size
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        hmLcf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmLcf, "", sgDBPath & "LCF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: LCF.Btr)", Contract
        On Error GoTo 0
        imLcfRecLen = Len(tmLcf) 'btrRecordLength(hmLcf)    'Get Cff size
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        hmRvf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmRvf, "", sgDBPath & "RVF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: RVF.Btr)", Contract
        On Error GoTo 0
        hmPhf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmPhf, "", sgDBPath & "PHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: PHF.Btr)", Contract
        On Error GoTo 0
    End If
    hmVsf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "VSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: VSF.Btr)", Contract
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf) 'btrRecordLength(hmVsf)    'Get Cff size
    If tgUrf(0).iRemoteUserID <= 0 Then
        hmSsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSsf, "", sgDBPath & "SSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: SSF.Btr)", Contract
        On Error GoTo 0
    End If
    If tgUrf(0).iRemoteUserID > 0 Then
        hmAsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmAsf, "", sgDBPath & "ASF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ASF.Btr)", Contract
        On Error GoTo 0
    End If
    hmVbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVbf, "", sgDBPath & "Vbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vbf.Btr)", Contract
    On Error GoTo 0
    imVBfRecLen = Len(tmVbf) 'btrRecordLength(hmAdf)    'Get Cff size


    '5/18/12: Only include dormant salesperson if referenced in contract
    'Moved from mInitActive
    Screen.MousePointer = vbHourglass
    
    mPopListKey
    
    ilRet = gObtainSalesperson()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    tgChfCntr.lCode = 0
    
    Contract.lbcSPerson(0).Clear 'Force list box to be populated
    Contract.mSPersonPop
    
    '5/18/12: Only include dormant salesperson if referenced in contract
    ReDim tmSvCntSalesperson(LBound(tmCntSalesperson) To UBound(tmCntSalesperson)) As SORTCODE
    For llSlf = LBound(tmCntSalesperson) To UBound(tmCntSalesperson) Step 1
        tmSvCntSalesperson(llSlf) = tmCntSalesperson(llSlf)
    Next llSlf
    
    Contract.mClearCtrlFields
    Contract.mClearPB
    For ilLoop = LBound(smDefSave) To UBound(smDefSave) Step 1
        smDefSave(ilLoop) = ""
        imDefSource(ilLoop) = -1
    Next ilLoop
    smFBITotal = "0"    'Fixed Price amount
    smIBCTotal = "0"    'Item contracted amount
    smTotalNA = "                                      "
    Contract.lacLnTotals(0).Caption = smTotalNA    'Same # blanks as in mCntrTotal anf cbcFilter
    'gHlfRead "CHF", tmChfHelp()
    'gHlfRead "CLF", tmClfHelp()
    smVehGp5CodeTag = ""
    ilRet = gPopMnfPlusFieldsBox(Contract, Contract.lbcAdvt, tmVehGp5Code(), smVehGp5CodeTag, "H5")
    Screen.MousePointer = vbHourglass
    ReDim tmRatedGp(LBound(tmVehGp5Code) To UBound(tmVehGp5Code)) As RATEDGP
    For ilLoop = LBound(tmVehGp5Code) To UBound(tmVehGp5Code) - 1 Step 1
        ilRet = gParseItem(tmVehGp5Code(ilLoop).sKey, 2, "\", slCode)
        tmRatedGp(ilLoop).iMnfVehGp5Rsch = Val(slCode)
        tmMnfSrchKey.iCode = tmRatedGp(ilLoop).iMnfVehGp5Rsch
        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If StrComp(Trim$(tmMnf.sRPU), "N", 1) = 0 Then
                tmRatedGp(ilLoop).sRated = "N"
            Else
                tmRatedGp(ilLoop).sRated = "Y"
            End If
        Else
            tmRatedGp(ilLoop).sRated = "Y"
        End If
    Next ilLoop
    smVehGp6CodeTag = ""
    Contract.lbcVehGp6.Clear
    Contract.mVehGp6Pop
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass  'Wait
    gAnyClustersDef
    gAnyRepDef
    Screen.MousePointer = vbHourglass

    Contract.lbcAdvt.Clear 'Force list box to be populated
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcAgency.Clear 'Force list box to be populated
    mAgencyPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcComp(0).Clear 'Force list box to be populated
    Contract.lbcComp(1).Clear
    mMnfPop 2
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcExcl(0).Clear 'Force list box to be populated
    Contract.lbcExcl(1).Clear 'Force list box to be populated
    mMnfPop 3
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcBusCat.Clear
    mMnfPop 1
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    Contract.lbcPot.Clear 'Force list box to be populated
    mMnfPop 4
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcSeg.Clear 'Force list box to be populated
    mMnfPop 5
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.mResvPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.mCommentPop True
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainRcfRifRdf()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = gObtainVef()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = gVffRead()
    bmAnyVehicleOnInsertion = False
    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then
        For ilVff = LBound(tgVff) To UBound(tgVff) - 1 Step 1
            If tgVff(ilVff).sOnInsertions <> "N" Then
                bmAnyVehicleOnInsertion = True
                Exit For
            End If
        Next ilVff
    End If
    smRBBooknameTag = ""
    ReDim tmRBBookname(0 To 0) As SORTCODE
    Contract.lbcRBBookName.Clear
    ilRet = Contract.mPopRBBooks()
    imSalesperson = False
    If (tgUrf(0).iSlfCode > 0) Then
        For ilLoop = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgMSlf(ilLoop).iCode = tgUrf(0).iSlfCode Then
                If StrComp(tgMSlf(ilLoop).sJobTitle, "S", 1) = 0 Then
                    imSalesperson = True
                End If
                Exit For
            End If
        Next ilLoop
    End If
    ilRet = gObtainUrf()
    Screen.MousePointer = vbHourglass
    Contract.lbcType.AddItem "Standard"
    If tgUrf(0).sResvType <> "H" Then
        Contract.lbcType.AddItem "Reservation"
    End If
    If tgUrf(0).sRemType <> "H" Then
        Contract.lbcType.AddItem "Remnant"
    End If
    If tgUrf(0).sDRType <> "H" Then
        Contract.lbcType.AddItem "Direct Response"
    End If
    If tgUrf(0).sPIType <> "H" Then
        Contract.lbcType.AddItem "Per Inquiry"
    End If
    If tgUrf(0).sPSAType <> "H" Then
        Contract.lbcType.AddItem "PSA"
    End If
    If tgUrf(0).sPromoType <> "H" Then
        Contract.lbcType.AddItem "Promo"
    End If
    Contract.lbcShowPrice.AddItem "Contract-No  Invoice-No"
    Contract.lbcShowPrice.AddItem "Contract-Yes Invoice-No"
    Contract.lbcShowPrice.AddItem "Contract-No  Invoice-Yes"
    Contract.lbcShowPrice.AddItem "Contract-Yes Invoice-Yes"
    smRevSetCodeTag(0) = ""
    smRevSetCodeTag(1) = ""
    smRevSetCodeTag(2) = ""
    smRevSetCodeTag(3) = ""
    smRevSetCodeTag(4) = ""
    Contract.lbcRevSet(0).Clear 'Force list box to be populated
    Contract.lbcRevSet(1).Clear 'Force list box to be populated
    Contract.lbcRevSet(2).Clear 'Force list box to be populated
    Contract.lbcRevSet(3).Clear 'Force list box to be populated
    Contract.lbcRevSet(4).Clear 'Force list box to be populated
    mRevSetPop 1
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    '3/6/18: Force load
    smShareBudgetCodeTag(0) = "~"
    smShareBudgetCodeTag(1) = "~"

    Contract.mShareBudgetPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gPopBooks
    mDemoPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mSocEcoPop
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Contract.lbcRateCard.Clear 'Force list box to be populated
    Contract.mRateCardPop
    If imTerminate Then
        imTerminate = False
        'Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    smBaseLen = ""
    If (tgUrf(0).iSlfCode <= 0) And (tgUrf(0).iRemoteUserID <= 0) Then
        'lbcBillCycle.AddItem "Broadcast"
        'If ((Asc(tgSpf.sOptionFields) And OFCALENDARBILL) = OFCALENDARBILL) Then
        '    lbcBillCycle.AddItem "Calendar"
        'End If
        'If ((Asc(tgSpf.sUsingFeatures9) And WEEKLYBILL) = WEEKLYBILL) Then
        '    lbcBillCycle.AddItem "Weekly"
        'End If
        Contract.mPopBillCycle False, False
        'Removed Fix on 12/27/02- Planning on using NTR instead
        'lbcBillCycle.AddItem "Fix B'cast"
        'lbcBillCycle.AddItem "Fix Cal"
        Contract.lbcInvGroup.AddItem "Only One"
        Contract.lbcInvGroup.AddItem "1 per Product"
        Contract.lbcInvGroup.AddItem "1 per Tag"
    Else
        Contract.lbcBillCycle.AddItem "Std B'cast"   '"Broadcast"
        Contract.lbcInvGroup.AddItem "Only One"
    End If
    Screen.MousePointer = vbHourglass
    gDayPop Contract.lbcDay
    '11/2/19:Moved to CFlight
    'Contract.lbcDGRun.AddItem "Week"
    'Contract.lbcDGRun.AddItem "2nd"
    'Contract.lbcDGRun.AddItem "3rd"
    'Contract.lbcDGRun.AddItem "4th"
    Contract.lbcDWP.Clear
    Contract.lbcDWP.AddItem "Daily"
    Contract.lbcDWP.AddItem "Weekly"
    'Contract.lbcDW.Clear
    'Contract.lbcDW.AddItem "Daily"
    'Contract.lbcDW.AddItem "Weekly"
    Contract.lbcPriceType(0).AddItem "[Dollars]"
    Contract.lbcPriceType(0).AddItem "Rate Card"
    Contract.lbcPriceType(0).AddItem "Avg Rate Card"
    Contract.lbcPriceType(0).AddItem "N/C"
    Contract.lbcPriceType(0).AddItem "MG"
    Contract.lbcPriceType(0).AddItem "Bonus"
    Contract.lbcPriceType(0).AddItem "Spinoff"
    Contract.lbcPriceType(0).AddItem "Package"
    Contract.lbcPriceType(0).AddItem "Recapturable"
    Contract.lbcPriceType(0).AddItem "ADU"
    Contract.lbcPriceType(1).AddItem "[Dollars]"
    Contract.lbcPriceType(1).AddItem "N/C"
    Contract.lbcPriceType(1).AddItem "MG"
    Contract.lbcPriceType(1).AddItem "Bonus"
    Contract.lbcPriceType(1).AddItem "Spinoff"
    Contract.lbcPriceType(1).AddItem "Package"
    Contract.lbcPriceType(1).AddItem "Recapturable"
    Contract.lbcPriceType(1).AddItem "ADU"
   '10843 remove dead code
'    If tgSpf.sCLnStdQt = "Y" Then
'        Contract.lbcQtPrice.AddItem "Average Price"
'        Contract.lbcQtPrice.AddItem "Total Price"
'        Contract.lbcQtPrice.AddItem "Quarter 1 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 2 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 3 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 4 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 5 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 6 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 7 Price"
'        Contract.lbcQtPrice.AddItem "Quarter 8 Price"
'        Contract.lbcQtPrice.height = gListBoxHeight(Contract.lbcQtPrice.ListCount, 10)
'        imLnTotalType = 2
'        Contract.lacPrice(0).Caption = Contract.lbcQtPrice.List(imLnTotalType)
'        Contract.lbcQtPrice.Move Contract.pbcLine.Left + Contract.pbcQTPrice.Left + Contract.pbcQTPrice.Width - Contract.lbcQtPrice.Width, Contract.pbcLine.Top + Contract.pbcQTPrice.Top - Contract.lbcQtPrice.height
'    End If
    Contract.height = Contract.cmcDone.Top + 5 * Contract.cmcDone.height / 3
    gCenterForm Contract
    'Traffic!plcHelp.Cls
    Contract.lbcStep(2).Clear  'Force list to be populated
    ReDim tmStep3Code(0 To 0) As SORTCODE   'lbcStep3Code.Clear
    smStep3CodeTag = "" 'lbcStep3Code.Tag = ""
    Contract.cbcDifference.Visible = False
'    mBuildXmlPaths smAgencyXMLPath, smStationXMLBasePath, smStationXMLWOPath
    If sgCntrScreen <> "Orders" Then
        Contract.cmcScale.Visible = True
        smAgencyXMLPath = mBuildAgencyXmlPath()
    Else
        '8033 no longer on order screen Handled in mSetBackgroundColor
       ' imcShare(1).Visible = False
    End If
    mStationXmlAllowedStart
    'will be blank if not allowed
    mBuildStationXMLPaths smStationXMLBasePath, smStationXMLWOPath
    Contract.cbcDifference.Clear
    imDiffMode = False
    Contract.mSetCommands
    imSPersonFirstTime = True   'First time at field-set default if required
    imProdFirstTime = True   'First time at field-set default if required
    imAgyFirstTime = True   'First time at field-set default if required
    imComp0FirstTime = True   'First time at field-set default if required
    imComp1FirstTime = True   'First time at field-set default if required
    imExcl0FirstTime = True   'First time at field-set default if required
    imExcl1FirstTime = True   'First time at field-set default if required
    imBusFirstTime = True
    imGuarFirstTime = True
    imBuyerFirstTime = True   'First time at field-set default if required
    imPhoneFirstTime = True   'First time at field-set default if required
    imDemoFirstTime = True   'First time at field-set default if required
    imInHotSpot = False
    lmLnHotSpot(1, 1) = tmLnCtrls(OV1INDEX).fBoxX '3675  '3555  'Left
    lmLnHotSpot(1, 2) = 15    'Top
    lmLnHotSpot(1, 3) = lmLnHotSpot(1, 1) + 150 '3675 + 150 'Right
    lmLnHotSpot(1, 4) = 15 + 180  'Bottom
    lmLnHotSpot(2, 1) = tmLnCtrls(OV1INDEX).fBoxX + 150 '3825  'Left
    lmLnHotSpot(2, 2) = 15    'Top
    lmLnHotSpot(2, 3) = lmLnHotSpot(2, 1) + 150 '3825 + 150 'Right
    lmLnHotSpot(2, 4) = 15 + 180  'Bottom
    lmLnHotSpot(3, 1) = tmLnCtrls(SPOTSINDEX).fBoxX '7470  'Left
    lmLnHotSpot(3, 2) = 15    'Top
    lmLnHotSpot(3, 3) = lmLnHotSpot(3, 1) + 150 '7470 + 150 'Right
    lmLnHotSpot(3, 4) = 15 + 180  'Bottom
    lmLnHotSpot(4, 1) = tmLnCtrls(SPOTSINDEX).fBoxX + 150 '7620  'Left
    lmLnHotSpot(4, 2) = 15    'Top
    lmLnHotSpot(4, 3) = lmLnHotSpot(4, 1) + 150 '7620 + 150 'Right
    lmLnHotSpot(4, 4) = 15 + 180  'Bottom

    'moved to: mResizeDemoBar
    'lmDBHotSpot(1, 1) = 3675  '3555  'Left
    'lmDBHotSpot(1, 2) = 15    'Top
    'lmDBHotSpot(1, 3) = 3675 + 150 'Right
    'lmDBHotSpot(1, 4) = 15 + 180  'Bottom
    'lmDBHotSpot(2, 1) = 3825  'Left
    'lmDBHotSpot(2, 2) = 15    'Top
    'lmDBHotSpot(2, 3) = 3825 + 150 'Right
    'lmDBHotSpot(2, 4) = 15 + 180  'Bottom
    'lmDBHotSpot(3, 1) = 7470  'Left
    'lmDBHotSpot(3, 2) = 15    'Top
    'lmDBHotSpot(3, 3) = 7470 + 150 'Right
    'lmDBHotSpot(3, 4) = 15 + 180  'Bottom
    'lmDBHotSpot(4, 1) = 7620  'Left
    'lmDBHotSpot(4, 2) = 15    'Top
    'lmDBHotSpot(4, 3) = 7620 + 150 'Right
    'lmDBHotSpot(4, 4) = 15 + 180  'Bottom
    '8032
    igStationXmlChoice = AGENCYXML
    'Moved to Traffic start up
    'mGetHideHiddenLinesFlag
'    '7673
'    mBuildXmlPaths smAgencyXMLPath, smStationXMLWOPath

    Contract.lbcDBBook.Clear
    Contract.edcDBBook.Text = ""
    sgBookNameCodeTag = ""
    '10967
On Error GoTo mInitErr
    tmDigitalVendorsToTest = mGetVendorsAPI()
    imCPMIgnoreScroll = False
    imCPMCtrlVisible = False
    DoEvents
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mBuildWkInfoSub(ilGetData As Integer, ilSetLnDates As Integer, ilGetAvails As Integer)
'
'   Where:
'       ilGetData(I)-True=Get demo data only if Demo Bar visible
'       ilGetAvails(I)- Get avails: Test imToggleAvail instead of this argument
'       imLnRowNo(I)- Line Row No
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim llDate As Long
    Dim llEDate As Long
    Dim ilInfoIndex As Integer
    Dim ilCount As Integer
    Dim llSpots As Long
    Dim ilDay As Integer
    Dim ilCff As Integer
    Dim ilWkIndex As Integer
    Dim ilRet As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfSocEco As Integer
    Dim ilMnfDemo As Integer
    Dim ilRdfCode As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    ReDim ilDays(0 To 6) As Integer
    ReDim ilAvDays(0 To 6) As Integer
    Dim llPop As Long
    Dim llAvgAud As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim llAvgAudAvg As Long
    Dim ilAvgRating As Integer
    Dim llTotalGRP As Long
    Dim llTotalGrImp As Long
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim ilPrevIndex As Integer
    Dim llCPP As Long
    Dim llCPM As Long
    Dim slLineType As String
    Dim llCompStartTime As Long
    Dim llCompEndTime As Long
    Dim llSpotCount As Long
    Dim ilAvail As Integer
    Dim ilInvCount As Integer
    Dim ilCompUnitSold As Integer
    Dim ilNoAvailsForBreak As Integer
    Dim ilNoAvailsForNBB As Integer
    Dim ilAdfCode As Integer
    Dim ilAdfCodeCount As Integer
    Dim llSpanStartDate As Long
    Dim llSpanEndDate As Long
    Dim ilAirWk As Integer
    Dim slDate As String
    Dim ilRowNo As Integer
    Dim ilTest As Integer
    Dim llCntrPop As Long
    Dim ilVefCode As Integer
    Dim ilSvLnRowNo As Integer
    Dim llPopEst As Long
    Dim slStr As String
    Dim llRafCode As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilCompAvailCount As Integer
    Dim blDefaultBookExit As Boolean
    Dim ilVef As Integer

    If bmInCancel Then
        Exit Sub
    End If
    If imInResearch Then
        Exit Sub
    End If
    If imInSummary Then
        Exit Sub
    End If
    If imInBuildWkInfo Then
        Exit Sub
    End If
    If imLnRowNo <= 0 Then
        Exit Sub
    End If
    If (smLnSave(1, imLnRowNo) = "") Then
        Exit Sub
    End If
    imInBuildWkInfo = True
    If lmHSStartDate = -1 Then
        If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
            Contract.mSetHdDate False, False
        Else
            Contract.mSetHdDate True, False
        End If
    End If
    For ilDay = 0 To 6 Step 1
        ilAvDays(ilDay) = True
    Next ilDay

    'Clear any old values
    'mSetDemoBarBooks
    'mInitBookDemoForLn imLnRowNo
    slLineType = mGetLineType(imLnRowNo)
    ilIndex = imLnSave(10, imLnRowNo)
    ilSvLnRowNo = imLnRowNo
    Do While ilIndex > 0
        For ilLoop = LBound(tmWkInfo(ilIndex).lSpotCount) To UBound(tmWkInfo(ilIndex).lSpotCount) Step 1
            tmWkInfo(ilIndex).lSpotCount(ilLoop) = 0
            tmWkInfo(ilIndex).iWkType(ilLoop) = 1   'Assume weekly
            tmWkInfo(ilIndex).lActPrice(ilLoop) = -10
            '3/30/18: changed to use imToggleAvail
            'If (tmWkInfo(ilIndex).lHSStartDate <> lmHSStartDate) Then
                tmWkInfo(ilIndex).iAvail(ilLoop) = -29999   'Not defined
            'End If
            tmWkInfo(ilIndex).lRCPrice(ilLoop) = -1
            If ilGetData Then
                tmWkInfo(ilIndex).lAvgAud(ilLoop) = 0
                tmWkInfo(ilIndex).lGRP(ilLoop) = 0
                tmWkInfo(ilIndex).lGrImp(ilLoop) = 0
                tmWkInfo(ilIndex).lPopEst(ilLoop) = 0
            End If
        Next ilLoop
        tmWkInfo(ilIndex).lHSStartDate = lmHSStartDate
        ilIndex = tmWkInfo(ilIndex).iNextIndex
    Loop
    If ilSetLnDates Then
        smLnSave(14, imLnRowNo) = ""
        smLnSave(15, imLnRowNo) = ""
        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                If smLnSave(14, imLnRowNo) = "" Then
                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), smLnSave(14, imLnRowNo)    'Week Start date
                End If
                gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), smLnSave(15, imLnRowNo)    'Week Start date
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
    Else
        If (smLnSave(14, imLnRowNo) = "") And (Contract.plcDemoBar.Visible = False) Then
            imInBuildWkInfo = False
            Exit Sub
        End If
    End If
    'llStartDate = gDateValue(smLnSave(14, imLnRowNo))
    'llStartMonDate = gDateValue(gObtainPrevMonday(smLnSave(14, imLnRowNo)))
    'llEndDate = gDateValue(smLnSave(15, imLnRowNo))
    llSpanStartDate = lmHSStartDate + 7 * (imHSCurWkNo - 1)
    llSpanEndDate = llSpanStartDate + 7 * 13 - 1
    'If llEndDate < llStartDate Then
    '    Exit Sub
    'End If
    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
        mCVehOptIndex imLnRowNo, imCntrVefCode, imVpfIndex, imTerminate
        blDefaultBookExit = False
        ilVef = gBinarySearchVef(imCntrVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).iDnfCode > 0 Then
                blDefaultBookExit = True
            End If
        End If
        If (Contract.plcDemoBar.Visible = False) Or blDefaultBookExit Then
            mSetDemoBarBooks imLnRowNo, ilDnfCode, imCppCpm, 2, imCntrVefCode, imVpfIndex, imTerminate, imChgMode   ' (mBuildWkInfo sets tgClf to DnfCode)
        Else
            mSetDemoBarBooks imLnRowNo, ilDnfCode, imCppCpm, 0, imCntrVefCode, imVpfIndex, imTerminate, imChgMode   ' (mBuildWkInfo sets tgClf to DnfCode)
        End If
        'If Contract.lbcDBBook.ListIndex >= 0 Then
        '    slNameCode = tgBookNameCode(Contract.lbcDBBook.ListIndex).sKey  'Traffic!lbcBookNameCode.List(Contract.lbcDBBook.ListIndex)
        '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
        '    ilDnfCode = Val(slCode)
        'Else
        '    ilDnfCode = -1
        'End If
        If tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode <> ilDnfCode Then
            tgClfCntr(imLnRowNo - 1).iResearch = False
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Then
                Else
                    If (Contract.lbcDBBook.ListIndex >= 0) And (Contract.lbcDBDemo.ListIndex >= 0) Then
                        imLnChg = True
                    End If
                End If
            End If
        End If
        If ilDnfCode > 0 Then
            tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode = ilDnfCode
            If Contract.lbcDBDemo.ListIndex >= 0 Then
                'Find matching name
                If (Contract.lbcDemo(0).ListIndex >= 1) Or (Contract.lbcDemo(1).ListIndex >= 1) Or (Contract.lbcDemo(2).ListIndex >= 1) Or (Contract.lbcDemo(3).ListIndex >= 1) Then
                    ilIndex = Contract.lbcDemo(imDemoMap(Contract.lbcDBDemo.ListIndex)).ListIndex
                    slNameCode = tgDemoCode(ilIndex - 1).sKey  'Traffic!Contract.lbcDemoCode.List(ilIndex - 1)
                Else
                    ilIndex = Contract.lbcDBDemo.ListIndex
                    slNameCode = tgDemoCode(ilIndex).sKey  'Traffic!Contract.lbcDemoCode.List(ilIndex - 1)
                End If
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfDemo = Val(slCode)
            Else
                ilMnfDemo = 0
            End If
        Else
            tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode = 0
            ilMnfDemo = 0
        End If
        If tgClfCntr(imLnRowNo - 1).ClfRec.iMnfDemo <> ilMnfDemo Then
            tgClfCntr(imLnRowNo - 1).iResearch = False
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Then
                Else
                    If (Contract.lbcDBBook.ListIndex >= 0) And (Contract.lbcDBDemo.ListIndex >= 0) Then
                        imLnChg = True
                    End If
                End If
            End If
        End If
        tgClfCntr(imLnRowNo - 1).ClfRec.iMnfDemo = ilMnfDemo
        If Contract.lbcDBSocEco.ListIndex >= 1 Then
            slNameCode = tgSocEcoCode(Contract.lbcDBSocEco.ListIndex - 1).sKey  'Traffic!lbcSocEcoCode.List(Contract.lbcDBSocEco.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMnfSocEco = Val(slCode)
        Else
            ilMnfSocEco = 0
        End If
        If imLnSave(4, imLnRowNo) = 1 Then
            llOvStartTime = CLng(gTimeToCurrency(smLnSave(2, imLnRowNo), False))
            llOvEndTime = CLng(gTimeToCurrency(smLnSave(3, imLnRowNo), True))
        Else
            llOvStartTime = 0
            llOvEndTime = 0
        End If
        If ilGetData And (Contract.plcDemoBar.Visible = True) Then
            If (ilDnfCode > 0) And (ilMnfDemo > 0) Then
                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
            End If
        End If
        ilRdfCode = imLnSave(1, imLnRowNo)
    Else
        llCntrPop = -1
        ReDim tlPkWkInfo(0 To 0) As PkWkInfo
        ReDim igLnMakingPack(0 To 0) As Integer
        '10389
        'If ilGetData And (Contract.plcDemoBar.Visible = True) Then
        If ilGetData And gVisibleTabNeedsResearch() Then
            'Build an array of line number that make up the package
            For ilTest = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                If (tgPackLn(ilTest).sType = "A") Or (tgPackLn(ilTest).sType = "O") Or (tgPackLn(ilTest).sType = "E") Then
                    '3/28/16
                    If tgSaf(0).sAudByPackage = "Y" Then
                        If tgPackLn(ilTest).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                            igLnMakingPack(UBound(igLnMakingPack)) = tgPackLn(ilTest).iLineNo
                            ReDim Preserve igLnMakingPack(0 To UBound(igLnMakingPack) + 1) As Integer
                            Exit For
                        End If
                    Else
                        If tgPackLn(ilTest).iLineNo = tgClfCntr(imLnRowNo - 1).ClfRec.iLine Then
                            ilIndex = tgPackLn(ilTest).iIndex
                            Do While ilIndex >= 0
                                igLnMakingPack(UBound(igLnMakingPack)) = tgPackLn(ilIndex).iLineNo
                                ReDim Preserve igLnMakingPack(0 To UBound(igLnMakingPack) + 1) As Integer
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                            Exit For
                        Else
                            ilIndex = tgPackLn(ilTest).iIndex
                            Do While ilIndex >= 0
                                If tgClfCntr(imLnRowNo).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                                    ilIndex = tgPackLn(ilTest).iIndex
                                    Do While ilIndex >= 0
                                        igLnMakingPack(UBound(igLnMakingPack)) = tgPackLn(ilIndex).iLineNo
                                        ReDim Preserve igLnMakingPack(0 To UBound(igLnMakingPack) + 1) As Integer
                                        ilIndex = tgPackLn(ilIndex).iIndex
                                    Loop
                                    Exit For
                                End If
                                ilIndex = tgPackLn(ilIndex).iIndex
                            Loop
                        End If
                    End If
                End If
            Next ilTest
            ReDim tmPkWkInfo(0 To UBound(igLnMakingPack)) As WKINFO
            'For ilRowNo = LBound(smLnSave, 2) To UBound(smLnSave, 2) - 1 Step 1
            For ilRowNo = imLB1Or2 To UBound(smLnSave, LINEBOUNDINDEX) - 1 Step 1
                For ilLoop = LBound(igLnMakingPack) To UBound(igLnMakingPack) - 1 Step 1
                    If igLnMakingPack(ilLoop) = tgClfCntr(ilRowNo - 1).ClfRec.iLine Then
                        If imLnSave(4, ilRowNo) = 1 Then
                            llOvStartTime = CLng(gTimeToCurrency(smLnSave(2, ilRowNo), False))
                            llOvEndTime = CLng(gTimeToCurrency(smLnSave(3, ilRowNo), True))
                        Else
                            llOvStartTime = 0
                            llOvEndTime = 0
                        End If
                        imLnRowNo = ilRowNo
                        mInitBookDemoForLn ilRowNo, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode
                        ilDnfCode = tgClfCntr(ilRowNo - 1).ClfRec.iDnfCode
                        ilMnfDemo = tgClfCntr(ilRowNo - 1).ClfRec.iMnfDemo
                        ilRdfCode = imLnSave(1, ilRowNo)
                        gFindMatch smLnSave(1, ilRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
                        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) < 0 Then
                            ilVefCode = 0
                        Else
                            slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(Contract.lbcLnVehicle(igTabMapIndex)))
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            ilVefCode = CInt(slCode)
                        End If
                        If (ilDnfCode > 0) And (ilMnfDemo > 0) And (ilVefCode > 0) Then
                            ilMnfSocEco = 0
                            ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
                            If tgSpf.sDemoEstAllowed <> "Y" Then
                                If llPop > 0 Then
                                    If llCntrPop = -1 Then
                                        llCntrPop = llPop
                                    Else
                                        If llCntrPop <> llPop Then
                                            llCntrPop = 0
                                        End If
                                    End If
                                End If
                            End If
                            ReDim llPkSpotCount(0 To 0) As Long
                            ReDim llPkActPrice(0 To 0) As Long
                            ReDim llPkAvgAud(0 To 0) As Long
                            ReDim llPkDate(0 To 0) As Long
                            ReDim llPkPopEst(0 To 0) As Long
                            For llDate = llSpanStartDate To llSpanEndDate Step 7
                                ilAirWk = False
                                For ilDay = 0 To 6 Step 1
                                    ilDays(ilDay) = False
                                Next ilDay
                                ilCff = tgClfCntr(ilRowNo - 1).iFirstCff
                                Do While ilCff <> -1
                                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                        'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
                                        'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
                                        llFlStartDate = tgCffCntr(ilCff).lStartDate
                                        llFlEndDate = tgCffCntr(ilCff).lEndDate
                                        If llDate < llFlStartDate Then
                                            Exit Do
                                        End If
                                        If (llDate + 6 >= llFlStartDate) And (llDate <= llFlEndDate) Then
                                            ilAirWk = True
                                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                                llSpots = 0
                                                For ilDay = 0 To 6 Step 1
                                                    If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                                                        llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                                        If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                                                            ilDays(ilDay) = True
                                                        End If
                                                    End If
                                                Next ilDay
                                            Else
                                                llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                                For ilDay = 0 To 6 Step 1
                                                    If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                                                        If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                                            ilDays(ilDay) = True
                                                        End If
                                                    End If
                                                Next ilDay
                                            End If
                                            'ilWkIndex = (llDate - lmHSStartDate) \ 7 + 1
                                            'tmPkWkInfo(ilLoop).iSpotCount(ilWkIndex) = llSpots
                                            'tmPkWkInfo(ilLoop).lActPrice(ilWkIndex) = 0
                                            llPkSpotCount(UBound(llPkSpotCount)) = llSpots
                                            llPkActPrice(UBound(llPkActPrice)) = 0
                                            llPkDate(UBound(llPkDate)) = llDate
                                            slStr = mGetLineType(ilRowNo, True)
                                            llRafCode = mGetRafCode(ilRowNo, tmRegionCode())
                                            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, ilRdfCode, llOvStartTime, llOvEndTime, ilDays(), slStr, llRafCode, llPkAvgAud(UBound(llPkAvgAud)), llPkPopEst(UBound(llPkAvgAud)), ilAudFromSource, llAudFromCode) 'tmPkWkInfo(ilLoop).lAvgAud(ilWkIndex))
                                            ReDim Preserve llPkSpotCount(0 To UBound(llPkSpotCount) + 1) As Long
                                            ReDim Preserve llPkActPrice(0 To UBound(llPkActPrice) + 1) As Long
                                            ReDim Preserve llPkAvgAud(0 To UBound(llPkAvgAud) + 1) As Long
                                            ReDim Preserve llPkDate(0 To UBound(llPkDate) + 1) As Long
                                            ReDim Preserve llPkPopEst(0 To UBound(llPkPopEst) + 1) As Long
                                            '10256
                                            mActOneLineInfoSet ilRowNo
                                            Exit Do
                                        End If
                                    End If
                                    ilCff = tgCffCntr(ilCff).iNextCff
                                Loop
                            Next llDate
                             If UBound(llPkSpotCount) > 0 Then
                                ReDim Preserve llPkSpotCount(0 To UBound(llPkSpotCount) - 1) As Long
                                ReDim Preserve llPkActPrice(0 To UBound(llPkActPrice) - 1) As Long
                                ReDim Preserve llPkAvgAud(0 To UBound(llPkAvgAud) - 1) As Long
                                ReDim Preserve llPkDate(0 To UBound(llPkDate) - 1) As Long
                                ReDim Preserve llPkPopEst(0 To UBound(llPkPopEst) - 1) As Long
                                ReDim ilRating(LBound(llPkSpotCount) To UBound(llPkSpotCount)) As Integer
                                ReDim llPkGRP(LBound(llPkSpotCount) To UBound(llPkSpotCount)) As Long
                                ReDim llPkGrImp(LBound(llPkSpotCount) To UBound(llPkSpotCount)) As Long
                                'gAvgAudToLnResearch smCntr1or2PlaceRating, False, llPop, llPkPopEst(), llPkSpotCount(), llPkActPrice(), llPkAvgAud(), llTotalCost, llAvgAudAvg, ilRating(), ilAvgRating, llPkGrImp(), llTotalGrImp, llPkGRP(), llTotalGRP, llCPP, llCPM, llPopEst
                                gAvgAudToLnResearch smCntr1or2PlaceRating, False, llPop, llPkPopEst(), llPkSpotCount(), llPkActPrice(), llPkAvgAud(), dlTotalCost, llAvgAudAvg, ilRating(), ilAvgRating, llPkGrImp(), llTotalGrImp, llPkGRP(), llTotalGRP, llCPP, llCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                                If tgSpf.sDemoEstAllowed = "Y" Then
                                    If llPopEst > 0 Then
                                        If llCntrPop = -1 Then
                                            llCntrPop = llPopEst
                                        Else
                                            If llCntrPop <> llPopEst Then
                                                llCntrPop = 0
                                            End If
                                        End If
                                    End If
                                End If
                                For ilDay = LBound(llPkDate) To UBound(llPkDate) Step 1
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lDate = llPkDate(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lSpotCount = llPkSpotCount(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lActPrice = llPkActPrice(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lAvgAud = llPkAvgAud(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lGRP = llPkGRP(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lGrImp = llPkGrImp(ilDay)
                                    tlPkWkInfo(UBound(tlPkWkInfo)).lPopEst = llPkPopEst(ilDay)
                                     ReDim Preserve tlPkWkInfo(0 To UBound(tlPkWkInfo) + 1) As PkWkInfo
                                Next ilDay
                            End If
                        End If
                        Exit For
                    End If
                Next ilLoop
            Next ilRowNo
            imLnRowNo = ilSvLnRowNo
        End If
        If imLnSave(4, imLnRowNo) = 1 Then
            llOvStartTime = CLng(gTimeToCurrency(smLnSave(2, imLnRowNo), False))
            llOvEndTime = CLng(gTimeToCurrency(smLnSave(3, imLnRowNo), True))
        Else
            llOvStartTime = 0
            llOvEndTime = 0
        End If
        For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
            If StrComp(smLnSave(20, imLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                ilRdfCode = tgMRdf(ilLoop).iCode
                Exit For
            End If
        Next ilLoop
    End If
    For llDate = llSpanStartDate To llSpanEndDate Step 7
        ilAirWk = False
        ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
                'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
                llFlStartDate = tgCffCntr(ilCff).lStartDate
                llFlEndDate = tgCffCntr(ilCff).lEndDate
                If llDate < llFlStartDate Then
                    Exit Do
                End If
                If (llDate + 6 >= llFlStartDate) And (llDate <= llFlEndDate) Then
                    ilAirWk = True
                    Exit Do
                End If
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
        ilWkIndex = (llDate - lmHSStartDate) \ 7 + 1
        ilInfoIndex = (ilWkIndex - 1) \ (UBound(tmWkInfo(1).lSpotCount) + 1) + 1
        ilIndex = imLnSave(10, imLnRowNo)
        ilCount = 0
        Do
            If ilIndex <= 0 Then
                ReDim Preserve tmWkInfo(0 To UBound(tmWkInfo) + 1) As WKINFO
                If ilCount = 0 Then
                    imLnSave(10, imLnRowNo) = UBound(tmWkInfo) - 1
                Else
                    tmWkInfo(ilPrevIndex).iNextIndex = UBound(tmWkInfo) - 1
                End If
                ilIndex = UBound(tmWkInfo) - 1
                tmWkInfo(ilIndex).iNextIndex = -1
                For ilLoop = LBound(tmWkInfo(ilIndex).lSpotCount) To UBound(tmWkInfo(ilIndex).lSpotCount) Step 1
                    tmWkInfo(ilIndex).lSpotCount(ilLoop) = 0
                    tmWkInfo(ilIndex).iWkType(ilLoop) = 1   'Assume weekly
                    tmWkInfo(ilIndex).iAvail(ilLoop) = -29999   'Not defined
                    tmWkInfo(ilIndex).lActPrice(ilLoop) = -10
                    tmWkInfo(ilIndex).lRCPrice(ilLoop) = -1
                    tmWkInfo(ilIndex).lAvgAud(ilLoop) = 0
                    tmWkInfo(ilIndex).lGRP(ilLoop) = 0
                    tmWkInfo(ilIndex).lGrImp(ilLoop) = 0
                    tmWkInfo(ilIndex).lPopEst(ilLoop) = 0
                    tmWkInfo(ilIndex).lHSStartDate = lmHSStartDate
                Next ilLoop
            Else
                If ilCount + 1 = ilInfoIndex Then
                    ilCount = ilCount + 1
                    Exit Do
                End If
                ilPrevIndex = ilIndex
                ilIndex = tmWkInfo(ilIndex).iNextIndex
            End If
            ilCount = ilCount + 1
        Loop While ilCount <> ilInfoIndex
        ilWkIndex = ilWkIndex - (ilCount - 1) * (UBound(tmWkInfo(ilIndex).lSpotCount) + 1)
        For ilDay = 0 To 6 Step 1
            ilDays(ilDay) = False
        Next ilDay
        If ilAirWk Then
            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                tmWkInfo(ilIndex).iWkType(ilWkIndex - 1) = 0
                llSpots = 0
                For ilDay = 0 To 6 Step 1
                    If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                        llSpots = llSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                        If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                            ilDays(ilDay) = True
                        End If
                    End If
                Next ilDay
            Else
                tmWkInfo(ilIndex).iWkType(ilWkIndex - 1) = 1
                llSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                For ilDay = 0 To 6 Step 1
                    If (llDate + ilDay >= llFlStartDate) And (llDate + ilDay <= llFlEndDate) Then
                        If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                            ilDays(ilDay) = True
                        End If
                    End If
                Next ilDay
            End If
            tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1) = tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1) + llSpots
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                Select Case tgCffCntr(ilCff).CffRec.sPriceType
                    Case "T"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = tgCffCntr(ilCff).CffRec.lActPrice
                    Case "N"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -1
                    Case "M"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -2
                    Case "B"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -3
                    Case "S"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -4
                    Case "P"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -5
                    Case "R"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -6
                    Case "A"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = -7
                End Select
                If tgCffCntr(ilCff).CffRec.sPriceType = "*" Then
                    slDate = Format$(llDate, "m/d/yy")
                    tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slDate, imCntrVefCode, imLnRowNo)
                    If tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) > 0 Then
                        slCode = mRCPriceAdjust(tmRcf, tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1), False, imLnRowNo, imDGSave(), 0, 0, Val(smLnSave(16, imLnRowNo)), llSpots, ilDays(), imLnSave(1, imLnRowNo))
                        slCode = gRoundStr(slCode, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                        tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = gStrDecToLong(slCode, 0)
                    End If
                Else
                    tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = tgCffCntr(ilCff).CffRec.lPropPrice
                End If
                'If (ilGetData Or ilGetAvails) And (Contract.plcDemoBar.Visible = True) Then
                If (ilGetData Or (imToggleAvail >= 0)) And (Contract.plcDemoBar.Visible = True) Then
                    If ilGetData Then
                        If (ilDnfCode > 0) And (imCntrVefCode > 0) And (ilMnfDemo > 0) Then
                            slStr = mGetLineType(imLnRowNo, True)
                            llRafCode = mGetRafCode(imLnRowNo, tmRegionCode())
                            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, imCntrVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, ilRdfCode, llOvStartTime, llOvEndTime, ilDays(), slStr, llRafCode, tmWkInfo(ilIndex).lAvgAud(ilWkIndex - 1), tmWkInfo(ilIndex).lPopEst(ilWkIndex - 1), ilAudFromSource, llAudFromCode)
                            '10256
                             mActOneLineInfoSet imLnRowNo
                        Else
                            tmWkInfo(ilIndex).lAvgAud(ilWkIndex - 1) = 0
                            tmWkInfo(ilIndex).lPopEst(ilWkIndex - 1) = 0
                        End If
                    End If
                    llEDate = llDate + 6
                    If llEDate > llFlEndDate Then
                        llEDate = llFlEndDate
                    End If
                    igComp1Code = 0
                    igComp2Code = 0
                    For ilDay = 0 To 6 Step 1
                        ilAvDays(ilDay) = ilDays(ilDay)
                    Next ilDay
                    ilAdfCode = 0
                    ilAdfCodeCount = 0
                    'If (tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = -29999) And (ilGetAvails) Then
                    If (tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = -29999) And (imToggleAvail >= 0) Then
                        If tgSpf.sCIncludeMissDB = "N" Then
                            mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(imLnRowNo - 1).ClfRec.iLine, imCntrVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), ilAvDays(), llDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoAvailsForBreak, ilNoAvailsForNBB, llCompStartTime, llCompEndTime, False, ilAdfCode, ilAdfCodeCount, ilCompAvailCount
                        Else
                            mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(imLnRowNo - 1).ClfRec.iLine, imCntrVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), ilAvDays(), llDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoAvailsForBreak, ilNoAvailsForNBB, llCompStartTime, llCompEndTime, True, ilAdfCode, ilAdfCodeCount, ilCompAvailCount
                        End If
                        tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = ilAvail
                    End If
                End If
            Else
                Select Case tgCffCntr(ilCff).CffRec.sPriceType
                    Case "T"
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = tgCffCntr(ilCff).CffRec.lActPrice
                    Case Else
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = 0
                End Select
                tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = 0

                If ilGetData And (Contract.plcDemoBar.Visible = True) And (UBound(igLnMakingPack) > 0) Then
                    ReDim lgLnCost(0 To 0) As Long
                    ReDim igLnAvgRating(0 To 0) As Integer
                    ReDim lgLnGrImp(0 To 0) As Long
                    ReDim lgLnGRP(0 To 0) As Long
                    If tgSpf.sDemoEstAllowed = "Y" Then
                        llCntrPop = -1
                    End If
                    For ilLoop = LBound(tlPkWkInfo) To UBound(tlPkWkInfo) - 1 Step 1
                        If llDate = tlPkWkInfo(ilLoop).lDate Then
                            lgLnCost(UBound(lgLnCost)) = 0
                            igLnAvgRating(UBound(igLnAvgRating)) = 0
                            lgLnGrImp(UBound(lgLnGrImp)) = tlPkWkInfo(ilLoop).lGrImp
                            lgLnGRP(UBound(lgLnGRP)) = tlPkWkInfo(ilLoop).lGRP
                            ReDim Preserve lgLnCost(0 To UBound(lgLnCost) + 1) As Long
                            ReDim Preserve igLnAvgRating(0 To UBound(igLnAvgRating) + 1) As Integer
                            ReDim Preserve lgLnGrImp(0 To UBound(lgLnGrImp) + 1) As Long
                            ReDim Preserve lgLnGRP(0 To UBound(lgLnGRP) + 1) As Long
                            If tgSpf.sDemoEstAllowed = "Y" Then
                                If llCntrPop = -1 Then
                                    llCntrPop = tlPkWkInfo(ilLoop).lPopEst
                                Else
                                    If llCntrPop <> tlPkWkInfo(ilLoop).lPopEst Then
                                        llCntrPop = 0
                                    End If
                                End If
                            End If
                        End If
                    Next ilLoop
                    'If UBound(lgLnCost) > 1 Then
                    If UBound(lgLnCost) > 0 Then
                        ReDim Preserve lgLnCost(0 To UBound(lgLnCost) - 1) As Long
                        ReDim Preserve igLnAvgRating(0 To UBound(igLnAvgRating) - 1) As Integer
                        ReDim Preserve lgLnGrImp(0 To UBound(lgLnGrImp) - 1) As Long
                        ReDim Preserve lgLnGRP(0 To UBound(lgLnGRP) - 1) As Long
                        'gResearchTotals smCntr1or2PlaceRating, False, llCntrPop, lgLnCost(), lgLnGrImp(), lgLnGRP(), CLng(tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1)), llTotalCost, ilAvgRating, tmWkInfo(ilIndex).lGrImp(ilWkIndex - 1), tmWkInfo(ilIndex).lGRP(ilWkIndex - 1), llCPP, llCPM, llAvgAud
                        gResearchTotals smCntr1or2PlaceRating, False, llCntrPop, lgLnCost(), lgLnGrImp(), lgLnGRP(), CLng(tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1)), dlTotalCost, ilAvgRating, tmWkInfo(ilIndex).lGrImp(ilWkIndex - 1), tmWkInfo(ilIndex).lGRP(ilWkIndex - 1), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
                    End If
                End If
            End If
        Else
            igComp1Code = 0
            igComp2Code = 0
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                mDPAllowedDays imLnSave(1, imLnRowNo), ilDays()
            Else
                For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If StrComp(smLnSave(20, imLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                        mDPAllowedDays tgMRdf(ilLoop).iCode, ilDays()
                        Exit For
                    End If
                Next ilLoop
            End If
            llEDate = llDate + 6
            For ilDay = 0 To 6 Step 1
                ilAvDays(ilDay) = ilDays(ilDay)
            Next ilDay
            ilAdfCode = 0
            ilAdfCodeCount = 0
            'If (tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = -29999) And (ilGetAvails) Then
            If (tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = -29999) And (imToggleAvail >= 0) Then
                If tgSpf.sCIncludeMissDB = "N" Then
                    mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(imLnRowNo - 1).ClfRec.iLine, imCntrVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), ilAvDays(), llDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoAvailsForBreak, ilNoAvailsForNBB, llCompStartTime, llCompEndTime, False, ilAdfCode, ilAdfCodeCount, ilCompAvailCount
                Else
                    mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(imLnRowNo - 1).ClfRec.iLine, imCntrVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, imLnRowNo)), ilAvDays(), llDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoAvailsForBreak, ilNoAvailsForNBB, llCompStartTime, llCompEndTime, True, ilAdfCode, ilAdfCodeCount, ilCompAvailCount
                End If
                tmWkInfo(ilIndex).iAvail(ilWkIndex - 1) = ilAvail
            End If
            slDate = Format$(llDate, "m/d/yy")
            tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = mComputeRCWkPrice(tmRcf.iCode, imLnSave(1, imLnRowNo), slDate, imCntrVefCode, imLnRowNo)
            If tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) > 0 Then
                slCode = mRCPriceAdjust(tmRcf, tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1), False, imLnRowNo, imDGSave(), 0, 0, Val(smLnSave(16, imLnRowNo)), 0, ilDays(), imLnSave(1, imLnRowNo))
                slCode = gRoundStr(slCode, "1.00", 0)    'gLongToStrDec(llWkPrice, 0) '& ".00"
                tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1) = gStrDecToLong(slCode, 0)
            End If
        End If
    Next llDate
    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
        '10389
        'If ilGetData And (Contract.plcDemoBar.Visible = True) Then
        If ilGetData And gVisibleTabNeedsResearch() Then
            ReDim ilRating(LBound(tmWkInfo(ilIndex).lSpotCount) To UBound(tmWkInfo(ilIndex).lSpotCount)) As Integer
            ilIndex = imLnSave(10, imLnRowNo)
            Do While ilIndex > 0
                'gAvgAudToLnResearch smCntr1or2PlaceRating, False, llPop, tmWkInfo(ilIndex).lPopEst(), tmWkInfo(ilIndex).lSpotCount(), tmWkInfo(ilIndex).lActPrice(), tmWkInfo(ilIndex).lAvgAud(), llTotalCost, llAvgAudAvg, ilRating(), ilAvgRating, tmWkInfo(ilIndex).lGrImp(), llTotalGrImp, tmWkInfo(ilIndex).lGRP(), llTotalGRP, llCPP, llCPM, llPopEst
                gAvgAudToLnResearch smCntr1or2PlaceRating, False, llPop, tmWkInfo(ilIndex).lPopEst(), tmWkInfo(ilIndex).lSpotCount(), tmWkInfo(ilIndex).lActPrice(), tmWkInfo(ilIndex).lAvgAud(), dlTotalCost, llAvgAudAvg, ilRating(), ilAvgRating, tmWkInfo(ilIndex).lGrImp(), llTotalGrImp, tmWkInfo(ilIndex).lGRP(), llTotalGRP, llCPP, llCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                ilIndex = tmWkInfo(ilIndex).iNextIndex
            Loop
        End If
        Erase ilRating
    Else
        Erase tlPkWkInfo
        Erase llPkSpotCount
        Erase llPkActPrice
        Erase llPkDate
        Erase llPkAvgAud
        Erase ilRating
        Erase llPkGRP
        Erase llPkGrImp
        Erase llPkPopEst
    End If
    If imLnRowNo > 0 Then
        mInitBookDemoForLn imLnRowNo, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode
    End If
    imInBuildWkInfo = False
End Sub

Public Sub mSelectChangeSub(CNTRNOINDEX As Integer, MERCHINDEX As Integer, PROMOINDEX As Integer, INTCOMMINDEX As Integer, CANCLAUSEINDEX As Integer, CHGREASONINDEX As Integer, OTHERCOMMINDEX As Integer, STATUSINDEX As Integer, REVNOINDEX As Integer, GUARINDEX As Integer, GROSSINDEX As Integer, PRODTAGINDEX As Integer, RATECARDINDEX As Integer, DEMOINDEX As Integer, DATEINDEX As Integer)
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilPos As Integer
    Dim ilSbfRet As Integer
    Dim ilLoopCount As Integer
    Dim ilCntRevNo As Integer
    Dim ilPropVer As Integer
    Dim slCntRevNo As String
    Dim slPropVer As String
    Dim slCntrNo As String
    Dim llCntrNo As Long
    Dim ilFound As Integer
    Dim slMsg As String
    Dim ilUpdateAllowed As Integer
    Dim ilStep3Index As Integer
    Dim ilStep2Index As Integer
    Dim ilStep1Index As Integer
    Dim ilStepChgd As Integer
    Dim ilSvFreezeState As Integer

    Contract.tmcClick.Enabled = False
    imSelectDelay = False
    imSelectUpdate = False
    Contract.mPaintStep
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        Contract.tmcClick.Enabled = False    'For safety
        Contract.pbcEatTab(1).Enabled = False
        Contract.pbcEatTab(0).Enabled = False
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        imInSelect = True
        'ilRet = btrUnlock(hmChf, BTRV_UNLOCK_MULTIPLE)
        ilRet = gDeleteLockRec_ByRlfCode(hmRlf, lmLockRecCode)
        If sgCntrScreen = "Orders" Then
            If (imStep1Index >= 0) And (imStep1Index <= 3) Then
                If (igWinStatus(CONTRACTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    imUpdateAllowed = False
                Else
                    imUpdateAllowed = True
                End If
            Else
                imUpdateAllowed = False
            End If
        Else
            If (imStep1Index >= 0) And (imStep1Index <= 5) Then
                If (igWinStatus(PROPOSALSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    imUpdateAllowed = False
                Else
                    imUpdateAllowed = True
                End If
            Else
                imUpdateAllowed = False
            End If
        End If
        '8/16/12: Handle case where contract was set as Calendar but is not allowed at this time
        Contract.mPopBillCycle False, False
        ilSvFreezeState = igFreezeState
        igFreezeState = 0
        lgSchChfCode = 0
        Do
            ilStep1Index = imStep1Index
            ilStep2Index = imStep2Index
            ilStep3Index = imStep3Index
            If (imStep1Index = 0) Then
                imStep3Index = -1
                ilStep3Index = -1
                ilRet = 2
                slStr = ""
            Else
                If (imStep2Index = 0) Or (imStep2Index = -1) Then    'By number
                    imStep2Index = 0
                    ilStep2Index = 0
                    If Contract.lbcStep(1).ListCount <= 0 Then
                        Contract.lbcStep(1).AddItem "[by Number]", 0
                    End If
                    Contract.lbcStep(1).ListIndex = 0
                    slStr = Contract.edcStepDropDown.Text    'Contract.lbcStep(2).Text
                    ilRet = gParseItem(slStr, 1, " ", slCntrNo)
                    If gCompNumberStr(slCntrNo, "2000000000") >= 0 Then
                        Beep
                        MsgBox "Contract " & slCntrNo & " too Large", vbOKOnly + vbExclamation, "Warning"
                        Contract.pbcHeader.Cls
                        Contract.pbcLine.Cls
                        Contract.pbcLnTotal.Cls
                        Contract.pbcLnDates.Cls
                        Contract.mPaintHdTitle
                        Contract.mPaintLnTitle
                        Contract.mPaintLnDates
                        Contract.mPaintLnTotal
                        imStep3Index = -2
                        Contract.mSetCommands
                        Screen.MousePointer = vbDefault
                        imChgMode = False
                        imInSelect = False
                        igFreezeState = ilSvFreezeState
                        Exit Sub
                    End If
                    llCntrNo = Val(slCntrNo)
                    tmChfSrchKey1.lCntrNo = llCntrNo
                    ilCntRevNo = 32000
                    ilPropVer = 32000
                    ilRet = gParseItem(slStr, 2, " ", slCntRevNo)
                    If ilRet = CP_MSG_NONE Then
                        If UCase(Left$(slCntRevNo, 1)) = "R" Then
                            slCntRevNo = Mid$(slCntRevNo, 2)
                            ilPos = InStr(slCntRevNo, "-")
                            If ilPos > 0 Then
                                ilCntRevNo = Val(Left$(slCntRevNo, ilPos - 1))
                            Else
                                ilCntRevNo = Val(slCntRevNo)
                            End If
                            ilRet = gParseItem(slStr, 3, " ", slPropVer)
                            If ilRet = CP_MSG_NONE Then
                                If UCase(Left$(slPropVer, 1)) = "V" Then
                                    slPropVer = Mid$(slPropVer, 2)
                                End If
                                ilPropVer = Val(slPropVer)
                            End If
                        ElseIf UCase(Left$(slCntRevNo, 1)) = "V" Then
                            slPropVer = Mid$(slCntRevNo, 2)
                            ilPropVer = Val(slPropVer)
                        Else
                            'ilCntRevNo = Val(slCntRevNo)
                            ilPos = InStr(slCntRevNo, "-")
                            If ilPos > 0 Then
                                ilCntRevNo = Val(Left$(slCntRevNo, ilPos - 1))
                            Else
                                ilCntRevNo = Val(slCntRevNo)
                            End If
                            ilRet = gParseItem(slStr, 3, " ", slPropVer)
                            If ilRet = CP_MSG_NONE Then
                                If UCase(Left$(slPropVer, 1)) = "V" Then
                                    slPropVer = Mid$(slPropVer, 2)
                                End If
                                ilPropVer = Val(slPropVer)
                            End If
                        End If
                    End If
                    tmChfSrchKey1.iCntRevNo = ilCntRevNo
                    tmChfSrchKey1.iPropVer = ilPropVer
                    ilRet = gCHF1GetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo)
                        If (ilCntRevNo = 32000) And (ilPropVer = 32000) Then
                            Exit Do
                        End If
                        If (ilCntRevNo = 32000) And (tmChf.iPropVer = ilPropVer) Then
                            Exit Do
                        End If
                        If (ilPropVer = 32000) And (tmChf.iCntRevNo = ilCntRevNo) Then
                            Exit Do
                        End If
                        If (tmChf.iPropVer = ilPropVer) And (tmChf.iCntRevNo = ilCntRevNo) Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet <> BTRV_ERR_NONE) Or (tmChf.lCntrNo <> llCntrNo) Or ((tmChf.iCntRevNo <> ilCntRevNo) And (ilCntRevNo <> 32000)) Or ((tmChf.iPropVer <> ilPropVer) And (ilPropVer <> 32000)) Then
                        MsgBox "Contract " & slCntrNo & " Not Found", vbOKOnly + vbExclamation, "Warning"
                        slStr = Contract.edcStepDropDown.Text    'Contract.lbcStep(2).Text
                        If slStr <> "" Then
                            Beep
                        End If
                        Contract.pbcHeader.Cls
                        Contract.pbcLine.Cls
                        Contract.pbcLnTotal.Cls
                        Contract.pbcLnDates.Cls
                        Contract.mPaintHdTitle
                        Contract.mPaintLnTitle
                        Contract.mPaintLnDates
                        Contract.mPaintLnTotal
                        imStep3Index = -2
                        Contract.mSetCommands
                        Screen.MousePointer = vbDefault
                        imChgMode = False
                        imInSelect = False
                        igFreezeState = ilSvFreezeState
                        Exit Sub
                    End If
                    ilFound = True
                    'Proposal "Revise Proposal"- disallow change to Order
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 2) And (tmChf.iCntRevNo > 0) Then
                        slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                        ilFound = False
                    End If
                    'Proposal "Revise Proposal"- disallow change to Order
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 2) And (tmChf.iCntRevNo = 0) And ((tmChf.sStatus = "G") Or (tmChf.sStatus = "H") Or (tmChf.sStatus = "N") Or (tmChf.sStatus = "O")) Then
                        slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                        ilFound = False
                    End If
                    'Proposal "New Version Proposal"- disallow change to Order
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 3) And (tmChf.iCntRevNo > 0) Then
                        slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                        ilFound = False
                    End If
                    'Proposal "New Version Proposal"- disallow change to Order
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 3) And (tmChf.iCntRevNo = 0) And ((tmChf.sStatus = "G") Or (tmChf.sStatus = "H") Or (tmChf.sStatus = "N") Or (tmChf.sStatus = "O")) Then
                        slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                        ilFound = False
                    End If
                    'Proposal "Revise Order"- disallow change to proposal
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 4) And (tmChf.iCntRevNo <= 0) And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Or (tmChf.sStatus = "D")) Then
                        slMsg = "Contract " & slCntrNo & " is a Proposal, Use 'Change an Existing Proposal' "
                        ilFound = False
                    End If
                    'Proposal "Change Hold to Order"- allow change to hold
                    If (sgCntrScreen <> "Orders") And (imStep1Index = 5) And (tmChf.sStatus <> "H") Then
                        slMsg = "Contract " & slCntrNo & " is Not currently set As Hold"
                        ilFound = False
                    End If
                    'Order "Revise Order"- disallow change to proposal
                    If (sgCntrScreen = "Orders") And (imStep1Index = 2) And (tmChf.iCntRevNo <= 0) And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Or (tmChf.sStatus = "D")) Then
                        slMsg = "Contract " & slCntrNo & " is a Proposal can't modify on the Order screen "
                        ilFound = False
                    End If
                    'Order "Change Hold to Order"- allow change to hold
                    If (sgCntrScreen = "Orders") And (imStep1Index = 3) And (tmChf.sStatus <> "H") Then  'And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Or (tmChf.sStatus = "D") Or (tmChf.sStatus = "N") Or (tmChf.sStatus = "O")) Then
                        slMsg = "Contract " & slCntrNo & " is Not currently set As Hold"
                        ilFound = False
                    End If
                    '10063
                    If (sgCntrScreen = "Orders") And (imStep1Index = 4) And (tmChf.iCntRevNo <= 0) And ((tmChf.sStatus = "W") Or (tmChf.sStatus = "C") Or (tmChf.sStatus = "I") Or (tmChf.sStatus = "D")) And igWinStatus(PROPOSALSJOB) = 0 Then
                        slMsg = "User does not have access to Proposal #" & slCntrNo & "."
                        ilFound = False
                    End If
                    If Not ilFound Then
                        Beep
                        Screen.MousePointer = vbDefault
                        MsgBox slMsg, vbOKOnly + vbExclamation, "Warning"
                        Contract.pbcHeader.Cls
                        Contract.pbcLine.Cls
                        Contract.pbcLnTotal.Cls
                        Contract.pbcLnDates.Cls
                        Contract.mPaintHdTitle
                        Contract.mPaintLnTitle
                        Contract.mPaintLnDates
                        Contract.mPaintLnTotal
                        imStep3Index = -2
                        Contract.mSetCommands
                        imChgMode = False
                        imInSelect = False
                        igFreezeState = ilSvFreezeState
                        Exit Sub
                    End If
                    '2/17/18: Activated from spot screen allowed to see programmatic buys, 6 and 7 from spot screen
                    If (igAlertCntrStatus <> 6) And (igAlertCntrStatus <> 7) Then
                        ilFound = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode(), tmChf.sSource)
                    Else
                        ilFound = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode(), tmChf.sSource)
                    End If
                    If Not ilFound Then
                        Screen.MousePointer = vbDefault
                        Beep
                        slMsg = "Contract " & slCntrNo & " is Not allowed to be viewed or changed"
                        MsgBox slMsg, vbOKOnly + vbExclamation, "Warning"
                        imStep3Index = -2
                        Contract.mSetCommands
                        imChgMode = False
                        imInSelect = False
                        igFreezeState = ilSvFreezeState
                        Exit Sub
                    End If
                    Contract.lbcStep(2).Clear
                    ReDim tmStep3Code(0 To 1) As SORTCODE   'Contract.lbcStep3Code.Clear
                    'Contract.lbcStep3Code.AddItem Trim$(Str$(tmChf.lCntrNo)) & " R" & Trim$(Str$(tmChf.iCntRevNo)) & " V" & Trim$(Str$(tmChf.iPropVer)) & "\" & Trim$(Str$(tmChf.lCode))
                    tmStep3Code(0).sKey = Trim$(str$(tmChf.lCntrNo)) & " R" & Trim$(str$(tmChf.iCntRevNo)) & "-" & Trim$(str$(tmChf.iExtRevNo)) & " V" & Trim$(str$(tmChf.iPropVer)) & "\" & Trim$(str$(tmChf.lCode))
                    Contract.lbcStep(2).AddItem Trim$(str$(tmChf.lCntrNo)) & " R" & Trim$(str$(tmChf.iCntRevNo)) & "-" & Trim$(str$(tmChf.iExtRevNo)) & " V" & Trim$(str$(tmChf.iPropVer))
                    Contract.lbcStep(2).ListIndex = 0
                    Contract.lbcStep(2).Tag = "1"
                    smStep3CodeTag = "1"    'Contract.lbcStep3Code.Tag = "1"
                    ilRet = 0
                    imStep3Index = 0
                    ilStep3Index = 0
                    Contract.cbcDifference.Visible = False
                    If sgCntrScreen <> "Orders" Then
                        Contract.cmcScale.Visible = True
                    End If
                    Contract.cbcDifference.Clear
                    imDiffMode = False
                Else
                    'ilRet = gOptionalLookAhead(Contract.edcStepDropDown, Contract.lbcStep(2), imBSMode, slStr)
                    imLbcArrowSetting = True
                    gMatchLookAhead Contract.edcStepDropDown, Contract.lbcStep(2), imBSMode, imComboBoxIndex
                    If Contract.lbcStep(2).ListIndex < 0 Then
                        'Beep
                        imStep3Index = -2
                        Contract.mSetCommands
                        Screen.MousePointer = vbDefault
                        imChgMode = False
                        imInSelect = False
                        igFreezeState = ilSvFreezeState
                        Exit Sub
                    End If
                    imStep3Index = Contract.lbcStep(2).ListIndex
                    ilStep3Index = imStep3Index
                    ilRet = 0
                End If
            End If
            Contract.pbcHeader.Cls
            Contract.pbcLine.Cls
            Contract.pbcLnTotal.Cls
            Contract.pbcLnDates.Cls
            Contract.mPaintHdTitle
            Contract.mPaintLnTitle
            Contract.mPaintLnDates
            Contract.mPaintLnTotal
            If ilRet = 0 Then
                If Not Contract.mReadRec(imStep3Index) Then
                    igFreezeState = ilSvFreezeState
                    'GoTo cbcSelectErr
                    imChgMode = False
                    imInSelect = False
                    Screen.MousePointer = vbDefault
                    'cmcCancel_Click
                    imStepBoxNo = -1
                    Contract.pbcStepSTab.SetFocus
                    Exit Sub
                End If
                ilFound = True
                slCntrNo = Trim$(str$(tgChfCntr.lCntrNo))
                'Proposal "Revise Proposal"- disallow change to Order
                If (sgCntrScreen <> "Orders") And (imStep1Index = 2) And (tgChfCntr.iCntRevNo > 0) Then
                    slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                    ilFound = False
                End If
                'Proposal "Revise Proposal"- disallow change to Order
                If (sgCntrScreen <> "Orders") And (imStep1Index = 2) And (tgChfCntr.iCntRevNo = 0) And ((tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "O")) Then
                    slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                    ilFound = False
                End If
                'Proposal "New Version Proposal"- disallow change to Order
                If (sgCntrScreen <> "Orders") And (imStep1Index = 3) And (tgChfCntr.iCntRevNo > 0) Then
                    slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                    ilFound = False
                End If
                'Proposal "New Version Proposal"- disallow change to Order
                If (sgCntrScreen <> "Orders") And (imStep1Index = 3) And (tgChfCntr.iCntRevNo = 0) And ((tgChfCntr.sStatus = "G") Or (tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "O")) Then
                    slMsg = "Contract " & slCntrNo & " is an Order, Use 'Revise an Existing Hold or Order' "
                    ilFound = False
                End If
                'Proposal "Revise Order"- disallow change to proposal
                If (sgCntrScreen <> "Orders") And (imStep1Index = 4) And (tgChfCntr.iCntRevNo <= 0) And ((tgChfCntr.sStatus = "W") Or (tgChfCntr.sStatus = "C") Or (tgChfCntr.sStatus = "I") Or (tgChfCntr.sStatus = "D")) Then
                    slMsg = "Contract " & slCntrNo & " is a Proposal, Use 'Change an Existing Proposal' "
                    ilFound = False
                End If
                'Proposal "Change Hold to Order"- allow change to hold
                If (sgCntrScreen <> "Orders") And (imStep1Index = 5) And (tgChfCntr.sStatus <> "H") Then
                    slMsg = "Contract " & slCntrNo & " is Not currently set As Hold"
                    ilFound = False
                End If
                'Order "Revise Order"- disallow change to proposal
                If (sgCntrScreen = "Orders") And (imStep1Index = 2) And (tgChfCntr.iCntRevNo <= 0) And ((tgChfCntr.sStatus = "W") Or (tgChfCntr.sStatus = "C") Or (tgChfCntr.sStatus = "I") Or (tgChfCntr.sStatus = "D")) Then
                    slMsg = "Contract " & slCntrNo & " is a Proposal can't modify on the Order screen "
                    ilFound = False
                End If
                'Order "Change Hold to Order"- allow change to hold
                If (sgCntrScreen = "Orders") And (imStep1Index = 3) And (tgChfCntr.sStatus <> "H") Then  'And ((tgChfCntr.sStatus = "W") Or (tgChfCntr.sStatus = "C") Or (tgChfCntr.sStatus = "I") Or (tgChfCntr.sStatus = "D") Or (tgChfCntr.sStatus = "N") Or (tgChfCntr.sStatus = "O")) Then
                    slMsg = "Contract " & slCntrNo & " is Not currently set As Hold"
                    ilFound = False
                End If
                If Not ilFound Then
                    igFreezeState = ilSvFreezeState
                    Beep
                    Screen.MousePointer = vbDefault
                    MsgBox slMsg, vbOKOnly + vbExclamation, "Warning"
                    imChgMode = False
                    imInSelect = False
                    imStepBoxNo = -1
                    Contract.lbcStep(2).Visible = False
                    Contract.pbcStepSTab.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass  'Wait
                If Not Contract.mReadClfCffRec() Then
                    igFreezeState = ilSvFreezeState
                    'GoTo cbcSelectErr
                    imChgMode = False
                    imInSelect = False
                    Screen.MousePointer = vbDefault
                    'cmcCancel_Click
                    imStepBoxNo = -1
                    Contract.pbcStepSTab.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass  'Wait
            End If
            Screen.MousePointer = vbHourglass  'Wait
            If Contract.lbcStep(2).Visible Then
                Contract.lbcStep(2).Visible = Not Contract.lbcStep(2).Visible
            End If
            If ilRet = 0 Then
                If imStep1Index = 1 Then    'Model new order
                    '10525
                    If bgModelNTR Then
                        ilSbfRet = mReadSbfRec(hmSbf, hmRvf, hmPhf, imStep1Index)
                    End If
                    tgChfCntr.lCode = 0     'note:mReadSbfRec does not read billing records if modeling
                    tgChfCntr.lCntrNo = 0
                    tgChfCntr.iCntRevNo = 0
                    tgChfCntr.iExtRevNo = 0
                    tgChfCntr.iPropVer = 1
                    tgChfCntr.iOHDDate(0) = 0
                    tgChfCntr.iOHDDate(1) = 0
                    tgChfCntr.iPropDate(0) = 0
                    tgChfCntr.iPropDate(1) = 0
                    lmChfRecPos = 0
                    lmCntrChfCode = 0
                    If sgCntrScreen = "Orders" Then
                        tgChfCntr.sStatus = "N"
                    Else
                        tgChfCntr.sStatus = "W"
                    End If
                    smOrigStatus = ""
                    'imStep3Index = -1  'don't set to -1 because of undo
                ElseIf (sgCntrScreen = "Orders") And (imStep1Index = 2) Then
                    If (tgChfCntr.sStatus = "O") Then   'And (tgChfCntr.sSchStatus <> "M") Then
                        tgChfCntr.sStatus = "N"
                    ElseIf (tgChfCntr.sStatus = "H") Then   'And (tgChfCntr.sSchStatus <> "M") Then
                        tgChfCntr.sStatus = "G"
                    End If
                ElseIf (sgCntrScreen <> "Orders") And (imStep1Index = 4) Then
                    If (tgChfCntr.sStatus = "O") And (tgChfCntr.sSchStatus <> "M") Then
                        tgChfCntr.sStatus = "W"
                    End If
                End If
                mStatusPop tgChfCntr.sStatus, tgChfCntr.iCntRevNo, imStep1Map()
                Screen.MousePointer = vbHourglass  'Wait
                mMoveHdRecToCtrl DATEINDEX, tmChf, tmEff, tmAdf, tmAgf, tmVsf, tmRcf, imTerminate
                If (tgSpf.sMktBase = "Y") And (Not imMktFound) Then
                    igFreezeState = ilSvFreezeState
                    slCntrNo = Trim$(str$(tgChfCntr.lCntrNo))
                    slMsg = "Contract " & slCntrNo & " not part of selected market"
                    Beep
                    Screen.MousePointer = vbDefault
                    MsgBox slMsg, vbOKOnly + vbExclamation, "Warning"
                    imChgMode = False
                    imInSelect = False
                    imStepBoxNo = -1
                    Contract.lbcStep(2).Visible = False
                    Contract.pbcStepSTab.SetFocus
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass  'Wait
                mMoveLnRecToCtrl imTerminate, tmRegionCode()
                Screen.MousePointer = vbHourglass  'Wait
                lmHSStartDate = -1
                Contract.mSetHdDate False, False  'Init Hot values
                Screen.MousePointer = vbHourglass  'Wait
                Contract.mInitDmShow
                mInitCEShow tmCECtrls()
                Contract.mInitSPShow
                Contract.mInitSBShow
                Contract.mInitRSShow
                Contract.mInitPTShow
                Contract.mInitNoShow
                Contract.mInitLnShow
                Screen.MousePointer = vbHourglass  'Wait
            Else
                imStep3Index = -1
                ilStep3Index = -1
                Contract.mClearCtrlFields
            End If
           '10843 remove dead code
'            If tgSpf.sCLnStdQt = "Y" Then
'                imLnTotalType = 2
'                Contract.lacPrice(0).Caption = Contract.lbcQtPrice.List(imLnTotalType)
'            End If
            ilUpdateAllowed = imUpdateAllowed
            imUpdateAllowed = True
            'For ilLoop = LBound(tmHdCtrls) To UBound(tmHdCtrls) Step 1
            For ilLoop = imLBCTCtrls To UBound(tmHdCtrls) Step 1
                Screen.MousePointer = vbHourglass  'Wait
                If (ilLoop <> CNTRNOINDEX) Or (ilRet = 0) Then
                    If ilLoop = MERCHINDEX Then
                        slStr = smComment(0)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = PROMOINDEX Then
                        slStr = smComment(1)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = INTCOMMINDEX Then
                        slStr = smComment(2)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = CANCLAUSEINDEX Then
                        slStr = smComment(3)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = CHGREASONINDEX Then
                        slStr = smComment(4)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = OTHERCOMMINDEX Then
                        slStr = smComment(5)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = STATUSINDEX Then
                        If (ilRet = 0) And (Contract.lbcStatus.ListIndex < 0) And (sgCntrScreen <> "Orders") And ((imStep1Index = 1) Or (imStep1Index = 4)) Then
                            imChgMode = True
                            Contract.lbcStatus.ListIndex = 0
                            imChgMode = True
                        ElseIf (ilRet = 0) And (Contract.lbcStatus.ListIndex < 0) And (sgCntrScreen = "Orders") And ((imStep1Index = 1) Or (imStep1Index = 2)) Then
                            imChgMode = True
                            Contract.lbcStatus.ListIndex = 1
                            imChgMode = True
                        End If
                        If smOrigStatus = "" Then
                            Contract.mSetShow ilLoop  'Set show strings
                        Else
                            If (sgCntrScreen = "Orders") Then
                                If smOrigStatus = "G" Then
                                    slStr = "Approved Hold"
                                ElseIf smOrigStatus = "H" Then
                                    slStr = "Hold"
                                ElseIf smOrigStatus = "N" Then
                                    slStr = "Approved Order"
                                ElseIf smOrigStatus = "W" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Working"
                                    Else
                                        slStr = "Working Proposal"
                                    End If
                                ElseIf smOrigStatus = "C" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Completed"
                                    Else
                                        slStr = "Completed Proposal"
                                    End If
                                ElseIf smOrigStatus = "I" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Unapproved"
                                    Else
                                        slStr = "Unapproved Proposal"
                                    End If
                                ElseIf smOrigStatus = "D" Then
                                    slStr = "Rejected"
                                Else
                                    slStr = "Order"
                                End If
                            Else
                                If smOrigStatus = "G" Then
                                    slStr = "Approved Hold"
                                ElseIf smOrigStatus = "H" Then
                                    slStr = "Hold"
                                ElseIf smOrigStatus = "N" Then
                                    slStr = "Approved Order"
                                ElseIf smOrigStatus = "O" Then
                                    slStr = "Order"
                                ElseIf smOrigStatus = "D" Then
                                    slStr = "Rejected"
                                ElseIf smOrigStatus = "W" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Working"
                                    Else
                                        slStr = "Working Proposal"
                                    End If
                                ElseIf smOrigStatus = "C" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Completed"
                                    Else
                                        slStr = "Completed Proposal"
                                    End If
                                ElseIf smOrigStatus = "I" Then
                                    If tgChfCntr.iCntRevNo > 0 Then
                                        slStr = "Rev Unapproved"
                                    Else
                                        slStr = "Unapproved Proposal"
                                    End If
                                End If
                            End If
                            gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                        End If
                    ElseIf ilLoop = REVNOINDEX Then
                        slStr = Contract.edcRevNo.Text
                        If ((sgCntrScreen = "Orders") And (imStep1Index < 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Then
                            slStr = Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & slStr
                        End If
                        If (slStr = "0") Or (slStr = "0-0") Then
                            slStr = ""
                        End If
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = GUARINDEX Then
                        slStr = smHdSave(2)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = GROSSINDEX Then
                        slStr = smHdSave(1)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    ElseIf ilLoop = PRODTAGINDEX Then
                        slStr = smHdSave(3)
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(ilLoop)
                    Else
                        Contract.mSetShow ilLoop  'Set show strings
                    End If
                Else    'Don't call mSetShow as it sets the contract number of blank
                    slStr = Contract.edcCntrNo.Text
                    gSetShow Contract.pbcHeader, slStr, tmHdCtrls(CNTRNOINDEX)
                End If
            Next ilLoop
            Contract.pbcHeader_Paint
            Contract.pbcLine_Paint
            imUpdateAllowed = ilUpdateAllowed
            igRepNetChgAllowed = False
            If ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                If Trim$(tgChfCntr.sRepDBID) <> "" Then
                    If tgChfCntr.sSchStatus = "F" Then
                        imUpdateAllowed = False
                    Else
                        igRepNetChgAllowed = True
                    End If
                End If
            End If
            '10686
            If tgChfCntr.sDelete = "Y" And mIsRevising() Then
                MsgBox "Switching to View Mode as contract has a newer version", vbOKOnly + vbExclamation, "Warning"
                imUpdateAllowed = False
                mSetBackgroundColor 10  'force to Yellow (view mode)
            End If
            '2/25/21
            If (Asc(tgSaf(0).sFeatures8) And PODADSERVERVIEWONLY) = PODADSERVERVIEWONLY Then
                If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Then
                Else
                    'Test if PIO contract and disallow change mode
                    If (tgChfCntr.sAdServerDefined = "Y") Or ((tgChfCntr.sPodSpotDefined = "Y") And (tgChfCntr.sSource = "I")) Then
                        MsgBox "Switching to View Mode as Contract contains Ad Server and Site Set to Only Allow Viewing", vbOKOnly + vbExclamation, "Warning"
                        imUpdateAllowed = False
                        mSetBackgroundColor 10  'force to Yellow (view mode)
                    End If
                End If
            End If
            mPopPodTarget
            '10525 removed index test from function.  Move here
            If imStep1Index <> 1 Then
                ilSbfRet = mReadSbfRec(hmSbf, hmRvf, hmPhf, imStep1Index)
            End If
            ilSbfRet = mReadMultiMediaRec(hmMsf, hmMgf, hmCxf, imStep1Index)
            Screen.MousePointer = vbHourglass  'Wait
            Contract.udcCBill.Action 4
            Contract.udcCBill.Action 2
            Contract.udcCGameInv.Action 4
            Contract.udcCGameInv.MultiMediaVefCode = imMultiMediaVefCode
            Contract.udcCGameInv.MultiMediaTypeItem = smMultiMediaTypeItem
            Contract.udcCGameInv.Action 2
            Contract.udcCDigital.Action 4
            Contract.udcCDigital.Action 2
            Contract.mPkgTotal
            Screen.MousePointer = vbHourglass  'Wait
            '3/13/20: Only compute research numbers as part of the load if set to continuously compute the numbers
            If ((Asc(tgSaf(0).sFeatures6) And CALCULATESEARCHONLINECHG) = CALCULATESEARCHONLINECHG) Then
                '12/11/19
                Contract.lacAudTotals_Click
            End If
            Contract.mCntrTotal
            Screen.MousePointer = vbHourglass  'Wait
            imComputeReq = False
            Contract.mSetCreditInfo
            Screen.MousePointer = vbHourglass  'Wait
            mSetSchStatus smOrigStatus
            Screen.MousePointer = vbHourglass  'Wait
            igPriDemoMnfDemo = 0
            mSetPriDemoResearch hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating
            Screen.MousePointer = vbHourglass  'Wait
            '3/13/20: add click to title if not computing research number continuously
            If ((Asc(tgSaf(0).sFeatures6) And CALCULATESEARCHONLINECHG) <> CALCULATESEARCHONLINECHG) Then
                If InStr(1, Contract!lacAudTotals.Caption, "Click") <= 0 Then
                    Contract.lacAudTotals.ForeColor = Red
                    Contract.lacAudTotals.Caption = "Click to Compute->" & Contract.lacAudTotals.Caption
                End If
            End If
            '4/3/07: Remove restriction on unscheduled Hold or Order:ttp 2316
            If (imRecLocked) Then
                imUpdateAllowed = False
                Contract.plcPower.Visible = False
                Contract.plcLine.Visible = True
                Contract.pbcLine.Visible = True
                Contract.pbcLnTotal.Visible = True
                Contract.vbcLines.Visible = True
                If Contract.laclnFrame.Visible Then
                    Contract.pbcLineArrow.Visible = True
                Else
                    Contract.pbcLineArrow.Visible = False
                End If
                Contract.cmcRates.Visible = False
            ElseIf (sgCntrScreen = "Orders") And ((smOrigStatus = "W") Or (smOrigStatus = "C") Or (smOrigStatus = "I")) Or (imRecLocked) Then
                imUpdateAllowed = False
                Contract.plcPower.Visible = False
                Contract.plcLine.Visible = True
                Contract.pbcLine.Visible = True
                Contract.pbcLnTotal.Visible = True
                Contract.vbcLines.Visible = True
                'lacLines.Visible = True
                If Contract.laclnFrame.Visible Then
                    Contract.pbcLineArrow.Visible = True
                Else
                    Contract.pbcLineArrow.Visible = False
                End If
                If ((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6)) Or (imRecLocked) Then
                Else
                    ilRet = MsgBox("Contract with Revised Status, can't alter", vbOKOnly + vbExclamation, "Message")
                    mSetBackgroundColor 10  'force to Yellow (view mode)
                End If
                Contract.cmcRates.Visible = False
            Else
                If Not imUpdateAllowed Then
                '    cmcPower.Visible = False
                    Contract.cmcRates.Visible = False
                Else
                        If (Contract.mTestFields(RATECARDINDEX, ALLMANDEFINED + NOMSG) = NO) Or (Contract.mTestFields(DEMOINDEX, ALLMANDEFINED + NOMSG) = NO) Then
                            Contract.cmcRates.Visible = False
                        Else
                            Contract.cmcRates.Visible = True
                        End If
                    '4/3/07:  reset contract status as Complete

                    If (sgCntrScreen <> "Orders") And ((smOrigStatus = "G") Or (smOrigStatus = "N")) Then
                        ilCntRevNo = tgChfCntr.iCntRevNo
                        tgChfCntr.sStatus = "C"
                        smOrigStatus = tgChfCntr.sStatus    'Save for erase test and mSaveRec
                        If tgChfCntr.iCntRevNo <= 0 Then
                            tgChfCntr.iCntRevNo = 1
                        End If
                        mStatusPop tgChfCntr.sStatus, tgChfCntr.iCntRevNo, imStep1Map()
                        mMoveHdRecToCtrl DATEINDEX, tmChf, tmEff, tmAdf, tmAgf, tmVsf, tmRcf, imTerminate
                        slStr = "Rev Completed"
                        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(STATUSINDEX)
                        tgChfCntr.iCntRevNo = ilCntRevNo
                        Contract.pbcHeader.Cls
                        Contract.pbcHeader_Paint
                    End If
                End If
            End If
            If (imStep1Index = 1) And (imUpdateAllowed) Then    'Model new order
                imLnChg = True
            Else
                imLnChg = False
            End If
            Contract.mSetCommands
            slStr = Contract.edcStepDropDown.Text
            ilRet = gParseItem(slStr, 1, " ", slCntrNo)
            Contract.mSetDifference slCntrNo
            Screen.MousePointer = vbHourglass  'Wait
            Contract.mSetRevNo
            imFirstTimeStep = False
            'Check if selection changed while processing current selection- if so, redo
        'Loop While (imStep3Index <> Contract.lbcStep(2).ListIndex) And ((imStep3Index <> 0) Or (Contract.lbcStep(2).ListIndex >= 0))
            ilStepChgd = False
            If (ilStep1Index <> imStep1Index) Or (ilStep2Index <> ilStep2Index) Or (ilStep3Index <> imStep3Index) Then
                ilStepChgd = True
            Else
                If (imStep3Index <> Contract.lbcStep(2).ListIndex) And ((imStep3Index <> 0) Or (Contract.lbcStep(2).ListIndex >= 0)) Then
                    ilStepChgd = True
                Else
                    If (imStep2Index <> Contract.lbcStep(1).ListIndex) And ((imStep2Index <> 0) Or (Contract.lbcStep(1).ListIndex >= 0)) Then
                        ilStepChgd = True
                    End If
                End If
            End If
        Loop While ilStepChgd
        Screen.MousePointer = vbDefault
        If (Contract.lbcStep(2).ListCount = 1) And (Contract.lbcStep(2).Visible) Then
            Contract.lbcStep(2).Visible = False
        End If
        '10951
        mSetDefaultTab
        imChgMode = False
    End If
    imInSelect = False
    igFreezeState = ilSvFreezeState
    Exit Sub

    On Error GoTo 0
    'If sgCntrScreen = "Orders" Then
    '    gShowHelpMess tmChfHelp(), CHFCNTRSELECT
    'Else
    '    gShowHelpMess tmChfHelp(), CHFPROPSELECT
    'End If
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    imInSelect = False
    Exit Sub
End Sub

Private Sub mSetPriceInTab(PRICEINDEX As Integer)
    Dim slChgPrice As String
    slChgPrice = tgUrf(0).sChgPrices
    tgUrf(0).sChgPrices = "I"
    If imLnRowNo >= UBound(tgClfCntr) + 1 Then
        bmSettingPriceInTab = True
        imLnBoxNo = PRICEINDEX
        Contract.mLnEnableBox PRICEINDEX
        Contract.mLnSetShow PRICEINDEX
        bmSettingPriceInTab = False
    End If
    tgUrf(0).sChgPrices = slChgPrice
End Sub

Private Function mObtainAudioTypeDefault(slVehicleName As String) As String
    Dim ilVefCode As Integer
    Dim ilRet As Integer
    Dim ilVff As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    mObtainAudioTypeDefault = ""
    gFindMatch Trim$(slVehicleName), 0, Contract.lbcLnVehicle(igTabMapIndex)
    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) < 0 Then
        ilVefCode = 0
    Else
        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = CInt(slCode)
    End If
    If ilVefCode > 0 Then
        ilVff = gBinarySearchVff(ilVefCode)
        If ilVff <> -1 Then
            mObtainAudioTypeDefault = tgVff(ilVff).sDefaultAudioType
        End If
    End If

End Function

Public Sub mUnloadSub()
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = btrClose(hmVbf)
    btrDestroy hmVbf
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
    Else
        ilRet = btrClose(hmAsf)
        btrDestroy hmAsf
    End If
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilRet = btrClose(hmLcf)
        btrDestroy hmLcf
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        ilRet = btrClose(hmPsf)
        btrDestroy hmPsf
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
    End If

    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf

    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Or ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        btrDestroy hmGif
    End If
    btrDestroy hmPvf
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmDef)
    btrDestroy hmDef
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    ilRet = btrClose(hmCef)
    btrDestroy hmCef
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hmCmf)
    btrDestroy hmCmf
    ilRet = btrClose(hmMsf)
    btrDestroy hmMsf
    ilRet = btrClose(hmMgf)
    btrDestroy hmMgf
    ilRet = btrClose(hmGhf)
    btrDestroy hmGhf
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    ilRet = btrClose(hmEff)
    btrDestroy hmEff
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    dnf_rst.Close
    
End Sub

Public Sub mAvailTab(Optional ilInLnRowNo As Integer = -1)
    Dim slLineType As String
    Dim ilRdfCode As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDay As Integer
    ReDim ilDays(0 To 6) As Integer
    ReDim ilAvDays(0 To 6) As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    Dim llSpotCount As Long
    Dim ilAvail As Integer
    Dim ilInvCount As Integer
    Dim ilCompUnitSold As Integer
    Dim ilNoBreak As Integer
    Dim ilNoNBBBreak As Integer
    Dim ilAdfCode As Integer
    Dim ilAdfCodeCount As Integer
    Dim ilCompBreakCount As Integer
    Dim ilLoop As Integer
    
    Dim llMoLnSDate As Long
    Dim llLnSDate As Long
    Dim llLnEDate As Long
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilVefCode As Integer
    Dim llCompStartTime As Long
    Dim llCompEndTime As Long
    Dim llUpper As Long
    Dim ilWeekIndex As Integer
    Dim llNext As Long
    Dim llPrevNext As Long
    Dim ilLnRowNo As Integer
    Dim ilStartRowNo As Integer
    Dim ilEndRowNo As Integer
    Dim llLatestDate As Long

    Dim ilVpf As Integer
    Dim slLength As String
    Dim llCompTime As Long
    Dim ilSvToggleAvail As Integer
    
    ilSvToggleAvail = imToggleAvail
    imToggleAvail = 0
    
    ilAdfCode = 0
    If Contract.lbcAdvt.ListIndex >= 1 Then
        slNameCode = tmCntAdvertiser(Contract.lbcAdvt.ListIndex - 1).sKey  'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        slCode = Trim$(slCode)
        ilAdfCode = Val(slCode)
    End If
    If ilInLnRowNo < 0 Then
        'ilStartRowNo = LBound(tgClfCntr) + 1
        'ilEndRowNo = UBound(tgClfCntr)
        'If (smLnSave(1, UBound(tgClfCntr) + 1) <> "") And (smLnSave(16, UBound(tgClfCntr) + 1) <> "") Then
        '    ilEndRowNo = UBound(tgClfCntr) + 1
        'Else
        '    ilEndRowNo = UBound(tgClfCntr)
        'End If
        ilStartRowNo = Contract.vbcLines.Value + 1  'Top location
        ilEndRowNo = Contract.vbcLines.Value + Contract.vbcLines.LargeChange + 1
        If ilEndRowNo > UBound(tgClfCntr) + 1 Then
            ilEndRowNo = UBound(tgClfCntr) + 1 'include blank row as it might have data
        End If
        ilStartRowNo = mRemapRowNo(ilStartRowNo)
        ilEndRowNo = mRemapRowNo(ilEndRowNo)
    Else
        ilStartRowNo = ilInLnRowNo
        ilEndRowNo = ilInLnRowNo
    End If
    ilLnRowNo = ilStartRowNo
    Do While ilLnRowNo <= ilEndRowNo
        slLineType = mGetLineType(ilLnRowNo)
        igComp1Code = 0
        If Contract.lbcComp(0).ListIndex >= 2 Then
            slNameCode = tgCompCode(Contract.lbcComp(0).ListIndex - 2).sKey   'Traffic!lbcCompCode.List(lbcComp(0).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                igComp1Code = Val(Trim$(slCode))
            End If
        End If
        igComp2Code = 0
        If Contract.lbcComp(1).ListIndex >= 2 Then
            slNameCode = tgCompCode(Contract.lbcComp(1).ListIndex - 2).sKey   'Traffic!lbcCompCode.List(lbcComp(1).ListIndex - 2)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If ilRet = CP_MSG_NONE Then
                igComp2Code = Val(Trim$(slCode))
            End If
        End If
        If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
            mDPAllowedDays imLnSave(1, ilLnRowNo), ilDays()
        Else
            For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If StrComp(smLnSave(20, ilLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                    mDPAllowedDays tgMRdf(ilLoop).iCode, ilDays()
                    Exit For
                End If
            Next ilLoop
        End If
        For ilDay = 0 To 6 Step 1
            ilAvDays(ilDay) = ilDays(ilDay)
        Next ilDay
        If imLnSave(4, ilLnRowNo) = 1 Then
            llOvStartTime = CLng(gTimeToCurrency(smLnSave(2, ilLnRowNo), False))
            llOvEndTime = CLng(gTimeToCurrency(smLnSave(3, ilLnRowNo), True))
        Else
            llOvStartTime = 0
            llOvEndTime = 0
        End If
        If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
            ilRdfCode = imLnSave(1, ilLnRowNo)
            gFindMatch smLnSave(1, ilLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
            If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) < 0 Then
                ilVefCode = 0
            Else
                slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = CInt(slCode)
            End If
        Else
            ilVefCode = 0
            For ilLoop = LBound(tgPackVeh) To UBound(tgPackVeh) - 1 Step 1
                If StrComp(Trim$(smLnSave(1, ilLnRowNo)), Trim$(tgPackVeh(ilLoop).sName), 1) = 0 Then
                    ilVefCode = tgPackVeh(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
            For ilLoop = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                If StrComp(smLnSave(20, ilLnRowNo), Trim$(tgMRdf(ilLoop).sName), 1) = 0 Then
                    ilRdfCode = tgMRdf(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
        End If
        If lmHSStartDate = -1 Then
            If imLnRowNo < UBound(tgClfCntr) + 1 Then   'New lines set after all fields entered
                Contract.mSetHdDate False, False
            Else
                Contract.mSetHdDate True, False
            End If
        End If
        ''llLnSDate = lmHSStartDate   'gDateValue(smLnSave(14, ilLnRowNo))
        'get all weeks from start to end. if only do the week looking at require that the array tgLnAvailInfo would need to be reconstructed to add  in missing previous weeks
        'llLnSDate = lmHSStartDate + 7 * (imHSCurWkNo - 1) 'gDateValue(smLnSave(14, ilLnRowNo))
        llLnSDate = lmHSStartDate
        'llLnEDate = lmHSStartDate + 7 * (imHSCurWkNo - 1))
        llLnEDate = lmHSStartDate + 7 * (imHSCurWkNo + 12) - 1 '7 * 13  '104 'gDateValue(smLnSave(15, ilLnRowNo))
        'This code was in to eliminate looking into weeks past the last created lcf week  It failed when start date was past the last lcf date
        llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
        If (llLatestDate < llLnEDate) And (llLatestDate > 0) Then
            If llLnSDate < llLatestDate Then
                llLnEDate = llLatestDate
            End If
        End If
        llMoLnSDate = gObtainPrevMondayLong(llLnSDate)
        For llDate = llMoLnSDate To llLnEDate Step 7
            If llDate < llLnSDate Then
                llSDate = llLnSDate
            Else
                llSDate = llDate
            End If
            llEDate = llDate + 6
            'If llEDate > llLnEDate Then
            '    llEDate = llLnEDate
            'End If
            
            If tgClfCntr(ilLnRowNo - 1).lFirstAvail = -1 Then
                llUpper = UBound(tgLnAvailInfo)
                ReDim Preserve tgLnAvailInfo(0 To llUpper + 1) As LNAVAILINFO
                tgClfCntr(ilLnRowNo - 1).lFirstAvail = llUpper
                llNext = llUpper
                For ilLoop = 0 To UBound(tgLnAvailInfo(0).iSAvail) Step 1
                    tgLnAvailInfo(llNext).iSAvail(ilLoop) = -29999
                Next ilLoop
                tgLnAvailInfo(llNext).sLineType = slLineType
                tgLnAvailInfo(llNext).lStartDate = llMoLnSDate
                tgLnAvailInfo(llNext).lNextIndex = -1
                ilWeekIndex = (llDate - tgLnAvailInfo(llNext).lStartDate) \ 7
            Else
                llNext = tgClfCntr(ilLnRowNo - 1).lFirstAvail
                ilWeekIndex = (llDate - tgLnAvailInfo(llNext).lStartDate) \ 7
                Do
                    If ilWeekIndex <= UBound(tgLnAvailInfo(0).iSAvail) Then
                        Exit Do
                    Else
                        llPrevNext = llNext
                        llNext = tgLnAvailInfo(llNext).lNextIndex
                        If llNext = -1 Then
                            llUpper = UBound(tgLnAvailInfo)
                            ReDim Preserve tgLnAvailInfo(0 To llUpper + 1) As LNAVAILINFO
                            tgLnAvailInfo(llPrevNext).lNextIndex = llUpper
                            llNext = llUpper
                            For ilLoop = 0 To UBound(tgLnAvailInfo(0).iSAvail) Step 1
                                tgLnAvailInfo(llNext).iSAvail(ilLoop) = -29999
                            Next ilLoop
                            tgLnAvailInfo(llNext).sLineType = slLineType
                            tgLnAvailInfo(llNext).lStartDate = llDate
                            tgLnAvailInfo(llNext).lNextIndex = -1
                            'Exit Do
                        'Else
                        '    ilWeekIndex = ilWeekIndex - (UBound(tgLnAvailInfo(0).iAvail) + 1)
                        End If
                        ilWeekIndex = ilWeekIndex - (UBound(tgLnAvailInfo(0).iSAvail) + 1)
                    End If
                Loop
            End If
            If ilWeekIndex < 0 Then
                tgClfCntr(ilLnRowNo - 1).lFirstAvail = -1
                llUpper = UBound(tgLnAvailInfo)
                ReDim Preserve tgLnAvailInfo(0 To llUpper + 1) As LNAVAILINFO
                tgClfCntr(ilLnRowNo - 1).lFirstAvail = llUpper
                llNext = llUpper
                For ilLoop = 0 To UBound(tgLnAvailInfo(0).iSAvail) Step 1
                    tgLnAvailInfo(llNext).iSAvail(ilLoop) = -29999
                Next ilLoop
                tgLnAvailInfo(llNext).sLineType = slLineType
                tgLnAvailInfo(llNext).lStartDate = llMoLnSDate
                tgLnAvailInfo(llNext).lNextIndex = -1
                ilWeekIndex = (llDate - tgLnAvailInfo(llNext).lStartDate) \ 7
            End If
            If slLineType <> "O" And slLineType <> "A" And slLineType <> "E" Then
                If (tgLnAvailInfo(llNext).iSAvail(ilWeekIndex) = -29999) Then
                    If tgSpf.sCIncludeMissDB = "N" Then
                        mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(ilLnRowNo - 1).ClfRec.iLine, ilVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, ilLnRowNo)), ilAvDays(), llSDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoBreak, ilNoNBBBreak, llCompStartTime, llCompEndTime, False, ilAdfCode, ilAdfCodeCount, ilCompBreakCount
                    Else
                        mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(ilLnRowNo - 1).ClfRec.iLine, ilVefCode, ilRdfCode, llOvStartTime, llOvEndTime, Val(smLnSave(16, ilLnRowNo)), ilAvDays(), llSDate, llEDate, llSpotCount, ilAvail, ilInvCount, ilCompUnitSold, ilNoBreak, ilNoNBBBreak, llCompStartTime, llCompEndTime, True, ilAdfCode, ilAdfCodeCount, ilCompBreakCount
                    End If
                    tgLnAvailInfo(llNext).iInv(ilWeekIndex) = ilInvCount
                    tgLnAvailInfo(llNext).lSpotCount(ilWeekIndex) = llSpotCount
                    tgLnAvailInfo(llNext).iSAvail(ilWeekIndex) = ilAvail
                    ilVpf = gBinarySearchVpf(tgClfCntr(ilLnRowNo - 1).ClfRec.iVefCode)
                    If ilVpf <> -1 Then
                        If tgVpf(ilVpf).sSCompType = "T" Then   'Time
                            tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) = ilCompBreakCount - ilCompUnitSold
                        ElseIf tgVpf(ilVpf).sSCompType = "B" Then   'Break
                            tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) = ilNoBreak - ilCompUnitSold
                        Else
                            'Not in same break
                            tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) = ilNoNBBBreak - ilCompUnitSold
                        End If
                        If tgVpf(ilVpf).sAdvtSep = "T" Then
                        Else
                            If ilNoBreak - ilAdfCodeCount < tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) Then
                                tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) = ilNoBreak - ilAdfCodeCount
                            End If
                        End If
                    End If
                End If
            End If
        Next llDate
        ilLnRowNo = ilLnRowNo + 1
    Loop
    imToggleAvail = ilSvToggleAvail
End Sub

Public Sub mResearchForTabs(VEHINDEX As Integer, DPINDEX As Integer, Optional ilInLnRowNo As Integer = 0)
    '10431 store old LineRowNo
    Dim ilStoreLine As Integer
    '12/5/19: add test
    If ((Asc(tgSaf(0).sFeatures6) And CALCULATESEARCHONLINECHG) = CALCULATESEARCHONLINECHG) Then
        imDBTotals = 6  'Weekly summary
        ilStoreLine = imLnRowNo
        imLnRowNo = ilInLnRowNo
        mTSummary imLnRowNo, hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imDBTotals, False, imCppCpm, imCntrVefCode, imVpfIndex, imTerminate, imChgMode, imInResearch, VEHINDEX, DPINDEX, imStep1Index, False, False
        imLnRowNo = ilStoreLine
    Else
        If InStr(1, Contract!lacAudTotals.Caption, "Click") <= 0 Then
            Contract.lacAudTotals.ForeColor = Red
            Contract.lacAudTotals.Caption = "Click to Compute->" & Contract.lacAudTotals.Caption
        End If
    End If
End Sub

Public Sub mLnBookNamePop(ilLnRowNo As Integer)
    Dim slSQLQuery As String
    Dim ilIndex As Integer
    Dim ilVefCode As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slStr As String
    
    Contract.lbcLnProgram(igTabMapIndex).Clear
    ilIndex = -1
    slStr = ""
    gFindMatch smLnSave(1, ilLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
    If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) < 0 Then
        ilVefCode = 0
    Else
        slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = CInt(slCode)
    End If
    slSQLQuery = "Select distinct dnfBookName, dnfBookDate, dnfCode From DRF_Demo_Rsrch_Data"
    slSQLQuery = slSQLQuery & " Inner Join DNF_Demo_Rsrch_Names On drfDnfCode = dnfCode "
    slSQLQuery = slSQLQuery & " Where drfVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " Order By dnfBookDate, dnfBookName"
    Set dnf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not dnf_rst.EOF
        Contract.lbcLnProgram(igTabMapIndex).AddItem Trim$(dnf_rst!dnfBookName) & ": " & Format(dnf_rst!dnfBookDate, sgShowDateForm)
        Contract.lbcLnProgram(igTabMapIndex).ItemData(Contract.lbcLnProgram(igTabMapIndex).NewIndex) = dnf_rst!dnfCode
        If tgClfCntr(ilLnRowNo - 1).ClfRec.iDnfCode = dnf_rst!dnfCode Then
            slStr = Contract.lbcLnProgram(igTabMapIndex).List(Contract.lbcLnProgram(igTabMapIndex).NewIndex)
        End If
        dnf_rst.MoveNext
    Loop
    If slStr <> "" Then
        gFindMatch slStr, 0, Contract.lbcLnProgram(igTabMapIndex)
        If gLastFound(Contract.lbcLnProgram(igTabMapIndex)) >= 0 Then
            Contract.lbcLnProgram(igTabMapIndex).ListIndex = gLastFound(Contract.lbcLnProgram(igTabMapIndex))
        End If
    End If
End Sub

Public Sub mSetDnfCode(ilOldDnfCode As Integer)
    Dim ilRet As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilClf As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilTerminate As Integer
    
    ilStartRow = imLB1Or2   'LBound(smLnSave, 2)
    If smLnSave(1, UBound(smLnSave, LINEBOUNDINDEX)) = "" Then
        ilEndRow = UBound(smLnSave, LINEBOUNDINDEX) - 1
    Else
        ilEndRow = UBound(smLnSave, LINEBOUNDINDEX)
    End If
    ilStartRow = ilStartRow - 1
    ilEndRow = ilEndRow - 1
    For ilClf = ilStartRow To ilEndRow Step 1
        If (ilOldDnfCode <> 0) And (tgClfCntr(ilClf).ClfRec.iDnfCode = ilOldDnfCode) Then
            mCVehOptIndex ilClf + 1, ilVefCode, ilVpfIndex, ilTerminate
            If ilVpfIndex <> -1 Then
                tmDrfSrchKey.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode
                tmDrfSrchKey.sDemoDataType = "D"    'Demo Data
                tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                tmDrfSrchKey.iVefCode = ilVefCode
                tmDrfSrchKey.sInfoType = ""        'Daypart
                tmDrfSrchKey.iRdfCode = 0
                ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tmDrf.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode) And (tmDrf.iVefCode = ilVefCode) Then
                    tgClfCntr(ilClf).ClfRec.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode
                    tgClfCntr(ilClf).lPop = 0
                    tgClfCntr(ilClf).iResearch = False
                    imLnChg = True
                Else
                    tmDrfSrchKey.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode
                    tmDrfSrchKey.sDemoDataType = "M"    'Demo Data
                    tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                    tmDrfSrchKey.iVefCode = ilVefCode
                    tmDrfSrchKey.sInfoType = ""        'Daypart
                    tmDrfSrchKey.iRdfCode = 0
                    ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tmDrf.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode) And (tmDrf.iVefCode = ilVefCode) Then
                        tgClfCntr(ilClf).ClfRec.iDnfCode = tgClfCntr(imLnRowNo - 1).ClfRec.iDnfCode
                        tgClfCntr(ilClf).lPop = 0
                        tgClfCntr(ilClf).iResearch = False
                        imLnChg = True
                    End If
                End If
            End If
        End If
    Next ilClf
End Sub

Public Function mGetWeekToolTip(ilRowNo As Integer, llDate As Long, ilIgnoreRightMove As Integer, Ctrl As Control) As String
    Dim llNext As Long
    Dim ilWeekIndex As Integer
    Dim slLineType As String
    
    If tgClfCntr(ilRowNo - 1).lFirstAvail = -1 Then
        Screen.MousePointer = vbHourglass
        bmInAvailTab = True
        mAvailTab ilRowNo
        bmInAvailTab = False
        Screen.MousePointer = vbDefault
    End If
    
    llNext = tgClfCntr(ilRowNo - 1).lFirstAvail
    If llNext = -1 Then
        mGetWeekToolTip = ""
        ilIgnoreRightMove = False
        Exit Function
    End If
    slLineType = mGetLineType(ilRowNo, False, True)
    If ((slLineType = "A") Or (slLineType = "O") Or (slLineType = "E")) Then
        mGetWeekToolTip = ""
        ilIgnoreRightMove = False
        Exit Function
    End If
    
    ilWeekIndex = (llDate - tgLnAvailInfo(llNext).lStartDate) \ 7
    Do While llNext <> -1
        If ilWeekIndex <= UBound(tgLnAvailInfo(0).iSAvail) Then
            Exit Do
        Else
            ilWeekIndex = ilWeekIndex - (UBound(tgLnAvailInfo(0).iSAvail) + 1)
            llNext = tgLnAvailInfo(llNext).lNextIndex
        End If
    Loop
    If llNext <> -1 Then
        If tgLnAvailInfo(llNext).iSAvail(ilWeekIndex) <> -29999 Then
            If (lmToolTipNext <> llNext) Or (imToolTipWeekIndex <> ilWeekIndex) Then
                Ctrl.ToolTipText = ""
            End If
            lmToolTipNext = llNext
            imToolTipWeekIndex = ilWeekIndex
            'If (Contract.tscLine.SelectedItem.Index = 8) Then
            If (Contract.tscLine.SelectedItem.Index = imTabMap(TABAVAIL)) Then
                mGetWeekToolTip = tgLnAvailInfo(llNext).iSAvail(ilWeekIndex)
            'ElseIf (Contract.tscLine.SelectedItem.Index = 9) Then
            ElseIf (Contract.tscLine.SelectedItem.Index = imTabMap(TABPROTECTION)) Then
                If tgLnAvailInfo(llNext).iSAvail(ilWeekIndex) < tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex) Then
                    mGetWeekToolTip = tgLnAvailInfo(llNext).iSAvail(ilWeekIndex)
                Else
                    mGetWeekToolTip = tgLnAvailInfo(llNext).iAPAvail(ilWeekIndex)
                End If
            Else
                mGetWeekToolTip = ""
            End If
        Else
            mGetWeekToolTip = ""
        End If
    Else
        mGetWeekToolTip = ""
    End If
    ilIgnoreRightMove = False
End Function

Public Sub mClearAvails(ilRowNo As Integer)
    Dim ilClf As Integer
    Dim llNext As Long
    Dim ilWeek As Integer
    Dim ilLnMax As Integer
    
    If ilRowNo = -1 Then
        If (smLnSave(1, UBound(tgClfCntr) + 1) <> "") And (smLnSave(16, UBound(tgClfCntr) + 1) <> "") Then
            ilLnMax = UBound(tgClfCntr)
        Else
            ilLnMax = UBound(tgClfCntr) - 1
        End If
        For ilClf = LBound(tgClfCntr) To ilLnMax Step 1
            llNext = tgClfCntr(ilClf).lFirstAvail
            Do While llNext <> -1
                For ilWeek = LBound(tgLnAvailInfo(llNext).iSAvail) To UBound(tgLnAvailInfo(llNext).iSAvail) Step 1
                    tgLnAvailInfo(llNext).iSAvail(ilWeek) = -29999
                Next ilWeek
                llNext = tgLnAvailInfo(llNext).lNextIndex
            Loop
        Next ilClf
    Else
        llNext = tgClfCntr(ilRowNo - 1).lFirstAvail
        Do While llNext <> -1
            For ilWeek = LBound(tgLnAvailInfo(llNext).iSAvail) To UBound(tgLnAvailInfo(llNext).iSAvail) Step 1
                tgLnAvailInfo(llNext).iSAvail(ilWeek) = -29999
            Next ilWeek
            llNext = tgLnAvailInfo(llNext).lNextIndex
        Loop
    End If
    '10843 remove dead code
  '  If tgSpf.sCLnStdQt <> "Y" Then  'disallow avails/protectiona nd research tab if using Quarters. Problem is that both use imLnTotalType value 2
        'If (Contract.tscLine.SelectedItem.Index = 8) Or (Contract.tscLine.SelectedItem.Index = 9) Then
        If (Contract.tscLine.SelectedItem.Index = imTabMap(TABAVAIL)) Or (Contract.tscLine.SelectedItem.Index = imTabMap(TABPROTECTION)) Then
            mAvailTab ilRowNo
        End If
  '  End If
End Sub

Public Sub mClearResearch(ilRowNo As Integer, ilVehIndex As Integer, ilDPIndex As Integer)
    Dim ilClf As Integer
    Dim llNext As Long
    Dim ilWeek As Integer
    Dim ilLnMax As Integer
    
    If sgCntrScreen = "Orders" Then
        Exit Sub
    End If
    If ilRowNo = -1 Then
        If (smLnSave(1, UBound(tgClfCntr) + 1) <> "") And (smLnSave(16, UBound(tgClfCntr) + 1) <> "") Then
            ilLnMax = UBound(tgClfCntr)
        Else
            ilLnMax = UBound(tgClfCntr) - 1
        End If
        For ilClf = LBound(tgClfCntr) To ilLnMax Step 1
            tgClfCntr(ilClf).iResearch = False
        Next ilClf
    Else
        tgClfCntr(ilRowNo - 1).iResearch = False
    End If
   '10843 remove dead code
   ' If tgSpf.sCLnStdQt <> "Y" Then  'disallow avails/protectiona nd research tab if using Quarters. Problem is that both use imLnTotalType value 2
        'If ((Contract.tscLine.SelectedItem.Index = 8) Or (Contract.tscLine.SelectedItem.Index = 9) Or (Contract.tscLine.SelectedItem.Index = 10)) And (tgSpf.sGUsePropSys = "Y") Then
        If ((Contract.tscLine.SelectedItem.Index = imTabMap(TABAVAIL)) Or (Contract.tscLine.SelectedItem.Index = imTabMap(TABPROTECTION)) Or (Contract.tscLine.SelectedItem.Index = imTabMap(TABRESEARCH))) And (tgSpf.sGUsePropSys = "Y") Then
            If ilRowNo = -1 Then
                mResearchForTabs ilVehIndex, ilDPIndex
            Else
                mResearchForTabs ilVehIndex, ilDPIndex, ilRowNo
            End If
        End If
   ' End If
End Sub

Public Sub mSetPriceShow(slStr As String, ilPRICEINDEX As Integer)
    Dim lfBoxW As Single
    
    If imLnTotalType <> 2 Then
        gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilPRICEINDEX)
    Else
        lfBoxW = tmLnCtrls(ilPRICEINDEX).fBoxW
        tmLnCtrls(ilPRICEINDEX).fBoxW = imResearchTogglePriceWidth
        gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilPRICEINDEX)
        tmLnCtrls(ilPRICEINDEX).fBoxW = lfBoxW
    End If

End Sub

Public Function mIsPackageForActOne(slLineType As String, ilLnNumber As Integer, Optional llPvfCode As Long = 0) As Boolean
'10256 10325 added optional
    Dim blRet As Boolean
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilVef As Integer
    
    blRet = False
    If imACT1CodesDefined Then
'        'it's already defined as a package
        If (slLineType = "A") Or (slLineType = "O") Or (slLineType = "E") Then
            blRet = True
        Else
            'entering a saved contract with 'AdServer' screen open causes error
On Error GoTo ERRBOX
            'test by vehicle name
            gFindMatch smLnSave(1, ilLnNumber), 0, Contract.lbcLnVehicle(igTabMapIndex)
            If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                If igTabMapIndex = 1 Then
                    slNameCode = tmVehicleCode1(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey
                ElseIf igTabMapIndex = 2 Then
                    slNameCode = tmVehicleCode2(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey
                Else
                    slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey
                End If
                gParseItem slNameCode, 2, "\", slCode
                ilCode = CInt(slCode)
                ilVef = gBinarySearchVef(ilCode)
                If ilVef <> -1 Then
                    If (tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0) Then
                        blRet = True
                        llPvfCode = tgMVef(ilVef).lPvfCode
                    End If
                End If
            End If
        End If
    End If
ERRBOX:

    mIsPackageForActOne = blRet
End Function

'10325
Private Sub mObtainAct1DefaultsForPackages(ilLnRow As Integer)
    'smLnSave  29 = lineup imLnSave 32= stored time 33= spots 34= ClearPct 35=filter
    Dim llPvfCode As Long
    Dim slSQLQuery As String
    Dim myRst As ADODB.Recordset
    Dim slLineup As String
    Dim ilTime As Integer
    Dim ilSpots As Integer
    Dim ilClear As Integer
    Dim ilFilter As Integer
    Dim slLineType As String
    
    'we don't want it to go in as a package, because it then won't get the pvfcode.  It is currently seen as 'S' where this function is called (mLnSetShowSub)
    slLineType = "S"
    If mIsPackageForActOne(slLineType, ilLnRow, llPvfCode) Then
        If llPvfCode > 0 Then
            slSQLQuery = "Select pvfACT1DaypartFilter as filter, pvfACT1LineupCode as lineup, pvfACT1StoreClearPct as clear, pvfACT1StoredSpots as spots, pvfACT1StoredTime as Time from PVF_Package_Vehicle WHERE pvfcode = " & llPvfCode
            Set myRst = gSQLSelectCall(slSQLQuery)
            If Not myRst.EOF Then
                If Not IsNull(myRst!lineup) Then
                    slLineup = Trim$(myRst!lineup)
                End If
                If Not IsNull(myRst!Filter) Then
                    ilFilter = -1
                    If myRst!Filter = "F" Then
                        ilFilter = 0
                    End If
                End If
                If Not IsNull(myRst!Clear) Then
                    ilClear = -1
                    If myRst!Clear = "C" Then
                        ilClear = 0
                    End If
                End If
                If Not IsNull(myRst!Spots) Then
                    ilSpots = -1
                    If myRst!Spots = "S" Then
                        ilSpots = 0
                    End If
                End If
                If Not IsNull(myRst!Time) Then
                    ilTime = -1
                    If myRst!Time = "T" Then
                        ilTime = 0
                    End If
                End If
            End If
            smLnSave(29, ilLnRow) = slLineup
            imLnSave(32, ilLnRow) = ilTime
            imLnSave(33, ilLnRow) = ilSpots
            imLnSave(34, ilLnRow) = ilClear
            imLnSave(35, ilLnRow) = ilFilter
        End If  'pvfcode
    End If 'is package
End Sub

'10347
Private Sub mBlockHiddenLineWriting()
    imCurrHightLightLnRowNo = -1
    Contract.pbcLine_Paint
    imPrevHightLightLnRowNo = -1
    MsgBox "Can't Alter Hidden Line", vbOKOnly + vbExclamation, "Warning"
    imLnRowNo = -1
    Contract.pbcClickFocus.SetFocus
End Sub
Private Function mIsRevising() As Boolean
    '10686
    Dim blRet As Boolean
    blRet = False
    If imStepBoxNo = 3 Then
        If sgCntrScreen = "Orders" Then
            If imStep1Index = 2 Then
                blRet = True
            End If
        ElseIf imStep1Index = 4 Then
            blRet = True
        End If
    End If
    mIsRevising = blRet
End Function
Private Sub mAdjustDeliveryStatus()
    '10967
    Dim ilVendor As Integer
    Dim ilMegaphone As Integer
    Dim slVendor As String
    Dim ilDeliveryStatus As Integer
    Dim llRow As Long
    Dim ilIndex As Integer
    Dim dlVendorInContract As Dictionary
    Dim olIndex As Variant
    
    Const VENDORMEGAPHONE = "Megaphone"
    'copied from contracts form
    Const ADVTINDEX = 8     'Advertiser control/index
    Const PRODUCTINDEX = 9  'Product control/index

On Error GoTo mRecErr
    Set dlVendorInContract = New Dictionary
    'modeling
     If imStep1Index = 1 Then
        tgChfCntr.iAvfCode = 0
        tgChfCntr.iAdServDlvyStatus = DeliveryStatus_NA
    End If
    'go through line to determine if contract tied to a vendor
    For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.rows - 1 Step 1
        'Dick's test that row has a record in it-then test if any change for that vendor
        If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
            If Not mIsCPMPackageVehicle(llRow) Then
                 'modeling-remove the 'sent to this vendor'
                 If imStep1Index = 1 Then
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMAVFCODE) = 0
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS) = DeliveryStatus_NA
                    'these don't get copied over, but being thorough
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGVEHICLE) = ""
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGCAP) = ""
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGPOSITION) = ""
                    Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGTOTAL) = ""
                End If
                slVendor = ""
                ilDeliveryStatus = Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS))
                ilVendor = mGetAvfCodeForVehicle(llRow)
                If ilVendor > 0 Then
                    If dlVendorInContract.Exists(ilVendor) Then
                        slVendor = dlVendorInContract(ilVendor)
                    Else
                        For ilIndex = 0 To UBound(tmDigitalVendorsToTest) - 1
                            If tmDigitalVendorsToTest(ilIndex).iCode = ilVendor Then
                                slVendor = tmDigitalVendorsToTest(ilIndex).sName
                                dlVendorInContract.Add ilVendor, slVendor
                                Exit For
                            End If
                        Next
                    End If
                End If
                'we didn't find a vendor we care about.
                If Len(slVendor) = 0 Then
                    ilVendor = 0
                End If
                'any change to cpms?
                If imCPMChg = -1 Then
                    'change of vehicle or new vehicle
                    If Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGVEHICLE) Then
                        'new or megaphone vehicle to another megaphone vehicle
                        If Len(slVendor) > 0 Then
                            mSetVendorToPushOrRepush llRow
                        'changed from vendor vehicle to one that is not this vendor.  But already sent(have avfcode) Need to send a delete
                        ElseIf Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMAVFCODE)) <> ilVendor Then
                            mSetVendorToPushOrRepush llRow
                        'never sent and now changed to non-vendor.
                        ElseIf ilVendor = 0 Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS) = DeliveryStatus_NA
                        End If
                    'test changed to megaphone fields on an update
                    Else
                        Select Case slVendor
                            Case VENDORMEGAPHONE
                                If ilDeliveryStatus > DeliveryStatus_NotPushed And ilDeliveryStatus < DeliveryStatus_RequiresRepush Then
                                    'Cap,Position,Priorityimpressions,startdate,enddate,totalcost
                                    If Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONCAP)) <> Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGCAP)) Or Contract.grdCPMBuy.TextMatrix(llRow, CPMPOSITION) <> Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGPOSITION) Or Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMPRIORITY)) <> Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGPRIORITY)) Or Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS)) <> Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGIMPRESSIONS)) Or Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) <> Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGSTARTDATE) Or Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) <> Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGENDDATE) Or Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST)) <> Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMORIGTOTAL)) Then
                                        'just a repush
                                        mSetVendorToPushOrRepush llRow
                                    End If
                                End If
                        End Select
                    End If 'change to vehicle or other fields
                'modeling doesn't set imCPMChg to true
                ElseIf imStep1Index = 1 And Len(slVendor) > 0 Then
                    mSetVendorToPushOrRepush llRow
                End If 'any change to lines
            End If 'not a package
        Else
            Exit For
        End If 'valid line
    Next
    For Each olIndex In dlVendorInContract
        slVendor = dlVendorInContract(olIndex)
        Select Case slVendor
            Case VENDORMEGAPHONE
                If tgChfCntr.iAdServDlvyStatus > DeliveryStatus_NotPushed And ilDeliveryStatus < DeliveryStatus_RequiresRepush Then
                    'should only need to test if line didn't reset,but it's quick
                    If tmHdCtrls(PRODUCTINDEX).iChg Or tmHdCtrls(ADVTINDEX).iChg Then
                        mSetVendorToPushOrRepush -1
                    End If
                End If
        End Select
    Next
    'no vendor in contract
    If Len(slVendor) = 0 Then
        tgChfCntr.iAdServDlvyStatus = DeliveryStatus_NA
    End If
    Exit Sub
mRecErr:
    gLogMsg "Digital vendor error in contracts-mAdjustDeliveryStatus. Contract not set to resend", "TrafficErrors.Txt", False
End Sub
Private Function mGetAvfCodeForVehicle(llRow As Long) As Integer
    Dim ilRet As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVff As Integer
    
On Error GoTo ERRBOX
    ilRet = 0
    ilVefCode = CInt(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE))
    If ilVefCode > 0 Then
        ilVff = gBinarySearchVff(ilVefCode)
        If ilVff <> -1 Then
            ilRet = tgVff(ilVff).iAvfCode
        End If
    End If
    mGetAvfCodeForVehicle = ilRet
    Exit Function
ERRBOX:
    ilRet = 0
    gLogMsg "Digital vendor error in contracts-mGetAveCodeByVehicle. Contract not set to resend", "TrafficErrors.Txt", False
End Function
Private Function mIsCPMPackageVehicle(llRow As Long) As Boolean
    Dim blRet As Boolean
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    
On Error GoTo ERRBOX
    blRet = False
    ilVefCode = CInt(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEFCODE))
    If ilVefCode > 0 Then
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            If (tgMVef(ilVef).sType = "P") Then
                blRet = True
            End If
        End If
    End If
   mIsCPMPackageVehicle = blRet
    Exit Function
ERRBOX:
    blRet = True
    gLogMsg "Digital vendor error in contracts-mIsCPMPackageVehicle. Contract not set to resend", "TrafficErrors.Txt", False
End Function
Private Sub mSetVendorToPushOrRepush(llRow As Long)
    ' not testing DeliveryStatus_NotPushed or DeliveryStatus_RequiresRepush as they don't change
    If llRow > -1 Then
        Select Case Val(Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS))
            'sent'
            Case DeliveryStatus_Pushed, DeliveryStatus_Partial, DeliveryStatus_IssueEncountered
                Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS) = DeliveryStatus_RequiresRepush
            Case DeliveryStatus_NA
                Contract.grdCPMBuy.TextMatrix(llRow, CPMDELIVERYSTATUS) = DeliveryStatus_NotPushed
        End Select
    End If
    Select Case tgChfCntr.iAdServDlvyStatus
        'sent'
        Case DeliveryStatus_Pushed, DeliveryStatus_Partial, DeliveryStatus_IssueEncountered
             tgChfCntr.iAdServDlvyStatus = DeliveryStatus_RequiresRepush
        Case DeliveryStatus_NA
             tgChfCntr.iAdServDlvyStatus = DeliveryStatus_NotPushed
    End Select
End Sub
Private Function mGetVendorsAPI() As VendorInfo()
    Dim tlVendorsToReturn() As VendorInfo
    Dim tlTempVendors() As VendorInfo
    Dim ilIndex As Integer
    Dim ilSubIndex As Integer
    Dim blIsActive As Boolean

    ReDim tlVendorsToReturn(0)
    tlTempVendors = gGetDigitalVendors(False)
    For ilIndex = 0 To UBound(tlTempVendors) - 1
        If tlTempVendors(ilIndex).oContractMethod = ContractMethodType.Vendor Then
            tlVendorsToReturn(UBound(tlVendorsToReturn)) = tlTempVendors(ilIndex)
            ReDim Preserve tlVendorsToReturn(UBound(tlVendorsToReturn) + 1)
        End If
    Next
    mGetVendorsAPI = tlVendorsToReturn
End Function
Private Sub mSetDefaultTab()
    '10951 the podcast tab is green and the airtime tab is black
    If Contract.tscLine.Tabs(imTabMap(TABPODCASTCPM)).Image = 8 And Contract.tscLine.Tabs(imTabMap(TABAIRTIME)).Image = 5 Then
        Contract.tscLine.Tabs(imTabMap(TABPODCASTCPM)).Selected = True
    Else
        Contract.tscLine.Tabs(imTabMap(TABAIRTIME)).Selected = True
    End If
End Sub
