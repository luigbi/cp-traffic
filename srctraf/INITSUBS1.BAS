Attribute VB_Name = "INITSUBS1"
' Copyright 1993 Counterpoint Software ® All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: InitSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Initialize subs and functions
Option Explicit
Option Compare Text

Type DRIVEPATH
    sFolder As String
    iPos As Integer
End Type

Global sgCommand As String


'*******************************************************
'*                                                     *
'*      Procedure Name:gEraseGlobalVar                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Erase global variables          *
'*                                                     *
'*******************************************************
Sub gEraseGlobalVar(ilEraseAll As Integer)
    Erase tgAdvertiser
    sgAdvertiserTag = ""
    Erase tgAgency
    sgAgencyTag = ""
    'Erase tgColAdvertiser
    'sgColAdvertiserTag = ""
    'Erase tgCopyAdvertiser
    'sgCopyAdvertiserTag = ""
    'Erase tgSpotAdvertiser
    'sgSpotAdvertiserTag = ""
    'Erase tmCntAdvertiser
    'smCntAdvertiserTag = ""
    'Erase tmCntAgency
    'smCntAgencyTag = ""
    'Erase tmCntRateCard
    'smCntRateCardTag = ""
    'Erase tmVehicleCode
    'smVehicleCodeTag = ""
    'Erase tmCntSalesperson
    'smCntSalespersonTag = ""
    'Erase tmCntBuyerCode
    'smCntBuyerCodeTag = ""
    Erase tgAvailCode
    sgAvailCodeTag = ""
    Erase tgBookNameCode
    sgBookNameCodeTag = ""
    Erase tgBusCatCode
    sgBusCatCodeTag = ""
    Erase tgBuyerCode
    sgBuyerCodeTag = ""
    Erase tgCommentCode
    sgCommentCodeTag = ""
    Erase tgCompCode
    sgCompCodeTag = ""
    Erase tgDemoCode
    sgDemoCodeTag = ""
    Erase tgCDemoCode
    sgCDemoCodeTag = ""
    Erase tgExclCode
    sgExclCodeTag = ""
    Erase tgManager
    sgManagerTag = ""
    Erase tgNameCode
    sgNameCodeTag = ""
    Erase tgPayableCode
    sgPayableCodeTag = ""
    Erase tgPlanner
    sgPlannerTag = ""
    Erase tgPotCode
    sgPotCodeTag = ""
    Erase tgSegCode
    sgSegCodeTag = ""
    Erase tgProdCode
    sgProdCodeTag = ""
    Erase tgRateCardCode
    sgRateCardCodeTag = ""
    Erase tgResvCode
    sgResvCodeTag = ""
    Erase tgRevSetCode
    sgRevSetCodeTag = ""
    Erase tgSalesperson
    sgSalespersonTag = ""
    Erase tgShareBudgetCode
    sgShareBudgetCodeTag = ""
    Erase tgSocEcoCode
    sgSocEcoCodeTag = ""
    Erase tgSPersonCombo
    sgSPersonComboTag = ""
    Erase tgShtTitleCode
    sgShtTitleCodeTag = ""
    Erase tgTmpSort
    sgTmpSortTag = ""
    Erase tgUser
    sgUserTag = ""
    Erase tgPopUrf
    sgPopUrfTag = ""
    Erase tgUserVehicle
    sgUserVehicleTag = ""
    Erase tgRCUserVehicle
    sgRCUserVehicleTag = ""
    Erase tgPopDnfInfo
    sgPopDnfInfoTag = ""

    Erase tgVehCombo
    sgVehComboTag = ""
    Erase tgVehicle
    sgVehicleTag = ""
    Erase tgLogVehicle
    sgLogVehicleTag = ""
    Erase tgMktCode     '2-22-02
    sgMktCode = ""
    Erase tgMNFCodeRpt  '10-18-02
    sgMNFCodeTagRpt = ""

    Erase tgVpf
    sgVpfStamp = ""

    Erase tgSOCode
    sgSOCodeTag = ""
    Erase lgVehComboCode
    sgVsfStamp = ""
    Erase igVirtVefCode
    sgVVefStamp = ""

    Erase tgMVef
    sgMVefStamp = ""

    Erase tgMCof
    sgMCofStamp = ""

    Erase tgMRcf
    Erase tgMRif
    Erase tgMRdf
    sgMRcfStamp = ""
    sgMRifStamp = ""
    sgMRdfStamp = ""

    Erase tgClf
    Erase tgCff
    Erase tgClfCntr
    Erase tgCffCntr
    Erase tgClfSchd
    Erase tgCffSchd
    Erase tgClfImpt
    Erase tgCffImpt
    Erase tgFBSbf
    Erase tgIBSbf
    Erase tgMBSbf
    Erase tgPBSbf
    Erase tgClfCntrProj
    Erase tgCffCntrProj
    Erase tgClfBud
    Erase tgCffBud
    Erase tgClfPostLog
    Erase tgCffPostLog
    Erase tgClfCopy
    Erase tgCffCopy
    Erase tgClfRC
    Erase tgCffRC
    Erase tgClfSpot
    Erase tgCffSpot
    Erase tgClfInv
    Erase tgCffInv

    Erase tgClfAA
    Erase tgCffAA
    Erase tgClfAC
    Erase tgCffAC
    Erase tgClfAD
    Erase tgCffAD
    Erase tgClfAP
    Erase tgCffAP
    Erase tgClfAS
    Erase tgCffAS
    Erase tgClfAV
    Erase tgCffAV
    Erase tgClfCB
    Erase tgCffCB
    Erase tgClfCP
    Erase tgCffCP
    Erase tgClfCT
    Erase tgCffCT
    Erase tgClfED
    Erase tgCffED
    Erase tgClfPA
    Erase tgCffPA
    Erase tgClfPC
    Erase tgCffPC
    Erase tgClfPS
    Erase tgCffPS
    Erase tgClfRI
    Erase tgCffRI
    Erase tgClfSP
    Erase tgCffSP
    Erase tgClfUS
    Erase tgCffUS

    Erase tgCommAgf
    sgCommAgfStamp = ""
    Erase tgCommAdf
    sgCommAdfStamp = ""
    Erase tgCompMnf
    sgCompMnfStamp = ""
    Erase tgRevSet
    sgRevSetStamp = ""
    Erase tgPotMnf
    sgPotMnfStamp = ""
    Erase tgBusCatMnf
    sgBusCatMnfStamp = ""
    Erase tgExclMnf
    sgExclMnfStamp = ""
    Erase tgShareBudget
    sgShareBudgetStamp = ""
    Erase tgDemoMnf
    sgDemoMnfStamp = ""
    Erase tgSocEcoMnf
    sgSocEcoMnfStamp = ""
    Erase tgAvailAnf
    sgAvailAnfStamp = ""

    Erase tgPrg

    Erase tgRPrg

    Erase igEstByUSADnfCode
    Erase igMktCode 'Market code
    Erase igMktVefCode
    Erase igLogMktCode
    Erase igCntrMktCode

    If ilEraseAll Then
        Erase tgUrf
        sgUrfStamp = ""
        Erase tgJobHelp
        Erase tgListHelp
        Set ogContractCreator = Nothing
    End If


End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gInitGlobalVar                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize global variables     *
'*                                                     *
'*******************************************************
Sub gInitGlobalVar()
'
'   gInitGlobalVar
'
    Dim ilNoFiles As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim llVer As Long
    Dim flVerMajor As Single
    Dim flVerMinor As Single
    Dim flWinVersion As Single
    Dim ilTraffic1 As Integer   'Number of controls in traffic1
    Dim ilMultiNm As Integer
    'Dim slVersion As String
    'Dim slChar As String
    'Dim ilPos As Integer

        '9459
    bgUseRadioButtons = False
    '5/2/19: Set count to 0, Only open and close schedule files when count is zero
    
    lgOpenSchFileCount = 0

    sgApplyVehPreemptRule = "N"
        sgVehPreemptRule = "N"
        
    sgCSVersion = "Version 5.1" ' created 10/02/01 at 12:00PM"
    'sgCSVersion = "Version 4.6 created 4/24/00 at 2:00PM"
    'sgSystemDate = Format$(gNow(), "m/d/yy")    'Used to reset date when exiting traffic
    sgCurDir = CurDir$
    'igResetSystemDate = False
    igNoMinForms = 1    'Number of min forms: Traffic
    ilNoFiles = SetHandleCount(60)
'    sgDBPath = "c:\traffic\data\"   'Data base path

    igRUDirectMode = 0
    igRUStartExeCopy = False
    '5666  only do this once
    If Len(sgShortDate) = 0 Then
        sgShortDate = gRegistryGetShortDate()
        If Len(sgShortDate) > 0 And sgShortDate <> MYSHORTDATE Then
            gRegistrySetShortDate (MYSHORTDATE)
        End If
    End If
    gObtainIniValues
    'slVersion = "V"
    'ilPos = InStr(sgCSVersion, " ") + 1
    'slChar = Mid$(sgCSVersion, ilPos, 1)
    'Do While slChar <> " "
    '    slVersion = slVersion & slChar
    '    ilPos = ilPos + 1
    '    slChar = Mid$(sgCSVersion, ilPos, 1)
    'Loop
    'If igTestSystem Then
    '    sgBrannerMsg = "Counterpoint ® TEST, " & slVersion
    'Else
    '    sgBrannerMsg = "Counterpoint ®, " & slVersion
    'End If
    gSetBrannerName

    sgSUName = ""
    sgCPName = ""

    sgPasswordAddition = ""


    For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
        igWinStatus(ilLoop) = 0
    Next ilLoop

    igStopCancel = False
    igExitTraffic = False
    igFromSignon = False
    igManUnload = False
    igNoMenuDisabled = 0    'Number of times mune disabled
    'lgDateCreditShown = 0   'Date that credit advt/agy screen shown

    sgDefaultTerms = "15 Days Upon Receipt"

    sgShowDateForm = "m/d/yy"

    igCalActive = 0 'False
    fgCalLeft = 100
    fgCalTop = 200
    igCalMonth = Month(gNow())
    igCalYear = Year(gAdjYear(Format$(gNow(), "m/d/yy")))
    igCalMonthType = 0  'Standard
    igCalByYear = False
    igCalAcross = True

    igCalcActive = 0 'False
    fgCalcLeft = Traffic.Width - 2775 - 100   '2625 width of calc form
    fgCalcTop = 200

    fgPadDeltaY = 255
    fgPadDeltaX = 330
    fgPadMinY = 15
    fgPadMinX = 15

    fgListHtArial825 = 210  '225  'Height of list item (arial 8.25 bold; DivideStyle=1-Black line; ListStyle=0-2D)
    fgListHtSerif825 = 195  '255  'Height of list item (san serif 8.25; Divide style=3=Raised; ListStyle=1-3D)
    igAvgCharWidth = 90
    igScrollBarWidth = 255

    sgBS = Chr$(8)      'Backspace
    sgTB = Chr$(9)      'Tab
    sgLF = Chr$(10)     'Line feed (newline)
    sgCR = Chr$(13)     'Carriage return
    sgRUCrLf = Chr$(141) & Chr$(138)

    sgCnCDelimiter = Chr$(9)    'Tab
    sgCnCNameBracket = ""   ' or 4 quotes to make a quote """"

    fgOffset = 150      'Distance from box outline to top of control
    fgBoxInsetX = 30    'X Margin from box outline to text
    fgBoxInsetY = 120   'Y Margin from box outline to text
    fgBoxStH = 330      'Height of standard box
    fgBoxAddH = 675     'Height of address box
    fgBoxNmAddH = 1020  'Height of name and address box
    fgBoxGridH = 180      'Height of grid area (distance from bottom of form letter to bottom of form box)
    fgBevelX = 45       'Bevel width in X direction on panel
    fgBevelY = 45       'Bevel width in Y direction on panel
    fgXBoxToBox = 45    'Distance from one Box (ctrl) to the next in the x-direction
    fgStDeltaY = 345    'Distance from standard control top to control top
    fgAddDeltaY = 690   'Distance from address control to top of control
    fgAddNmDeltaY = 1035    'Distance from address/name control to top of control
    fgPanelAdj = 90    'Adjustment to window around panel
    fgBoxNoLabelH = 240 'Height of area without label
    igAlignCharWidth = 90
    lgPercentAdjH = 95
    lgPercentAdjW = 95

    igPopAdfAgfDormant = True
    igPopExternalAdvt = True

    sgBrowserTitle = ""

    igMNmCallSource = CALLNONE
    igVehCallSource = CALLNONE
    igETypeCallSource = CALLNONE
    igENameCallSource = CALLNONE
    igANmCallSource = CALLNONE
    igMcdCallSource = CALLNONE
    igCmmCallSource = CALLNONE
    igSofCallSource = CALLNONE
    igArfCallSource = CALLNONE
    igSlfCallSource = CALLNONE
    'igVsfCallSource = CALLNONE
    igAgyCallSource = CALLNONE
    igAdvtCallSource = CALLNONE
    igAdvtProdCallSource = CALLNONE
    igVehOptCallSource = CALLNONE
    igPersonnelCallSource = CALLNONE
    sgMNmName = ""
    sgVehName = ""
    sgETypeName = ""
    sgENameName = ""
    sgANmName = ""
    sgMcdName = ""
    sgCmmTitle = ""
    sgArfName = ""
    sgAgyName = ""
    sgAdvtName = ""
    sgSlfName = ""
    sgAdvtProdName = ""
    sgPersonnelName = ""
    'sgVsfName = ""

    igSSFBaseLen = 20

    igShowCatOrNames = 0    'Initialize by Categories

    igEtfRecsCreated = False    'Event type records created flag

    igLegalTime(0) = Asc("A")     'A
    igLegalTime(1) = Asc("a")     'a
    igLegalTime(2) = Asc("P")     'P
    igLegalTime(3) = Asc("p")    'p
    igLegalTime(4) = Asc("N")     'N for noon
    igLegalTime(5) = Asc("M")     'M
    igLegalTime(6) = Asc("m")    'm
    igLegalTime(7) = Asc("n")    'n
    igLegalTime(8) = Asc(":")     ':

    igLegalLength(0) = Asc(":")   ':
    igLegalLength(1) = Asc("'")   ''
    igLegalLength(2) = Asc("""")   '"
    igLegalLength(3) = Asc(" ")   'space
    igLegalLength(4) = Asc("H")
    igLegalLength(5) = Asc("h")
    igLegalLength(6) = Asc("M")
    igLegalLength(7) = Asc("m")
    igLegalLength(8) = Asc("S")
    igLegalLength(9) = Asc("s")

    igEditBoxMaxCharacters = 0
    lgEditBoxMaxValue = 0

    igLibType = 0   'Library
    igViewType = 0  'On air

    igSlfFirstNameFirst = False

    sgVpfStamp = ""
    sgMVefStamp = ""
    sgMktClusterDef = "N"
    sgRepDef = "N"
    ReDim igMktVefCode(0 To 0) As Integer
    ReDim igRepVefCode(0 To 0) As Integer

    ReDim igEstByUSADnfCode(0 To 0) As Integer

    sgVehicleTag = ""
    sgLogVehicleTag = ""
    'sgCAgfStamp = ""
    'sgCAdfStamp = ""
    'sgCAdfAllStamp = ""
    'sgCAgfCxfStamp = ""
    'sgCSlfStamp = ""
    'sgCCdfStamp = ""
    sgCommAgfStamp = ""
    sgCommAdfStamp = ""
    'sgCommUrfStamp = ""
    sgCompMnfStamp = ""

    'If (GetWinFlags() And 16384) > 0 Then   'NT
    '    flWinVersion = 4#
    '    fgWinVersion = flWinVersion
    'Else
        llVer = GetVersion() And &HFFFF&
        flVerMajor = llVer Mod 256
        flVerMinor = (llVer \ 256) / 100
        flWinVersion = flVerMajor + flVerMinor
        fgWinVersion = flWinVersion
    'End If
    If flWinVersion >= 3.1 Then
        If flWinVersion > 3.9 Then
            igSystemResources = 2000    'Arbitrary number
        Else
            igSystemResources = 600
        End If
    Else
        igSystemResources = 475
    End If

    If lgRUTimeOut <= 0 Then
        lgRUTimeOut = 900
        sgRUTimeOut = "15 Minutes"
    End If

'    igNoLJWinRes(RATECARDSJOB) = 78 'Number of controls plus 1 for main window
'    igNoLJWinRes(PROPOSALSJOB) = 166
'    igNoLJWinRes(CONTRACTSJOB) = 166
'    igNoLJWinRes(PROGRAMMINGJOB) = 35
'    igNoLJWinRes(SPOTSJOB) = 114
'    igNoLJWinRes(COPYJOB) = 45
'    igNoLJWinRes(LOGSJOB) = 23
'    igNoLJWinRes(POSTLOGSJOB) = 40
'    igNoLJWinRes(INVOICESJOB) = 21
'    igNoLJWinRes(COLLECTIONSJOB) = 72
'    igNoLJWinRes(BUDGETSJOB) = 80
'    igNoLJWinRes(SLSPCOMMSJOB) = 45
'    igNoLJWinRes(STATIONFEEDJOB) = 50

    ilTraffic1 = 10
    ilMultiNm = 50
'    igNoLJWinRes(VEHICLESLIST) = 29 + ilTraffic1
'    igNoLJWinRes(VEHICLEGROUPSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(AGENCIESLIST) = 49 + ilTraffic1 'Advertiser plus traffic1
'    igNoLJWinRes(ADVERTISERSLIST) = 83 + ilTraffic1
'    igNoLJWinRes(COMPETITIVESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(BUSCATEGORIESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(POTENTIALCODESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(EXCLUSIONSLIST) = ilMultiNm + ilTraffic1
'
'    igNoLJWinRes(SALESSOURCESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(SALESREGIONSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(SALESOFFICESLIST) = 32 + ilTraffic1
'    igNoLJWinRes(SALESTEAMSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(SALESPEOPLELIST) = 38 + ilTraffic1
'    igNoLJWinRes(RESEARCHLIST) = 30 + ilTraffic1
'    igNoLJWinRes(DEMOSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(COMPETITORSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(REVENUESETSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(BOILERPLATESLIST) = 30 + ilTraffic1
'
'    igNoLJWinRes(EVENTTYPESLIST) = 25 + ilTraffic1
'    igNoLJWinRes(EVENTNAMESLIST) = 35 + ilTraffic1
'    igNoLJWinRes(AVAILNAMESLIST) = 21 + ilTraffic1
'    igNoLJWinRes(FEEDTYPESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(GENRESLIST) = ilMultiNm + ilTraffic1
'
'    igNoLJWinRes(MEDIADEFINITIONSLIST) = 30 + ilTraffic1
'    igNoLJWinRes(ANNOUNCERNAMESLIST) = ilMultiNm + ilTraffic1
'
'    igNoLJWinRes(MISSEDREASONSLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(ITEMBILLINGTYPESLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(INVOICESORTLIST) = ilMultiNm + ilTraffic1
'    igNoLJWinRes(LOCKBOXESLIST) = 25 + ilTraffic1
'    igNoLJWinRes(EDISERVICESLIST) = 25 + ilTraffic1
'    igNoLJWinRes(TRANSACTIONSLIST) = ilMultiNm + ilTraffic1
'
'    igNoLJWinRes(SITELIST) = 145
'    igNoLJWinRes(USERLIST) = 116
'
'    igNoExeWinRes(RCTERMSEXE) = 93   'Rate card terms
'    igNoExeWinRes(GRIDCALCEXE) = 9  'Grid calculations
'    igNoExeWinRes(CAVAILEXE) = 22   'Avails for contract
'    igNoExeWinRes(CBILLEXE) = 48    'billing via contract
'    igNoExeWinRes(CCANCELEXE) = 11  'Cancel lines
'    igNoExeWinRes(CRESALLOEXE) = 21 'Reservation/Allocation Contract
'    igNoExeWinRes(CREVNOEXE) = 9    'Revision contract number
'    igNoExeWinRes(CNTRSCHEXE) = 9 + ilTraffic1 'Contract scheduling
'    igNoExeWinRes(PEVENTEXE) = 49   'Program events
'    igNoExeWinRes(PRGDATESEXE) = 24   'Program events
'    igNoExeWinRes(LINKSEXE) = 35 + ilTraffic1   'Links
'    igNoExeWinRes(LINKDLVYEXE) = 27   'Delivery links
'    igNoExeWinRes(LINKSDEFEXE) = 17   'Link definitions
'    igNoExeWinRes(PRGSCHEXE) = 9 + ilTraffic1 'Program scheduling
'    igNoExeWinRes(COPYINVEXE) = 16  'Copy inventory
'    igNoExeWinRes(COPYRATOEXE) = 19 'Copy ratio
'    igNoExeWinRes(COPYSPOTEXE) = 14 'Spots by copy
'    igNoExeWinRes(PURGEEXE) = 30
'    igNoExeWinRes(POSTADJTEXE) = 19
'    igNoExeWinRes(POSTITEMEXE) = 31
'    igNoExeWinRes(VEHOPTEXE) = 167  'Vehicle options
'    igNoExeWinRes(ADVTPRODEXE) = 18 + ilTraffic1 'Spots by copy
'    igNoExeWinRes(RPTNOSELEXE) = 25 + 18 + 3 + ilTraffic1 'Report+rptnosel
'    igNoExeWinRes(RPTSELEXE) = 98 + 3 + ilTraffic1 'Spots by copy
'    igNoExeWinRes(CALCEXE) = 9
'    igNoExeWinRes(CALENDAREXE) = 15
'    igNoExeWinRes(LOCKSEXE) = 19
'    igNoExeWinRes(PRGDELEXE) = 16
'    igNoExeWinRes(UNSCHDEXE) = 25
'    igNoExeWinRes(COPYASGNEXE) = 23  'Copy assign
'    igNoExeWinRes(RPTSELCTEXE) = 112 + 3 + ilTraffic1 'RptSelCT

'    sgTDBNames(1) = "AVR.BTR"
'    sgTDBNames(2) = "CBF.BTR"
'    sgTDBNames(3) = "CPR.BTR"
'    sgTDBNames(4) = "GRF.BTR"
'    sgTDBNames(5) = "JSR.BTR"
'    sgTDBNames(6) = "ODF.BTR"
'    sgTDBNames(7) = "RVR.BTR"
'    sgTDBNames(8) = "SWF.BTR"
'    sgTDBNames(9) = "ANR.BTR"
'    sgTDBNames(10) = "SVR.BTR"
'    sgTDBNames(11) = "IVR.BTR"
'    sgTDBNames(12) = "TXR.BTR"

    ReDim tgUrf(0 To 0) As URF

    'csiInitName

    'hgDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    Do While csiHandleValue(0, 3) = 0
        '7/6/11
        Sleep 1000
    Loop

    'Orlando meeting- ignore help messages
    'hgHlf = CBtrvTable(ONEHANDLE)
    'For ilLoop = 0 To 20 Step 1
    '    ilRet = ilRet
    'Next ilLoop
    'ilRet = btrOpen(hgHlf, "", sgDBPath & "Hlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'gHlfRead "JOB", tgJobHelp()
    'gHlfRead "LST", tgListHelp()

'   Missed reason
    igMnfMissed = 0
    igPLAnfCode = 0

    ReDim igMktCode(0 To 0) As Integer
    lgMtfNoRecs = 0
    lgDpfNoRecs = 0
    
    sgSQLDateForm = "yyyy-mm-dd"
    sgSQLTimeForm = "hh:mm:ss"
    
    Exit Sub
                  
    ilRet = err.Number
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gInitStdAlone                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Code required for stand alone   *
'*                     applications- from Traffic and  *
'*                     signon                          *
'*                                                     *
'*******************************************************
Sub gInitStdAlone(frm As Form, slUser As String, ilTestSystem As Integer)
    Dim ilRet As Integer
    Dim slVehType As String
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilRecLen As Integer

    If igStdAloneMode Then
        ilRet = csiSetAlloc("NAMES", 0, 2)
    End If
    igTestSystem = ilTestSystem
    sgSystemDate = gAdjYear(Format$(gNow(), "m/d/yy"))    'Used to reset date when exiting traffic
    igResetSystemDate = False
    gInitGlobalVar   'Initialize global variables
    ReDim tgJobHelp(0 To 0) As HLF
    ReDim tgListHelp(0 To 0) As HLF
    gSpfRead
    sgCPName = gGetCSIName("CPNAME")
    sgSUName = gGetCSIName("SUNAME")
    If (sgCPName = "") Or (sgSUName = "") Then
        hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
            ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUrfDecrypt tlUrf
                If tlUrf.iCode = 1 Then
                    sgCPName = Trim$(tlUrf.sName)
                End If
                ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUrfDecrypt tlUrf
                    If tlUrf.iCode = 2 Then
                        sgSUName = Trim$(tlUrf.sName)
                    End If
                End If
            End If
        End If
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        ilRet = csiSetName("CPName", sgCPName)
        ilRet = csiSetName("SUName", sgSUName)
    End If
    igUpdateAllowed = True
    sgUserName = slUser
    gUrfRead frm, sgUserName, True, tgUrf(), False  'Obtain user records
    sgUserName = Trim$(tgUrf(0).sName)
    If igRUDirectMode = 1 Then
        tgUrf(0).iRemoteUserID = 0
        'Disallow Contract Projections in Direct Mode (1/4/2000)
        igWinStatus(PROPOSALSJOB) = 1
    ElseIf (igRUDirectMode = 2) And (tgUrf(0).iRemoteUserID > 0) Then
        igWinStatus(PROPOSALSJOB) = 1
    End If
    gObtainVehicleName tgUrf(0).iDefVeh, sgUserDefVehicleName, slVehType
    sgUserSalespersonName = gGetCSIName("SLSPNAME")
    If sgUserSalespersonName = "" Then
        gObtainSalespersonName tgUrf(0).iSlfCode, sgUserSalespersonName
        ilRet = csiSetName("SLSPName", sgUserSalespersonName)
    End If
    ilRet = gVpfRead()    'Obtain preference table
    'gObtainVehComboList
    gObtainVirtVehList
    gInitSuperUser tgUrf(0)
    If tgUrf(0).sClnMoYr = "M" Then
        igCalByYear = False
    Else
        igCalByYear = True
    End If
    Select Case tgUrf(0).sClnType
        Case "S"
            igCalMonthType = 0
        Case "R"
            igCalMonthType = 1
        Case "U"
            igCalMonthType = 2
        Case "D"
            igCalMonthType = 3
        Case "C"
            igCalMonthType = 4
    End Select
    If tgUrf(0).sClnLayout = "A" Then
        igCalAcross = True
    Else
        igCalAcross = False
    End If
    fgCalLeft = tgUrf(0).iClnLeft
    fgCalTop = tgUrf(0).iClnTop
    fgCalcLeft = tgUrf(0).iClcLeft
    fgCalcTop = tgUrf(0).iClcTop

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gInitSuperUser                  *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize Restrictions for     *
'*                     superuser which might not to set*
'*                                                     *
'*******************************************************
Sub gInitSuperUser(tlUrf As URF)
    If (Trim$(tlUrf.sName) = sgCPName) Or (Trim$(tlUrf.sName) = sgSUName) Then
        tlUrf.iDefVeh = 0
        tlUrf.sGrid = "I"
        tlUrf.sPrice = "I"
        tlUrf.sCredit = "I"
        tlUrf.sPayRate = "I"
        tlUrf.sMerge = "I"
        tlUrf.sDelete = "N"
        tlUrf.sHideSpots = "I"
        tlUrf.sChgBilled = "I"
        tlUrf.sChgCntr = "I"
        tlUrf.sChgCrRt = "I"
        tlUrf.sBouChk = "I"
        'If (Trim$(tlUrf.sName) = sgCPName) Then
            tlUrf.sReprintLogAlert = "Y"
            tlUrf.sIncompAlert = "Y"
            tlUrf.sCompAlert = "Y"
            tlUrf.sSchAlert = "Y"
            tlUrf.sHoldAlert = "Y"
            tlUrf.sRateCardAlert = "Y"
            tlUrf.sResearchAlert = "Y"
            tlUrf.sAvailAlert = "Y"
            tlUrf.sCrdChkAlert = "Y"
            tlUrf.sDeniedAlert = "Y"
            tlUrf.sCrdLimitAlert = "Y"
            tlUrf.sMoveAlert = "Y"
            tlUrf.sPrgmmaticAlert = "I"
            tlUrf.sAdvanceAvails = "I"
            tlUrf.sShowNRMsg = "Y"
        'End If
        tlUrf.sWorkToDead = "Y"
        tlUrf.sWorkToComp = "Y"
        tlUrf.sWorkToHold = "Y"
        tlUrf.sWorkToOrder = "Y"
        tlUrf.sCompToIncomp = "Y"
        tlUrf.sCompToDead = "Y"
        tlUrf.sCompToHold = "Y"
        tlUrf.sCompToOrder = "Y"
        tlUrf.sIncompToDead = "Y"
        tlUrf.sIncompToComp = "Y"
        tlUrf.sIncompToHold = "Y"
        tlUrf.sIncompToOrder = "Y"
        tlUrf.sDeadToWork = "Y"
        tlUrf.sHoldToOrder = "Y"
        tlUrf.sReviseCntr = "Y"
        tlUrf.sResvType = "I"
        tlUrf.sRemType = "I"
        tlUrf.sDRType = "I"
        tlUrf.sPIType = "I"
        tlUrf.sPSAType = "I"
        tlUrf.sPromoType = "I"
        tlUrf.sRefResvType = "I"
        tlUrf.sUseComputeCMC = "H"
        tlUrf.sRegionCopy = "I"
        tlUrf.sChgPrices = "I"
        tlUrf.sActFlightButton = "I"
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainIniValues                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Ini Values               *
'*                                                     *
'*******************************************************
Sub gObtainIniValues()
    Dim ilRet As Integer
    Dim ilFound As Integer  'Traffic.Ini found in current directory
    Dim slFileName As String
    Dim slReturn As String * 130
    Dim slLocation As String
    Dim ilMasterDB As Integer
    Dim slStr As String
    Dim slDSN As String
    Dim ilTimeOut As Integer
    Dim slTimeOut As String
    'Dim fs As Object  D.S. 03/18/03
    Dim fs As New FileSystemObject

    sgIniPath = gSetPathEndSlash(sgIniPath, True)

    If igTestSystem Then
        slLocation = "TestLocations"
    Else
        slLocation = "Locations"
    End If

    ilFound = True
    ilMasterDB = False
    On Error GoTo gObtainIniValuesErr
    If igDirectCall = -1 Then
        slFileName = sgIniPath & "Traffic.Ini"
    Else
        slFileName = CurDir$ & "\Traffic.Ini"
    End If
    ilRet = GetPrivateProfileString("Locations", "UsingSockets", "Not Found", slReturn, 128, slFileName)
    If UCase(Trim$(slReturn)) = "Y" Then
        bgUsingSockets = True
    Else
        bgUsingSockets = False
    End If
    'When running in debug mode- the error is not occurring-so test return string
    ilRet = GetPrivateProfileString(slLocation, "Database", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "Traffic.Ini"
        Else
            slFileName = CurDir$ & "\Traffic.Ini"
        End If
        'When running in debug mode- the error is not occurring-so test return string
        ilRet = GetPrivateProfileString(slLocation, "Master Database", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            ilFound = False
        Else
            ilMasterDB = True
        End If
    End If
    If ilFound Then
        If Not ilMasterDB Then
            sgDBPath = Left$(slReturn, ilRet)
            sgMDBPath = ""
            sgSDBPath = ""
            igRetrievalDB = 0
        Else
            sgDBPath = ""
            sgMDBPath = Left$(slReturn, ilRet)
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Slave Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
                'End
                sgSDBPath = ""
            Else
                sgSDBPath = Left$(slReturn, ilRet)
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Retrieval Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                gMsgBox "Traffic.Ini 'Retrieval Database' Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
                End
            End If
            slStr = Left$(slReturn, ilRet)
            If StrComp(slStr, "Slave", 1) = 0 Then
                igRetrievalDB = 1
            Else
                igRetrievalDB = 0
            End If
        End If

        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "ServerDatabase", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgServerDatabase = ""
        Else
            sgServerDatabase = Left$(slReturn, ilRet)
            'If right$(sgServerDatabase, 1) <> "\" Then
            '    sgServerDatabase = sgServerDatabase & "\"
            'End If
            sgServerDatabase = gSetPathEndSlash(sgServerDatabase, True)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Temp Report Files", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgTDBPath = ""
        Else
            sgTDBPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Reports", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Reports' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgRptPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Save Reports", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Save Reports' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgRptSavePath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Import", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Import' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgImportPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Export", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Export' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgExportPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Archive", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgArchivePath = ""
        Else
            sgArchivePath = Left$(slReturn, ilRet)
        End If
        
        
        '1-6-05 add option to indicate separate automation import/export path
        'if not found, default to the general import/export paths
        igProphetImportPathFlag = 0      'flag indicating that the path needs to be changed in the browser to the automation import path loc.

        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "ProphetImport", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgProphetImportPath = sgImportPath
        Else
            sgProphetImportPath = Left$(slReturn, ilRet)
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "ProphetExport", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgProphetExportPath = sgExportPath
        Else
            sgProphetExportPath = Left$(slReturn, ilRet)
        End If
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Logo", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Logo' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgLogoPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Help", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Help' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgHelpPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Exe' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgExePath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Download Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgDLExePath = ""
        Else
            sgDLExePath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Download Rpt", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgDLRptPath = ""
        Else
            sgDLRptPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "DDF", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'DDF' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgDDFPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString("DDELinks", "DelayCount", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'DelayCount' DDELinks Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        igDDEDelay = Val(Left$(slReturn, ilRet))
        On Error Resume Next
        ilRet = GetPrivateProfileString("General", "Using Time Stamp", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            igUsingTimeStamp = True
        Else
            If Left$(slReturn, ilRet) = "No" Then
                igUsingTimeStamp = False
            Else
                igUsingTimeStamp = True
            End If
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString("General", "RemoteUserDirectMode", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            igRUDirectMode = 0
            On Error Resume Next
            If igTestSystem Then
                ilRet = GetPrivateProfileString("TestRemote", "Traffic.Ini", "Not Found", slReturn, 128, slFileName)
            Else
                ilRet = GetPrivateProfileString("Remote", "Traffic.Ini", "Not Found", slReturn, 128, slFileName)
            End If
            If Left$(slReturn, ilRet) = "Not Found" Then
                igRUDirectMode = 2
            Else
                igRUDirectMode = 0
            End If
        Else
            If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
                igRUDirectMode = 1
            Else
                igRUDirectMode = 2
            End If
        End If


    Else
        slFileName = "Traffic.Ini"
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Database", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            slFileName = "Traffic.Ini"
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Master Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                If igDirectCall = -1 Then
                    gMsgBox "Traffic.Ini missing or 'Database' Location Definition Missing from Traffic.Ini in " & sgIniPath, vbCritical + vbOKOnly, "Traffic.Ini Error"
                Else
                    gMsgBox "Traffic.Ini missing or 'Database' Location Definition Missing from Traffic.Ini in " & CurDir$, vbCritical + vbOKOnly, "Traffic.Ini Error"
                End If
                End
            End If
            ilMasterDB = True
        End If
        If Not ilMasterDB Then
            sgDBPath = Left$(slReturn, ilRet)
            sgMDBPath = ""
            sgSDBPath = ""
            igRetrievalDB = 0
        Else
            sgDBPath = ""
            sgMDBPath = Left$(slReturn, ilRet)
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Slave Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
                'End
                sgSDBPath = ""
            Else
                sgSDBPath = Left$(slReturn, ilRet)
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Retrieval Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                gMsgBox "Traffic.Ini 'Retrieval Database' Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
                End
            End If
            slStr = Left$(slReturn, ilRet)
            If StrComp(slStr, "Slave", 1) = 0 Then
                igRetrievalDB = 1
            Else
                igRetrievalDB = 0
            End If
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Temp Report Files", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgTDBPath = ""
        Else
            sgTDBPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Reports", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Reports' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgRptPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Save Reports", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Save Reports' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgRptSavePath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Import", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Import' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgImportPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Export", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Export' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgExportPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Archive", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgArchivePath = ""
        Else
            sgArchivePath = Left$(slReturn, ilRet)
        End If
        
        '1-6-05 add option to indicate separate automation import/export path
        'if not found, default to the general import/export paths
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "ProphetImport", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgProphetImportPath = sgImportPath
        End If
        sgProphetImportPath = Left$(slReturn, ilRet)
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "ProphetExport", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgProphetExportPath = sgExportPath
        End If
        sgProphetExportPath = Left$(slReturn, ilRet)
        
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Logo", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Logo' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgLogoPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Help", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Help' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgHelpPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'Exe' Location Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        sgExePath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Download Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgDLExePath = ""
        Else
            sgDLExePath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Download Rpt", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            'MsgBox "Traffic.Ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "Traffic.Ini Error"
            'End
            sgDLRptPath = ""
        Else
            sgDLRptPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "DDF", "Not Found", slReturn, 128, slFileName)
        sgDDFPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString("DDELinks", "DelayCount", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "Traffic.Ini 'DelayCount' DDELinks Definition Missing", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
        igDDEDelay = Val(Left$(slReturn, ilRet))
        On Error Resume Next
        ilRet = GetPrivateProfileString("General", "Using Time Stamp", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            igUsingTimeStamp = True
        Else
            If (UCase$(Left$(slReturn, ilRet)) = "NO") Or (UCase$(Left$(slReturn, ilRet)) = "FALSE") Then
                igUsingTimeStamp = False
            Else
                igUsingTimeStamp = True
            End If
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString("General", "RemoteUserDirectMode", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            igRUDirectMode = 0
            On Error Resume Next
            If igTestSystem Then
                ilRet = GetPrivateProfileString("TestRemote", "Traffic.Ini", "Not Found", slReturn, 128, slFileName)
            Else
                ilRet = GetPrivateProfileString("Remote", "Traffic.Ini", "Not Found", slReturn, 128, slFileName)
            End If
            If Left$(slReturn, ilRet) = "Not Found" Then
                igRUDirectMode = 2
            Else
                igRUDirectMode = 0
            End If
        Else
            If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
                igRUDirectMode = 1
            Else
                igRUDirectMode = 2
            End If
        End If
    End If
    On Error GoTo 0
    If Not ilMasterDB Then
        'If right$(sgDBPath, 1) <> "\" Then
        '    sgDBPath = sgDBPath & "\"
        'End If
        sgDBPath = gSetPathEndSlash(sgDBPath, False)
    Else
        'If right$(sgMDBPath, 1) <> "\" Then
        '    sgMDBPath = sgMDBPath & "\"
        'End If
        sgMDBPath = gSetPathEndSlash(sgMDBPath, False)
        If sgSDBPath <> "" Then
            'If right$(sgSDBPath, 1) <> "\" Then
            '    sgSDBPath = sgSDBPath & "\"
            'End If
            sgSDBPath = gSetPathEndSlash(sgSDBPath, False)
        End If
        If StrComp(sgMDBPath, sgSDBPath, 1) = 0 Then
            sgSDBPath = ""
        End If
        'Set sgDBPath used for time stamps,
        'it will be replaced from the open calls in csi_io.dll
        'by sg-DBPath
        If igRetrievalDB = 0 Then
            sgDBPath = sgMDBPath
        Else
            sgDBPath = sgSDBPath
        End If
    End If
    If igTestSystem Then
        If InStr(1, sgDBPath, "Test", vbTextCompare) = 0 Then
            gMsgBox "Using Test System but Database = reference in Traffic.Ini [TestLocation] area does not contain the word Test in the Path", vbCritical + vbOKOnly, "Traffic.Ini Error"
            End
        End If
    End If
    If Len(sgTDBPath) = 0 Then
        sgTDBPath = sgDBPath
    End If
    'If right$(sgTDBPath, 1) <> "\" Then
    '    sgTDBPath = sgTDBPath & "\"
    'End If
    'If right$(sgRptPath, 1) <> "\" Then
    '    sgRptPath = sgRptPath & "\"
    'End If
    'If right$(sgRptSavePath, 1) <> "\" Then
    '    sgRptSavePath = sgRptSavePath & "\"
    'End If
    'If right$(sgImportPath, 1) <> "\" Then
    '    sgImportPath = sgImportPath & "\"
    'End If
    'If right$(sgExportPath, 1) <> "\" Then
    '    sgExportPath = sgExportPath & "\"
    'End If
    If sgArchivePath = "" Then
        sgArchivePath = sgDBPath & "Archive\"
    End If
    'If right$(sgArchivePath, 1) <> "\" Then
    '    sgArchivePath = sgArchivePath & "\"
    'End If
    '1-6-05
    'If right$(sgProphetImportPath, 1) <> "\" Then
    '    sgProphetImportPath = sgProphetImportPath & "\"
    'End If
    'If right$(sgProphetExportPath, 1) <> "\" Then
    '    sgProphetExportPath = sgProphetExportPath & "\"
    'End If
    
    'If right$(sgLogoPath, 1) <> "\" Then
    '    sgLogoPath = sgLogoPath & "\"
    'End If
    'If right$(sgHelpPath, 1) <> "\" Then
    '    sgHelpPath = sgHelpPath & "\"
    'End If
    'If right$(sgExePath, 1) <> "\" Then
    '    sgExePath = sgExePath & "\"
    'End If
    'If Len(sgDLExePath) > 0 Then
    '    If right$(sgDLExePath, 1) <> "\" Then
    '        sgDLExePath = sgDLExePath & "\"
    '    End If
    '    If StrComp(sgExePath, sgDLExePath, 1) = 0 Then
    '        sgDLExePath = ""
    '    End If
    'End If
    'If Len(sgDLRptPath) > 0 Then
    '    If right$(sgDLRptPath, 1) <> "\" Then
    '        sgDLRptPath = sgDLRptPath & "\"
    '    End If
    '    If StrComp(sgRptPath, sgDLRptPath, 1) = 0 Then
    '        sgDLRptPath = ""
    '    End If
    'End If
    'If right$(sgDDFPath, 1) <> "\" Then
    '    sgDDFPath = sgDDFPath & "\"
    'End If
    
    sgDBPath = gSetPathEndSlash(sgDBPath, True)
    sgMDBPath = gSetPathEndSlash(sgMDBPath, True)
    sgSDBPath = gSetPathEndSlash(sgSDBPath, True)
    sgTDBPath = gSetPathEndSlash(sgTDBPath, True)
    sgRptPath = gSetPathEndSlash(sgRptPath, True)
    sgRptSavePath = gSetPathEndSlash(sgRptSavePath, True)
    sgImportPath = gSetPathEndSlash(sgImportPath, True)
    sgExportPath = gSetPathEndSlash(sgExportPath, True)
    sgArchivePath = gSetPathEndSlash(sgArchivePath, True)
    sgProphetImportPath = gSetPathEndSlash(sgProphetImportPath, True)
    sgProphetExportPath = gSetPathEndSlash(sgProphetExportPath, True)
    sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    sgHelpPath = gSetPathEndSlash(sgHelpPath, True)
    sgExePath = gSetPathEndSlash(sgExePath, True)
    sgDLExePath = gSetPathEndSlash(sgDLExePath, True)
    sgDLRptPath = gSetPathEndSlash(sgDLRptPath, True)
    sgDDFPath = gSetPathEndSlash(sgDDFPath, True)
    
    If Len(sgDLExePath) > 0 Then
        If StrComp(sgExePath, sgDLExePath, 1) = 0 Then
            sgDLExePath = ""
        End If
    End If
    If Len(sgDLRptPath) > 0 Then
        If StrComp(sgRptPath, sgDLRptPath, 1) = 0 Then
            sgDLRptPath = ""
        End If
    End If
    
    'D.S. 03/18/03 Set fs = CreateObject("scripting.FileSystemObject")
    If Not fs.FolderExists(Left$(sgDBPath, Len(sgDBPath) - 1)) Then
        gMsgBox ("Database = " & sgDBPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgRptPath, Len(sgRptPath) - 1)) Then
        gMsgBox ("Reports = " & sgRptPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgRptSavePath, Len(sgRptSavePath) - 1)) Then
        gMsgBox ("Save Reports = " & sgRptSavePath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgImportPath, Len(sgImportPath) - 1)) Then
        gMsgBox ("Import = " & sgImportPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgExportPath, Len(sgExportPath) - 1)) Then
        gMsgBox ("Export = " & sgExportPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    
    '1-7-05 insure the automation folders exists
    If Not fs.FolderExists(Left$(sgProphetImportPath, Len(sgProphetImportPath) - 1)) Then
        gMsgBox ("ProphetImport = " & sgProphetImportPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgProphetExportPath, Len(sgProphetExportPath) - 1)) Then
        gMsgBox ("ProphetExport = " & sgProphetExportPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgLogoPath, Len(sgLogoPath) - 1)) Then
        gMsgBox ("Logo = " & sgLogoPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgHelpPath, Len(sgHelpPath) - 1)) Then
        gMsgBox ("Help = " & sgHelpPath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    If Not fs.FolderExists(Left$(sgExePath, Len(sgExePath) - 1)) Then
        gMsgBox ("Exe = " & sgExePath & " path within Traffic.Ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    'D.S. 03/18/03
    If Not fs.FolderExists(Left$(sgDBPath & "Messages\", Len(sgDBPath + "Messages\") - 1)) Then
        'MsgBox ("Messages = " & sgDBPath & "Messages\" & " path is invalid.  Counterpoint will attempt to create the directory")
        slFileName = sgDBPath & "Messages"
        fs.CreateFolder (slFileName)
        If Not fs.FolderExists(slFileName) Then
            gMsgBox ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder."), vbOKOnly + vbApplicationModal, "gObtainIniValue"
        End If
    End If

    If Not gLoadOptionTrafficThenAffiliate(slLocation, "Name", slStr) Then
        If igTestSystem Then
            gMsgBox "Traffic.Ini [TestLocations] 'Name' key is missing. Application will now shutdown.", vbCritical
        Else
            gMsgBox "Traffic.Ini [Locations] 'Name' key is missing. Application will now shutdown.", vbCritical
        End If
        End
    End If
    Set cnn = New ADODB.Connection
    slDSN = slStr
    cnn.Open "DSN=" & slDSN
    
    
    '5/8/19: Add timeout setting
    ilTimeOut = 600
    If gLoadOptionTrafficThenAffiliate(slLocation, "TimeOut", slTimeOut) Then
        ilTimeOut = Val(slTimeOut)
    End If
    cnn.CommandTimeout = ilTimeOut

    
    Exit Sub
gObtainIniValuesErr:
    ilFound = False
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRUIniValues              *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Remote User Ini Values   *
'*                                                     *
'*******************************************************
Sub gObtainRUIniValues()
    Dim ilRet As Integer
    Dim slFileName As String
    Dim slReturn As String * 130
    Dim slLocation As String

    If igTestSystem Then
        slLocation = "TestRemote"
    Else
        slLocation = "Remote"
    End If

    slFileName = CurDir$ & "\Traffic.Ini"
    sgRUImportPath = ""
    sgRUExportPath = ""
    sgRUCmmlPath = ""
    lgRUCheckSec = 30000   'Every 30 seconds (1000=1sec)
    sgIniPath = ""
    igRUMode = 0   'Dedicated
    On Error Resume Next
    ilRet = GetPrivateProfileString(slLocation, "Import", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        sgRUImportPath = ""
    Else
        sgRUImportPath = Left$(slReturn, ilRet)
    End If
    On Error Resume Next
    ilRet = GetPrivateProfileString(slLocation, "Export", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        sgRUExportPath = ""
    Else
        sgRUExportPath = Left$(slReturn, ilRet)
    End If
    On Error Resume Next
    ilRet = GetPrivateProfileString(slLocation, "TimeOut", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        lgRUTimeOut = 900
        sgRUTimeOut = "15 Minutes"
    Else
        lgRUTimeOut = 60 * Val(Left$(slReturn, ilRet))
        If lgRUTimeOut <= 0 Then
            lgRUTimeOut = 86400
            sgRUTimeOut = "24 Hours"
        Else
            sgRUTimeOut = Left$(slReturn, ilRet) & " Minutes"
        End If
    End If
    If lgRUTimeOut <= 0 Then
        lgRUTimeOut = 900
        sgRUTimeOut = "15 Minutes"
    End If
    On Error Resume Next
    ilRet = GetPrivateProfileString(slLocation, "XFerWorking", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        igRUXferWorking = False
    Else
        If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
            igRUXferWorking = True
        Else
            igRUXferWorking = False
        End If
    End If
    On Error Resume Next
    ilRet = GetPrivateProfileString(slLocation, "Dial-Up", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        igRUDialUp = False
    Else
        If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
            igRUDialUp = True
        Else
            igRUDialUp = False
        End If
    End If

    If tgUrf(0).iRemoteUserID <= 0 Then
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Commands", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgRUCmmlPath = ""
        Else
            sgRUCmmlPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "CheckSec", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            lgRUCheckSec = 30000   'Every 30 sec
        Else
            If Val(Left$(slReturn, ilRet)) > 60 Then
                lgRUCheckSec = 60000
            Else
                lgRUCheckSec = 1000 * Val(Left$(slReturn, ilRet))
            End If
        End If
    Else
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Traffic.Ini", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            sgIniPath = ""
        Else
            sgIniPath = Left$(slReturn, ilRet)
        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Control", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            igRUMode = 0   'Dedicated
        Else
            If StrComp(Left$(slReturn, ilRet), "Remote", 1) = 0 Then
                igRUMode = 1   'Direct
            ElseIf StrComp(Left$(slReturn, ilRet), "Manual", 1) = 0 Then
                igRUMode = 2   'Manual
            Else
                igRUMode = 0   'Dedicated
            End If
        End If
    End If
    If (right$(sgIniPath, 1) <> "\") And (Len(sgIniPath) > 0) Then
        sgIniPath = sgIniPath & "\"
    End If
    If (right$(sgRUExportPath, 1) <> "\") And (Len(sgRUExportPath) > 0) Then
        sgRUExportPath = sgRUExportPath & "\"
    End If
    If (right$(sgRUImportPath, 1) <> "\") And (Len(sgRUImportPath) > 0) Then
        sgRUImportPath = sgRUImportPath & "\"
    End If
    If (right$(sgRUCmmlPath, 1) <> "\") And (Len(sgRUCmmlPath) > 0) Then
        sgRUCmmlPath = sgRUCmmlPath & "\"
    End If
    Exit Sub
End Sub



Public Sub gSetBrannerName()
    Dim slVersion As String
    Dim ilPos As Integer
    Dim slChar As String

    slVersion = "V"
    ilPos = InStr(sgCSVersion, " ") + 1
    slChar = Mid$(sgCSVersion, ilPos, 1)
    Do While slChar <> " "
        slVersion = slVersion & slChar
        ilPos = ilPos + 1
        If ilPos > Len(sgCSVersion) Then
            Exit Do
        End If
        slChar = Mid$(sgCSVersion, ilPos, 1)
    Loop
    If igTestSystem Then
        sgBrannerMsg = "Counterpoint ® TEST, " & slVersion
    Else
        sgBrannerMsg = "Counterpoint ®, " & slVersion
    End If

End Sub

Public Sub gSetBkgdMode()
    igBkgdProg = 10
End Sub

'***************************************************************************************
'*
'* Procedure Name: gSetPathEndSlash
'*
'* Created: 10/02/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Install the final back slash if it does not already exist.
'*
'***************************************************************************************
Public Function gSetPathEndSlash(ByVal slInPath As String, ilAdjDrivePath As Integer) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If Len(slPath) > 0 Then
        If right$(slPath, 1) <> "\" Then
            slPath = slPath + "\"
        End If
        If ilAdjDrivePath Then
            slPath = gAdjustDrivePath(slPath)
        End If
    End If
    gSetPathEndSlash = slPath
End Function
Public Function gAdjustDrivePath(slInDrivePath As String) As String
    'This routine is used to check if the application is running on the server and the drive must be converted
    Dim slCurDir As String
    Dim slDrivePath As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer

    slDrivePath = Trim$(slInDrivePath)
    gAdjustDrivePath = slDrivePath
    slCurDir = Trim$(sgCurDir)
    If InStr(1, slDrivePath, "//", vbTextCompare) > 0 Then
        Exit Function
    End If
    If InStr(1, slDrivePath, "C:", vbTextCompare) > 0 Then
        Exit Function
    End If
    If InStr(1, slDrivePath, "D:", vbTextCompare) > 0 Then
        Exit Function
    End If
    If (Len(slDrivePath) <= 1) Or (Len(slCurDir) <= 1) Then
        Exit Function
    End If
    If StrComp(Left$(slCurDir, 1), Left(slDrivePath, 1), vbTextCompare) = 0 Then
        Exit Function
    End If
    'Find common parts
    ReDim tlField(0 To 0) As DRIVEPATH
    ReDim tlFieldCD(0 To 0) As DRIVEPATH
    ilPos1 = 1
    Do
        ilPos1 = InStr(ilPos1, slDrivePath, "\", vbTextCompare)
        If ilPos1 > 0 Then
            ilPos2 = InStr(ilPos1 + 1, slDrivePath, "\", vbTextCompare)
            If ilPos2 > 0 Then
                tlField(UBound(tlField)).sFolder = Mid(slDrivePath, ilPos1 + 1, ilPos2 - ilPos1 - 1)
                tlField(UBound(tlField)).iPos = ilPos1
                ReDim Preserve tlField(0 To UBound(tlField) + 1) As DRIVEPATH
            Else
                If ilPos1 < Len(slDrivePath) Then
                    tlField(UBound(tlField)).sFolder = Mid(slDrivePath, ilPos1 + 1)
                    tlField(UBound(tlField)).iPos = ilPos1
                Else
                    ReDim Preserve tlField(0 To UBound(tlField) - 1) As DRIVEPATH
                End If
                Exit Do
            End If
        Else
            Exit Do
        End If
        ilPos1 = ilPos2 - 1
    Loop While ilPos2 > 0
    ilPos1 = 1
    Do
        ilPos1 = InStr(ilPos1, slCurDir, "\", vbTextCompare)
        If ilPos1 > 0 Then
            ilPos2 = InStr(ilPos1 + 1, slCurDir, "\", vbTextCompare)
            If ilPos2 > 0 Then
                tlFieldCD(UBound(tlFieldCD)).sFolder = Mid(slCurDir, ilPos1 + 1, ilPos2 - ilPos1 - 1)
                tlFieldCD(UBound(tlFieldCD)).iPos = ilPos1
                ReDim Preserve tlFieldCD(0 To UBound(tlFieldCD) + 1) As DRIVEPATH
            Else
                If ilPos1 < Len(slCurDir) Then
                    tlFieldCD(UBound(tlFieldCD)).sFolder = Mid(slCurDir, ilPos1 + 1)
                    tlFieldCD(UBound(tlFieldCD)).iPos = ilPos1
                Else
                    ReDim Preserve tlFieldCD(0 To UBound(tlFieldCD) - 1) As DRIVEPATH
                End If
                Exit Do
            End If
        Else
            Exit Do
        End If
        ilPos1 = ilPos2 - 1
    Loop While ilPos2 > 0
    If (UBound(tlField) >= 3) And (UBound(tlFieldCD) >= 3) Then
        If StrComp(tlField(UBound(tlField) - 1).sFolder, tlFieldCD(UBound(tlFieldCD) - 1).sFolder, vbTextCompare) = 0 Then
            gAdjustDrivePath = Left$(slCurDir, tlFieldCD(UBound(tlFieldCD) - 1).iPos) & Mid$(slDrivePath, tlField(UBound(tlField) - 1).iPos + 1)
        Else
            If StrComp(tlField(UBound(tlField) - 2).sFolder, tlFieldCD(UBound(tlFieldCD) - 2).sFolder, vbTextCompare) = 0 Then
                gAdjustDrivePath = Left$(slCurDir, tlFieldCD(UBound(tlFieldCD) - 2).iPos) & Mid$(slDrivePath, tlField(UBound(tlField) - 2).iPos + 1)
            Else
                If StrComp(tlField(UBound(tlField) - 3).sFolder, tlFieldCD(UBound(tlFieldCD) - 3).sFolder, vbTextCompare) = 0 Then
                    gAdjustDrivePath = Left$(slCurDir, tlFieldCD(UBound(tlFieldCD) - 3).iPos) & Mid$(slDrivePath, tlField(UBound(tlField) - 3).iPos + 1)
                End If
            End If
        End If
    End If
End Function
