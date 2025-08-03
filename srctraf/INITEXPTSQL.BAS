Attribute VB_Name = "INITEXPTSQL"
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

    sgCSVersion = "Version " & App.Major & "." & App.Minor
    sgCurDir = CurDir$
    ilNoFiles = SetHandleCount(60)
    igRUDirectMode = 0
    gObtainIniValues
    sgSUName = ""
    sgCPName = ""
    sgPasswordAddition = ""

    For ilLoop = LBound(igWinStatus) To UBound(igWinStatus) Step 1
        igWinStatus(ilLoop) = 0
    Next ilLoop

    sgBS = Chr$(8)      'Backspace
    sgTB = Chr$(9)      'Tab
    sgLF = Chr$(10)     'Line feed (newline)
    sgCR = Chr$(13)     'Carriage return
    sgRUCrLf = Chr$(141) & Chr$(138)

    sgCnCDelimiter = Chr$(9)    'Tab
    sgCnCNameBracket = ""   ' or 4 quotes to make a quote """"
    igSSFBaseLen = 20
    igShowCatOrNames = 0    'Initialize by Categories
    igEtfRecsCreated = False    'Event type records created flag
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
    sgCommAgfStamp = ""
    sgCommAdfStamp = ""
    sgCompMnfStamp = ""

    llVer = GetVersion() And &HFFFF&
    flVerMajor = llVer Mod 256
    flVerMinor = (llVer \ 256) / 100
    flWinVersion = flVerMajor + flVerMinor
    fgWinVersion = flWinVersion

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


    ilTraffic1 = 10
    ilMultiNm = 50

    ReDim tgUrf(0 To 0) As URF

    'csiInitName

    'hgDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
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
    Exit Sub
                  
    ilRet = Err.Number
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
        If (Trim$(tlUrf.sName) = sgCPName) Then
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
        End If
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
    Dim ilFound As Integer  'ExportSQL.ini found in current directory
    Dim slFileName As String
    Dim slReturn As String * 130
    Dim slLocation As String
    Dim ilMasterDB As Integer
    Dim slStr As String
    Dim slDSN As String
    'Dim fs As Object  D.S. 03/18/03
    Dim fs As New FileSystemObject


    If igTestSystem Then
        slLocation = "TestLocations"
    Else
        slLocation = "Locations"
    End If

    igUsingTimeStamp = True

    ilFound = True
    ilMasterDB = False
    On Error GoTo gObtainIniValuesErr
    If igDirectCall = -1 Then
        slFileName = sgIniPath & "ExportSQL.ini"
    Else
        slFileName = CurDir$ & "\ExportSQL.ini"
    End If
    'When running in debug mode- the error is not occurring-so test return string
    ilRet = GetPrivateProfileString(slLocation, "Database", "Not Found", slReturn, 128, slFileName)
    If Left$(slReturn, ilRet) = "Not Found" Then
        If igDirectCall = -1 Then
            slFileName = sgIniPath & "ExportSQL.ini"
        Else
            slFileName = CurDir$ & "\ExportSQL.ini"
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
'        Else
'            sgDBPath = ""
'            sgMDBPath = Left$(slReturn, ilRet)
'            On Error Resume Next
'            ilRet = GetPrivateProfileString(slLocation, "Slave Database", "Not Found", slReturn, 128, slFileName)
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'                'End
'                sgSDBPath = ""
'            Else
'                sgSDBPath = Left$(slReturn, ilRet)
'            End If
'            On Error Resume Next
'            ilRet = GetPrivateProfileString(slLocation, "Retrieval Database", "Not Found", slReturn, 128, slFileName)
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                gMsgBox "ExportSQL.ini 'Retrieval Database' Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'                End
'            End If
'            slStr = Left$(slReturn, ilRet)
'            If StrComp(slStr, "Slave", 1) = 0 Then
'                igRetrievalDB = 1
'            Else
'                igRetrievalDB = 0
'            End If
        End If
        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "ServerDatabase", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            sgServerDatabase = ""
'        Else
'            sgServerDatabase = Left$(slReturn, ilRet)
'            If right$(sgServerDatabase, 1) <> "\" Then
'                sgServerDatabase = sgServerDatabase & "\"
'            End If
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Temp Report Files", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgTDBPath = ""
'        Else
'            sgTDBPath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Reports", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgRptPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Save Reports", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Save Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgRptSavePath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Import", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Import' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgImportPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Export", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Export' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgExportPath = Left$(slReturn, ilRet)
'
'        '1-6-05 add option to indicate separate automation import/export path
'        'if not found, default to the general import/export paths
'        igProphetImportPathFlag = 0      'flag indicating that the path needs to be changed in the browser to the automation import path loc.
'
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "ProphetImport", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            sgProphetImportPath = sgImportPath
'        Else
'            sgProphetImportPath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "ProphetExport", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            sgProphetExportPath = sgExportPath
'        Else
'            sgProphetExportPath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Logo", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Logo' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgLogoPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Help", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Help' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgHelpPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "ExportSQL.ini 'Exe' Location Definition Missing", vbCritical + vbOKOnly, "ExportSQL.ini Error"
            End
        End If
        sgExePath = Left$(slReturn, ilRet)
        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Download Exe", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgDLExePath = ""
'        Else
'            sgDLExePath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Download Rpt", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgDLRptPath = ""
'        Else
'            sgDLRptPath = Left$(slReturn, ilRet)
'        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "DDF", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "ExportSQL.ini 'DDF' Location Definition Missing", vbCritical + vbOKOnly, "ExportSQL.ini Error"
            End
        End If
        sgDDFPath = Left$(slReturn, ilRet)
        On Error Resume Next
'        ilRet = GetPrivateProfileString("DDELinks", "DelayCount", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'DelayCount' DDELinks Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        igDDEDelay = Val(Left$(slReturn, ilRet))
'        On Error Resume Next
'        ilRet = GetPrivateProfileString("General", "Using Time Stamp", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            igUsingTimeStamp = True
'        Else
'            If Left$(slReturn, ilRet) = "No" Then
'                igUsingTimeStamp = False
'            Else
'                igUsingTimeStamp = True
'            End If
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString("General", "RemoteUserDirectMode", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            igRUDirectMode = 0
'            On Error Resume Next
'            If igTestSystem Then
'                ilRet = GetPrivateProfileString("TestRemote", "ExportSQL.ini", "Not Found", slReturn, 128, slFileName)
'            Else
'                ilRet = GetPrivateProfileString("Remote", "ExportSQL.ini", "Not Found", slReturn, 128, slFileName)
'            End If
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                igRUDirectMode = 2
'            Else
'                igRUDirectMode = 0
'            End If
'        Else
'            If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
'                igRUDirectMode = 1
'            Else
'                igRUDirectMode = 2
'            End If
'        End If
    Else
        slFileName = "ExportSQL.ini"
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Database", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            slFileName = "ExportSQL.ini"
            On Error Resume Next
            ilRet = GetPrivateProfileString(slLocation, "Master Database", "Not Found", slReturn, 128, slFileName)
            If Left$(slReturn, ilRet) = "Not Found" Then
                If igDirectCall = -1 Then
                    gMsgBox "ExportSQL.ini missing or 'Database' Location Definition Missing from ExportSQL.ini in " & sgIniPath, vbCritical + vbOKOnly, "ExportSQL.ini Error"
                Else
                    gMsgBox "ExportSQL.ini missing or 'Database' Location Definition Missing from ExportSQL.ini in " & CurDir$, vbCritical + vbOKOnly, "ExportSQL.ini Error"
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
'        Else
'            sgDBPath = ""
'            sgMDBPath = Left$(slReturn, ilRet)
'            On Error Resume Next
'            ilRet = GetPrivateProfileString(slLocation, "Slave Database", "Not Found", slReturn, 128, slFileName)
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'                'End
'                sgSDBPath = ""
'            Else
'                sgSDBPath = Left$(slReturn, ilRet)
'            End If
'            On Error Resume Next
'            ilRet = GetPrivateProfileString(slLocation, "Retrieval Database", "Not Found", slReturn, 128, slFileName)
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                gMsgBox "ExportSQL.ini 'Retrieval Database' Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'                End
'            End If
'            slStr = Left$(slReturn, ilRet)
'            If StrComp(slStr, "Slave", 1) = 0 Then
'                igRetrievalDB = 1
'            Else
'                igRetrievalDB = 0
'            End If
        End If
        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Temp Report Files", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgTDBPath = ""
'        Else
'            sgTDBPath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Reports", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgRptPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Save Reports", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Save Reports' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgRptSavePath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Import", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Import' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgImportPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Export", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Export' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgExportPath = Left$(slReturn, ilRet)
'
'        '1-6-05 add option to indicate separate automation import/export path
'        'if not found, default to the general import/export paths
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "ProphetImport", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            sgProphetImportPath = sgImportPath
'        End If
'        sgProphetImportPath = Left$(slReturn, ilRet)
'
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "ProphetExport", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            sgProphetExportPath = sgExportPath
'        End If
'        sgProphetExportPath = Left$(slReturn, ilRet)
'
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Logo", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Logo' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgLogoPath = Left$(slReturn, ilRet)
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Help", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'Help' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        sgHelpPath = Left$(slReturn, ilRet)
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "Exe", "Not Found", slReturn, 128, slFileName)
        If Left$(slReturn, ilRet) = "Not Found" Then
            gMsgBox "ExportSQL.ini 'Exe' Location Definition Missing", vbCritical + vbOKOnly, "ExportSQL.ini Error"
            End
        End If
        sgExePath = Left$(slReturn, ilRet)
        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Download Exe", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgDLExePath = ""
'        Else
'            sgDLExePath = Left$(slReturn, ilRet)
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString(slLocation, "Download Rpt", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            'MsgBox "ExportSQL.ini 'Slave Database' Location Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            'End
'            sgDLRptPath = ""
'        Else
'            sgDLRptPath = Left$(slReturn, ilRet)
'        End If
        On Error Resume Next
        ilRet = GetPrivateProfileString(slLocation, "DDF", "Not Found", slReturn, 128, slFileName)
        sgDDFPath = Left$(slReturn, ilRet)
        On Error Resume Next
'        ilRet = GetPrivateProfileString("DDELinks", "DelayCount", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            gMsgBox "ExportSQL.ini 'DelayCount' DDELinks Definition Missing", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'        igDDEDelay = Val(Left$(slReturn, ilRet))
'        On Error Resume Next
'        ilRet = GetPrivateProfileString("General", "Using Time Stamp", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            igUsingTimeStamp = True
'        Else
'            If (UCase$(Left$(slReturn, ilRet)) = "NO") Or (UCase$(Left$(slReturn, ilRet)) = "FALSE") Then
'                igUsingTimeStamp = False
'            Else
'                igUsingTimeStamp = True
'            End If
'        End If
'        On Error Resume Next
'        ilRet = GetPrivateProfileString("General", "RemoteUserDirectMode", "Not Found", slReturn, 128, slFileName)
'        If Left$(slReturn, ilRet) = "Not Found" Then
'            igRUDirectMode = 0
'            On Error Resume Next
'            If igTestSystem Then
'                ilRet = GetPrivateProfileString("TestRemote", "ExportSQL.ini", "Not Found", slReturn, 128, slFileName)
'            Else
'                ilRet = GetPrivateProfileString("Remote", "ExportSQL.ini", "Not Found", slReturn, 128, slFileName)
'            End If
'            If Left$(slReturn, ilRet) = "Not Found" Then
'                igRUDirectMode = 2
'            Else
'                igRUDirectMode = 0
'            End If
'        Else
'            If StrComp(Left$(slReturn, ilRet), "Yes", 1) = 0 Then
'                igRUDirectMode = 1
'            Else
'                igRUDirectMode = 2
'            End If
'        End If
    End If
    On Error GoTo 0
    If Not ilMasterDB Then
        If right$(sgDBPath, 1) <> "\" Then
            sgDBPath = sgDBPath & "\"
        End If
    Else
        If right$(sgMDBPath, 1) <> "\" Then
            sgMDBPath = sgMDBPath & "\"
        End If
        If sgSDBPath <> "" Then
            If right$(sgSDBPath, 1) <> "\" Then
                sgSDBPath = sgSDBPath & "\"
            End If
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
'    If igTestSystem Then
'        If InStr(1, sgDBPath, "Test", vbTextCompare) = 0 Then
'            gMsgBox "Using Test System but Database = reference in ExportSQL.ini [TestLocation] area does not contain the word Test in the Path", vbCritical + vbOkOnly, "ExportSQL.ini Error"
'            End
'        End If
'    End If

    sgTDBPath = ""
    If Len(sgTDBPath) = 0 Then
        sgTDBPath = sgDBPath
    End If
    If right$(sgTDBPath, 1) <> "\" Then
        sgTDBPath = sgTDBPath & "\"
    End If
    If right$(sgRptPath, 1) <> "\" Then
        sgRptPath = sgRptPath & "\"
    End If
'    If right$(sgRptSavePath, 1) <> "\" Then
'        sgRptSavePath = sgRptSavePath & "\"
'    End If
'    If right$(sgImportPath, 1) <> "\" Then
'        sgImportPath = sgImportPath & "\"
'    End If
'    If right$(sgExportPath, 1) <> "\" Then
'        sgExportPath = sgExportPath & "\"
'    End If
'
'    '1-6-05
'    If right$(sgProphetImportPath, 1) <> "\" Then
'        sgProphetImportPath = sgProphetImportPath & "\"
'    End If
'    If right$(sgProphetExportPath, 1) <> "\" Then
'        sgProphetExportPath = sgProphetExportPath & "\"
'    End If
'
'    If right$(sgLogoPath, 1) <> "\" Then
'        sgLogoPath = sgLogoPath & "\"
'    End If
'    If right$(sgHelpPath, 1) <> "\" Then
'        sgHelpPath = sgHelpPath & "\"
'    End If
    If right$(sgExePath, 1) <> "\" Then
        sgExePath = sgExePath & "\"
    End If
'    If Len(sgDLExePath) > 0 Then
'        If right$(sgDLExePath, 1) <> "\" Then
'            sgDLExePath = sgDLExePath & "\"
'        End If
'        If StrComp(sgExePath, sgDLExePath, 1) = 0 Then
'            sgDLExePath = ""
'        End If
'    End If
'    If Len(sgDLRptPath) > 0 Then
'        If right$(sgDLRptPath, 1) <> "\" Then
'            sgDLRptPath = sgDLRptPath & "\"
'        End If
'        If StrComp(sgRptPath, sgDLRptPath, 1) = 0 Then
'            sgDLRptPath = ""
'        End If
'    End If
    If right$(sgDDFPath, 1) <> "\" Then
        sgDDFPath = sgDDFPath & "\"
    End If
    'D.S. 03/18/03 Set fs = CreateObject("scripting.FileSystemObject")
    If Not fs.FolderExists(Left$(sgDBPath, Len(sgDBPath) - 1)) Then
        gMsgBox ("Database = " & sgDBPath & " path within ExportSQL.ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
'    If Not fs.FolderExists(Left$(sgRptPath, Len(sgRptPath) - 1)) Then
'        gMsgBox ("Reports = " & sgRptPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgRptSavePath, Len(sgRptSavePath) - 1)) Then
'        gMsgBox ("Save Reports = " & sgRptSavePath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgImportPath, Len(sgImportPath) - 1)) Then
'        gMsgBox ("Import = " & sgImportPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgExportPath, Len(sgExportPath) - 1)) Then
'        gMsgBox ("Export = " & sgExportPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    '1-7-05 insure the automation folders exists
'    If Not fs.FolderExists(Left$(sgProphetImportPath, Len(sgProphetImportPath) - 1)) Then
'        gMsgBox ("ProphetImport = " & sgProphetImportPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgProphetExportPath, Len(sgProphetExportPath) - 1)) Then
'        gMsgBox ("ProphetExport = " & sgProphetExportPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgLogoPath, Len(sgLogoPath) - 1)) Then
'        gMsgBox ("Logo = " & sgLogoPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
'    If Not fs.FolderExists(Left$(sgHelpPath, Len(sgHelpPath) - 1)) Then
'        gMsgBox ("Help = " & sgHelpPath & " path within ExportSQL.ini is invalid, please correct"), vbOkOnly + vbApplicationModal, "gObtainIniValue"
'    End If
    If Not fs.FolderExists(Left$(sgExePath, Len(sgExePath) - 1)) Then
        gMsgBox ("Exe = " & sgExePath & " path within ExportSQL.ini is invalid, please correct"), vbOKOnly + vbApplicationModal, "gObtainIniValue"
    End If
    
    If Not gLoadOptionTrafficThenAffiliate(slLocation, "Name", slStr) Then
        gMsgBox "Traffic.Ini [Locations] 'Name' key is missing. Application will now shutdown.", vbCritical
        End
    End If
    Set cnn = New ADODB.Connection
    slDSN = slStr
    cnn.Open "DSN=" & slDSN
    
    'D.S. 03/18/03
'    slFileName = sgDBPath & "Messages"
'    If Not fs.FolderExists(Left$(slFileName, Len(slFileName) - 1)) Then
'        'MsgBox ("Messages = " & sgDBPath & "Messages\" & " path is invalid.  Counterpoint will attempt to create the directory")
'        slFileName = sgDBPath & "Messages"
'        fs.CreateFolder (slFileName)
'        If Not fs.FolderExists(slFileName) Then
'            lMsgBox ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder."), vbOkOnly + vbApplicationModal, "gObtainIniValue"
             'gLogMsg ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder."), "ExportSQL.txt", False
'        End If
'    End If
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

    slFileName = CurDir$ & "\ExportSQL.ini"
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
        ilRet = GetPrivateProfileString(slLocation, "ExportSQL.ini", "Not Found", slReturn, 128, slFileName)
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






Public Sub lLogMsg(sMsg As String, sFileName As String, iKill As Integer)

    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'iKill = True then delete the file first, iKill = False then append to the file
    
    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slToFile As String
    
    slToFile = "D:\CSI\" & sFileName
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
    slFullMsg = Format$(Now, "mm/dd/yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " " & sMsg
    Print #hlLogFile, slFullMsg
    Close hlLogFile
    Exit Sub
    
'Error:
'    ilRet = 1
'    Resume Next
    
End Sub


