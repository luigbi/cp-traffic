Attribute VB_Name = "RPTCR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcr.bas on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSsfSrchKey                  tmGhfSrchKey0                 tmGsfSrchKey0             *
'*  tmGsfSrchKey4                 tmVpfSrchKey                  imLstRecLen               *
'*  tmLdfList                                                                             *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCRGet.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report Get Data for Crystal screen code
Option Explicit
Option Compare Text
'Spot Projection
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering

Dim hmAdf As Integer            'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAdf As ADF

Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1            'SDF record image (key 1)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim tmVefSrchKey As INTKEY0     'VEF key 0 image

Dim imAVefCode() As Integer
Dim hmVLF As Integer            'Vehicle Link file handle
Dim tmVlf As VLF                'VLF record image
Dim tmVlfSrchKey0 As VLFKEY0            'VLF by selling vehicle record image
Dim imVlfRecLen As Integer        'VLF record length
  
Type VLFSORT
    sKey As String * 5
    lAvailTime As Long
    tVlf As VLF
End Type
Dim tmVlfSort() As VLFSORT
Dim tmVlfSortMF() As VLFSORT
Dim tmVlfSortSa() As VLFSORT
Dim tmVlfSortSu() As VLFSORT

Dim hmRcf As Integer            'Rate Card file handle

Dim tmRif As RIF
Dim hmRif As Integer            'Rate Card items file handle
Dim imRifRecLen As Integer      'RIF record length

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmZeroGrf As GRF              'initialized Generic recd

Dim tmCbf As CBF
Dim hmCbf As Integer
Dim imCbfRecLen As Integer        'CBF record length

Dim tmBvf As BVF                  'Budgets by office & vehicle
Dim hmBvf As Integer
Dim imBvfRecLen As Integer        'BVF record length

'Copy Report
Dim hmCpr As Integer            'Copy Report file handle
Dim tmCpr() As CPR                'CPR record image
Dim tmTCpr As CPR
Dim imCprRecLen As Integer        'CPR record length

'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length

' Copy Combo Inventory File
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length

' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer     'TZF record length

'  Media code File
Dim hmMcf As Integer        'Media file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length

'12-27-04
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0
Dim tmSSFSrchKey3 As SSFKEY3
Dim tmAvail As AVAILSS

Dim hmCrf As Integer        'Copy Rotation file handle
Dim tmCrf As CRF
Dim imCrfRecLen As Integer

Dim hmCaf As Integer        'Copy Rotation by Game or Team file handle
Dim tmCaf As CAF
Dim imCafRecLen As Integer
Dim tmCafSrchKey1 As LONGKEY0 'CAF key record image

Dim hmRsf As Integer        'Copy replacement file handle
Dim tmRsf As RSF
Dim imRsfRecLen As Integer
Dim tmRsfSrchKey1 As LONGKEY0 'RSF key record image

Dim hmCvf As Integer
Dim tmCvf As CVF
Dim imCvfRecLen As Integer
Dim tmCvfSrchKey0 As LONGKEY0

Dim hmTxr As Integer
Dim tmTxr As TXR
Dim imTxrRecLen As Integer
Dim tmTxrSrchKey0 As TXRKEY0

Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmCombineGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim tmCombineGsf As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

Type PLAYLIST
    sType As String * 1 'Vehicle type
    iVefCode As Integer 'Conventional or Selling Vehicle
    iLogCode As Integer 'Log Vehicle
    iSAFirstIndex As Integer 'Selected airing codes for selling vehicle
End Type
Type SATABLE
    iAirCode As Integer
    iSellCode As Integer
    iStartDate(0 To 1) As Integer
    iNextIndex As Integer
End Type
Dim tmSATable() As SATABLE
Type BUDGETWKS  'array of budget start/end weeks for a qtr
    iMnfCode As Integer      'budget name code
    lbvfCode As Integer     'budget code
    iYear As Integer        'year of budget record
    'iStartWks(1 To 13) As Integer   'array of start weeks for week, month or qtr
    iStartWks(0 To 13) As Integer   'array of start weeks for week, month or qtr. Index zero ignored
    'iEndWks(1 To 13) As Integer     'array of end weeks for week, month or qtr
    iEndWks(0 To 13) As Integer     'array of end weeks for week, month or qtr. Index zero ignored
    iIndex As Integer       '1 = base year, 2 = compare budget #1, 3 = compare bdg #2, etc.
End Type

Dim tmLogGen() As LOGGEN
Dim tmVpf As VPF                'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle

Dim hmOdf As Integer        'One day file
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdf As ODF            'ODF record image

Dim hmLst As Integer        'affiliate log file

Dim hmCnf As Integer        'Copy Instr file
Dim imCnfRecLen As Integer  'Copy Instr record length
Dim tmCnf As CNF
Dim tmCnfSrchKey As CNFKEY0

Dim hmClf As Integer        'Line file
Dim tmClf As CLF
Dim imClfRecLen As Integer  'Line record length
Dim tmClfSrchKey1 As CLFKEY1

Dim hmCHF As Integer        'Cnt Header file
Dim tmChf As CHF
Dim imCHFRecLen As Integer  'Line record length

Dim hmBof As Integer        'Blackout file
Dim tmBof As BOF
Dim imBofRecLen As Integer  ' record length

Dim hmLlf As Integer        'live log activity header file
Dim tmLlf As LLF
Dim imLlfRecLen As Integer  ' record length
Dim tmLlfList() As LLF

Dim hmLdf As Integer        'live log activity detail file
Dim tmLdf As LDF
Dim imLdfRecLen As Integer  ' record length
Dim tmLDFSrchKey1 As LONGKEY0

Dim hmLcf As Integer        'Library calendar file
Dim tmLcf As LCF
Dim imLcfRecLen As Integer  ' record length
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim tmLcfSrchKey2 As LCFKEY2

Dim hmLef As Integer        'library event file
Dim tmLef As LEF
Dim imLefRecLen As Integer  ' record length
Dim tmLefSrchKey0 As LEFKEY0

Dim hmLvf As Integer        'library version file
Dim tmLvf As LVF
Dim imLvfRecLen As Integer  ' record length
Dim tmLvfSrchKey As LONGKEY0     'LVF key record image

Dim hmLtf As Integer        'library title file
Dim tmLtf As LTF
Dim imLtfRecLen As Integer  ' record length
Dim tmLtfSrchKey As INTKEY0

Dim hmScr As Integer        'Date: 10/10/2018   SCR file
Dim tmScr As SCR
Dim imScrRecLen As Integer
Dim tmScrSrchKey As INTKEY0

Dim hmPrf As Integer        'Product file
Dim tmPrf As PRF
Dim imPrfRecLen As Integer  ' record length

Dim hmSif As Integer        'short title file
Dim tmSif As SIF
Dim imSifRecLen As Integer  ' record length

Dim hmSef As Integer        'Split entry
Dim tmSef As SEF
Dim imSefRecLen As Integer  ' record length
Dim tmSefSrchKey1 As SEFKEY1

Dim hmMsg As Integer        'message handle
Dim tmLibVersions() As LIBVERSIONS
'array of unique library versions for one vehicle with their earliest and latest dates used
Type LIBVERSIONS
    iVefCode As Integer
    lLvfCode As Long            '4-16-14 was integer
    iLvfVersion As Integer      '4-47-14
    lEarliestDate As Long
    lLatestDate As Long
End Type

Dim tmSellAirList() As SELLAIRLIST
'need to keep all vehicles within an advertiser together to show on the Playlist by ISCI (option)
Type SELLAIRLIST
    iAdfCode As Integer         'advt code
    sType As String * 1
    iVefCode As Integer         'selling, conventional, game vehicle code
    iVefAirCode As Integer      'airing vehicle code
End Type

Dim tmAiringInv() As AIRING_INV
Type AIRING_INV
    iVefAirCode As Integer
    lAirMFInv As Long
    lAirSatInv As Long
    lAirSunInv As Long
    lSellMFInv As Long
    lSellSatInv As Long
    lSellSunInv As Long
End Type

Dim tmSSFWeek() As SSFWEEK
Type SSFWEEK
    iVefCode As Integer
    lSSFCode As Long
    iDate(0 To 1) As Integer
    iDayOfWeek As Integer
End Type

Dim imUseAnfCodes() As Integer          'avail name codes to include in Airing Inventory
Dim imInclAnfCodes As Integer           'include or exclude anf codes

'5/22/15: Handle generic copy assigned to airing vehicle
Private Type RSFSORT
    iRotNo As Integer
    tRsf As RSF
End Type

'Date: 10/4/2018 Library days/times for each vehicle FYM
Type LibDaysTimes
    iVefCode As Integer
    lLvfCode As Long
    iLibVersion As Integer          'library version (lvf)
    iStartTime(0 To 1) As Integer   'start time of program (lcf)
    iLen(0 To 1) As Integer          'length of the program (lvf)
    iDays(0 To 6) As Integer        'valid days of program
    bProcessed As Boolean           'processed indicator (true/false)
End Type
Dim tmLibDaysTimes() As LibDaysTimes

Type EVENTCOMPLETE
    iVefCode As Integer
    lAirDate As Long
    iGameNo As Integer
    sLogComplete As String * 1      'logs completed and posted (Y/N)
End Type

Dim slDTStrings() As String         'processed DAY / TIME array
Dim slMonday() As String
Dim slTuesday() As String
Dim slWednesday() As String
Dim slThursday() As String
Dim slFriday() As String
Dim slSaturday() As String
Dim slSunday() As String

'TTP 10791 - Copy Rotations by Advertiser report: add special export option
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmExport As Integer
Dim smExportStatus As String

Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfTable() As ANF
Dim tmAnfSrchKey As INTKEY0     'ANF record image
Dim imAnfRecLen As Integer      'ANF record length
Dim tmChfSrchKey As LONGKEY0            'CHF record image


'       Prepass to find expired or active library versions used by vehicle
'       for a specified date span
'
Public Sub gCreatePgmLibrary()
    Dim ilRet As Integer
    Dim ilLoopOnVehicle As Integer
    Dim llStartDate As Long
    Dim slStartDate As String
    Dim ilStartDate(0 To 1) As Integer
    Dim llEndDate As Long
    Dim slEndDate As String
    Dim ilEndDate(0 To 1) As Integer
    Dim llLoopOnVer As Long
    Dim ilExpired As Integer        'true if expired, false if active
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim llLogDate As Long
    Dim illoop As Integer
    Dim ilFoundVer As Integer
    Dim ilUpperVer As Integer
    Dim ilProcess As Integer
    Dim ilNotTFN As Integer
    Dim ilfirstTime As Integer
    Dim llDate As Long
    Dim llStartOfSearch As Long
    Dim ilStartOfSearch(0 To 1) As Integer
    Dim lStartTime As Long
    Dim iLen0 As Integer
    Dim iLen1 As Integer
    
    Dim llCounter As Long       'Date: 1/6/2020 converted to long; getting overflow error using integer
    Dim ilDayOfWeek As Integer
    Dim slDayOfWeek As String
    Dim slStartTimeA As String      'AM/PM
    Dim slTimePlusLength As String
    Dim slFinalDTString As String
    Dim slCurDay As String
    Dim slLibLen As String
    Dim ilDay As Integer
    
    Dim tlLcfSrchKey As LCFKEY0
    Dim hlLcf As Integer        'Library calendar file
    Dim tlLcf As LCF
    Dim ilLcfRecLen As Integer  ' record length
    Dim slStartTimeL As String 'TTP 9944

    hmLvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: LvF.BTR)", RptSel
    On Error GoTo 0

    imLcfRecLen = Len(tmLcf)    'Save Library calendar
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: LCF.BTR)", RptSel
    On Error GoTo 0

    imLefRecLen = Len(tmLef)    'Save Library calendar
    hmLef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: LEF.BTR)", RptSel
    On Error GoTo 0
    
    imGrfRecLen = Len(tmGrf)    'Save prepass GRF record length
    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: GRF.BTR)", RptSel
    On Error GoTo 0

    imLtfRecLen = Len(tmLtf)    'Save library title record length
    hmLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: LTF.BTR)", RptSel
    On Error GoTo 0
    
    imLtfRecLen = Len(tmScr)    'Save library days and times record length
    hmScr = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmScr, "", sgDBPath & "Scr.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePgmLibraryErr
    gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrOpen: SCR.BTR)", RptSel
    On Error GoTo 0

    ilExpired = True
    If RptSel!rbcSelC4(0).Value = True Then         'active
        ilExpired = False
    End If
    
'    slStartDate = RptSel!edcSelCFrom.Text      'determine earliest/latest dates to retrieve
    '12-11-19 change to csi calendar control
    slStartDate = RptSel!CSI_CalFrom.Text      'determine earliest/latest dates to retrieve

    If Trim$(slStartDate) = "" Then
        slStartDate = "1/1/1970"
    End If
    llStartDate = gDateValue(slStartDate)
    gPackDateLong llStartDate, ilStartDate(0), ilStartDate(1)
'    slEndDate = RptSel!edcSelCFrom1.Text
    slEndDate = RptSel!CSI_CalTo.Text
    If Trim(slEndDate) = "" Then
        slEndDate = "12/31/2069"
    End If
    llEndDate = gDateValue(slEndDate)
    gPackDateLong llEndDate, ilEndDate(0), ilEndDate(1)
    
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    'assume active retrieval
    ilStartOfSearch(0) = ilStartDate(0)
    ilStartOfSearch(1) = ilStartDate(1)
    If ilExpired Then
        llStartOfSearch = gDateValue("1/1/1970")
        gPackDateLong llStartOfSearch, ilStartOfSearch(0), ilStartOfSearch(1)
    Else
        llStartOfSearch = llStartDate
    End If
       
    For ilLoopOnVehicle = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
        If RptSel!lbcSelection(0).Selected(ilLoopOnVehicle) Then
            slNameCode = tgVehicle(ilLoopOnVehicle).sKey    'Traffic!lbcVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            ilVefCode = Val(slCode)
            
            ReDim tmLibVersions(0 To 0) As LIBVERSIONS
            ReDim tmLibDaysTimes(0 To 0) As LibDaysTimes        'Date: 10/5/2018    array for unique library version, start time, length, day   FYM
            
            ilUpperVer = 0
            tmLcfSrchKey2.iVefCode = ilVefCode
            tmLcfSrchKey2.iLogDate(0) = 0   'ilStartDate(0)
            tmLcfSrchKey2.iLogDate(1) = 0   'ilStartDate(1)
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
            ilfirstTime = True
            Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode)
                If tmLcf.iLogDate(0) <= 7 And tmLcf.iLogDate(1) = 0 Then     'm-s tfn records
                    llLogDate = gDateValue("12/31/2069")
                    'llDate = llStartDate
                    llDate = llLogDate
                Else
                    gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llLogDate
                    llDate = llLogDate
                End If
                
' Date: 6/25/2019   FYM
' Commented out after comparing with version 7.0; no need for this test
'
'                ilProcess = False
'                If llDate >= llStartDate And llDate <= llEndDate Then
'                    ilProcess = True
'                Else
'                    'check if TFN exists
'                    If (llDate >= llEndDate) Then
'                        For ilDay = 1 To 7 Step 1
'                            tlLcfSrchKey.iType = 0
'                            tlLcfSrchKey.sStatus = "C"
'                            tlLcfSrchKey.iVefCode = ilVefCode
'                            tlLcfSrchKey.iLogDate(0) = ilDay  '1=Monday; 2= Tuesday;...
'                            tlLcfSrchKey.iLogDate(1) = 0
'                            tlLcfSrchKey.iSeqNo = 1
'                            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'                            If (ilRet <> BTRV_ERR_NONE) Then 'And (tlLcf.sStatus = "C") And (tlLcf.iVefCode = ilVefCode) And (tlLcf.iType = 0) Then
'                                ilProcess = True
'                                Exit For
'                            End If
'                        Next ilDay
'                    End If
'                End If
                
                'ilFoundVer = False
                If tmLcf.sStatus = "C" Then 'And ilProcess Then             'current vs pending
                    For illoop = 0 To 49                '10-2-18 0 based array
                        If tmLcf.lLvfCode(illoop) > 0 Then
                            tmLvfSrchKey.lCode = tmLcf.lLvfCode(illoop)
                            ilRet = btrGetEqual(hmLvf, tmLvf, Len(tmLvf), tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilRet <> BTRV_ERR_NONE Then
                                tmLvf.iVersion = -1
                                ilFoundVer = True
                            End If
                            
                            If tmLcf.iLogDate(0) <= 7 Then
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iDays(0) = tmLcf.iLogDate(0)
                            Else
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iDays(0) = tmLcf.iLogDate(0)
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iDays(1) = tmLcf.iLogDate(1)
                            End If
                            
                            'Date: 5/7/2019 set process indicator to false
                            'tmLibDaysTimes(UBound(tmLibDaysTimes)).bProcessed = False
                            
                            'Date: 10/5/2018    find the matching tmLibDaysTimes entry and set the day of week
                            If UBound(tmLibDaysTimes) = 0 Then          'first time for every Library Version
                                'Date: 10/5/2018    adding new array entry for the unique library version, start time, length, day
                                tmLibDaysTimes(0).iVefCode = tmLcf.iVefCode
                                tmLibDaysTimes(0).lLvfCode = tmLcf.lLvfCode(illoop)
                                tmLibDaysTimes(0).iLibVersion = tmLvf.iVersion
                                'set length (0,1)
                                tmLibDaysTimes(0).iLen(0) = tmLvf.iLen(0)
                                tmLibDaysTimes(0).iLen(1) = tmLvf.iLen(1)
                                'set StartTime for log
                                tmLibDaysTimes(0).iStartTime(0) = tmLcf.iTime(0, illoop)
                                tmLibDaysTimes(0).iStartTime(1) = tmLcf.iTime(1, illoop)
                                'set date
                                tmLibDaysTimes(0).iDays(0) = tmLcf.iLogDate(0)
                                ReDim Preserve tmLibDaysTimes(0 To UBound(tmLibDaysTimes) + 1) As LibDaysTimes
                            Else
                                'Date: 10/5/2018    adding new array entry for the unique library version, start time, length, day
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iVefCode = tmLcf.iVefCode
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).lLvfCode = tmLcf.lLvfCode(illoop)
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iLibVersion = tmLvf.iVersion     'tmLcf.lLvfCode(ilLoop)
                                'set length (0,1)
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iLen(0) = tmLvf.iLen(0)
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iLen(1) = tmLvf.iLen(1)
                                'set StartTime for log
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iStartTime(0) = tmLcf.iTime(0, illoop)
                                tmLibDaysTimes(UBound(tmLibDaysTimes)).iStartTime(1) = tmLcf.iTime(1, illoop)
                                ReDim Preserve tmLibDaysTimes(0 To UBound(tmLibDaysTimes) + 1) As LibDaysTimes
                            End If
                            
                            ilFoundVer = False
                            For llLoopOnVer = LBound(tmLibVersions) To UBound(tmLibVersions) - 1
                                If tmLcf.iVefCode = tmLibVersions(llLoopOnVer).iVefCode And tmLcf.lLvfCode(illoop) = tmLibVersions(llLoopOnVer).lLvfCode Then
                                    'get the version of this library
                                    If tmLibVersions(llLoopOnVer).iLvfVersion = tmLvf.iVersion Then
                                        'test for earliest/latest dates
                                        If llDate < tmLibVersions(llLoopOnVer).lEarliestDate Then
                                            tmLibVersions(llLoopOnVer).lEarliestDate = llDate
                                        End If
                                        If llLogDate > tmLibVersions(llLoopOnVer).lLatestDate Then
                                            tmLibVersions(llLoopOnVer).lLatestDate = llLogDate
                                        End If
                                        ilFoundVer = True
                                        Exit For    'get next version within the day
                                    End If
                                End If
                            Next llLoopOnVer
                            If Not ilFoundVer Then
                                tmLibVersions(ilUpperVer).iVefCode = tmLcf.iVefCode
                                tmLibVersions(ilUpperVer).lLvfCode = tmLcf.lLvfCode(illoop)
                                tmLibVersions(ilUpperVer).iLvfVersion = tmLvf.iVersion
                                'test for earliest/latest dates
                                'tmLibVersions(ilUpperVer).lEarliestDate = llLogDate
                                'If Format(llDate, "m/d/yy") = "12/31/69" Then
                                If Format(llDate, "ddddd") = "12/31/69" Then
                                    tmLibVersions(ilUpperVer).lEarliestDate = llStartDate
                                Else
                                    tmLibVersions(ilUpperVer).lEarliestDate = llDate
                                End If
                                tmLibVersions(ilUpperVer).lLatestDate = llLogDate
                                ReDim Preserve tmLibVersions(0 To UBound(tmLibVersions) + 1) As LIBVERSIONS
                                ilUpperVer = UBound(tmLibVersions)
                            End If
                        Else
                            Exit For        'ilLoop (0-49)
                        End If
                    Next illoop
                End If
                If ilfirstTime Then
                    ilfirstTime = False
                    tmLcfSrchKey2.iVefCode = ilVefCode
                    'tmLcfSrchKey2.iLogDate(0) = ilStartDate(0)
                    'tmLcfSrchKey2.iLogDate(1) = ilStartDate(1)
                    tmLcfSrchKey2.iLogDate(0) = 0   'ilStartOfSearch(0)
                    tmLcfSrchKey2.iLogDate(1) = 0   'ilStartOfSearch(1)
                    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    ilRet = ilRet
                    gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
                    ilRet = ilRet
                Else
                    ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Loop            '(ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode)
            'vehicle complete; find all the events for each version and write out prepass records
            For llLoopOnVer = 0 To UBound(tmLibVersions) - 1
                ilProcess = False
                If ilExpired Then
                    If tmLibVersions(llLoopOnVer).lLatestDate >= llStartDate And tmLibVersions(llLoopOnVer).lLatestDate <= llEndDate Then
                    'If tmLibVersions(llLoopOnVer).lLatestDate < llStartDate Then
                        ilProcess = True
                    End If
                Else                'active
                    If tmLibVersions(llLoopOnVer).lLatestDate >= llStartDate And tmLibVersions(llLoopOnVer).lEarliestDate <= llEndDate Then         'the library version ends before requested start
                        'still active
                        ilProcess = True
                    End If
                End If
                If ilProcess Then
                    'setup fields that are common to all events within the library version
                    tmGrf.iVefCode = tmLibVersions(llLoopOnVer).iVefCode
                    tmGrf.lCode4 = tmLibVersions(llLoopOnVer).lLvfCode
                    gPackDateLong tmLibVersions(llLoopOnVer).lEarliestDate, tmGrf.iStartDate(0), tmGrf.iStartDate(1)
                    gPackDateLong tmLibVersions(llLoopOnVer).lLatestDate, tmGrf.iDate(0), tmGrf.iDate(1)
                    tmGrf.iCode2 = 0
                    tmLefSrchKey0.iSeqNo = 0
                    tmLefSrchKey0.iStartTime(0) = 0
                    tmLefSrchKey0.iStartTime(1) = 0
                    tmLefSrchKey0.lLvfCode = tmGrf.lCode4   'lvf code
                    ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
                    
                    slFinalDTString = "": ReDim slDTStrings(0)
                    Do While ilRet = BTRV_ERR_NONE And tmLef.lLvfCode = tmGrf.lCode4
                        tmGrf.iCode2 = tmGrf.iCode2 + 1         'increment seq # so report is in proper order
                        tmGrf.iTime(0) = tmLef.iStartTime(0)        'start time of event
                        tmGrf.iTime(1) = tmLef.iStartTime(1)
                        tmGrf.iMissedTime(0) = tmLef.iLen(0)        'event length
                        tmGrf.iMissedTime(1) = tmLef.iLen(1)
                        'tmGrf.iPerGenl(1) = tmLef.iMaxUnits         ' # units for avail
                        'tmGrf.iPerGenl(2) = tmLef.ianfCode          'avail name
                        'tmGrf.iPerGenl(3) = tmLef.iEnfCode          'event name
                        'tmGrf.iPerGenl(4) = tmLef.iEtfCode          'event type
                        tmGrf.iPerGenl(0) = tmLef.iMaxUnits         ' # units for avail
                        tmGrf.iPerGenl(1) = tmLef.ianfCode          'avail name
                        tmGrf.iPerGenl(2) = tmLef.iEnfCode          'event name
                        tmGrf.iPerGenl(3) = tmLef.iEtfCode          'event type
                        tmGrf.lLong = tmLef.lEvtIDCefCode           'event ID
                        'tmGrf.lDollars(1) = tmLef.lCefCode          '3-24-11 show event comment (mainly for RPS automation)
                        'tmGrf.lDollars(2) = tmLef.lCefCode          'debug only
                        tmGrf.lDollars(0) = tmLef.lCefCode          '3-24-11 show event comment (mainly for RPS automation)
                        tmGrf.lDollars(1) = tmLef.lCefCode          'debug only
                        'grfGenDate = generation date for filtering
                        'grfGenTime = Generation time for filtering
                        'grfStartDate - start date of library version
                        'grfDate - end date of library version
                        'grfCode2 = seq # for sorting
                        'grfCode4 = Library Version
                        'grfVEfCode = vehicle code
                        'grfTime = start time of event within library
                        'grfMissedTime = length of event
                        'grfPerGenl(1) = # units for avai
                        'grfPerGenl(2) = Anf (Named AVail) code
                        'grfPerGenl(3) = Event name (ENF) code
                        'grfPerGenl(4) = event type (ETF) code
                        'grfLong = Event ID to CEF table (may not be used)
                        'grfDollars(1) = Event comments (from CEF), mainly to see the automation RPS comments
                        
                        ReDim slMonday(0): ReDim slTuesday(0): ReDim slWednesday(0): ReDim slThursday(0): ReDim slFriday(0): ReDim slSaturday(0): ReDim slSunday(0)
                        slDayOfWeek = ""
                        slCurDay = "": slLibLen = ""
                        For llCounter = 0 To UBound(tmLibDaysTimes) - 1
                            If tmLibDaysTimes(llCounter).iLibVersion = tmLibVersions(llLoopOnVer).iLvfVersion And _
                                tmLibDaysTimes(llCounter).lLvfCode = tmLibVersions(llLoopOnVer).lLvfCode And _
                                tmLibDaysTimes(llCounter).bProcessed = False Then                                       'Date: 5/8/2019 FYM - processed flag
                                
                                'get start times for library version
                                gUnpackTime tmLibDaysTimes(llCounter).iStartTime(0), tmLibDaysTimes(llCounter).iStartTime(1), "A", 1, slStartTimeA
                                
                                'bug found in gUnpackTime returning "12M" instead of "12AM/12PM"
                                If StrComp(slStartTimeA, "12M", vbTextCompare) = 0 Then
                                    slStartTimeA = "12AM"
                                ElseIf StrComp(slStartTimeA, "12N", vbTextCompare) = 0 Then
                                    slStartTimeA = "12PM"
                                End If
                                
                                gUnpackDate tmLibDaysTimes(llCounter).iDays(0), tmLibDaysTimes(llCounter).iDays(1), slCurDay
                                If tmLibDaysTimes(llCounter).iDays(0) <= 7 Then
                                    ilDayOfWeek = tmLibDaysTimes(llCounter).iDays(0)
                                Else
                                    ilDayOfWeek = (gWeekDayStr(slCurDay) + 1)
                                End If

                                slDayOfWeek = mGetDayOfWeek(ilDayOfWeek)        'returns Mo ... Su
                                
                                'get length of the program then add it to end time
                                gUnpackLength tmLibDaysTimes(llCounter).iLen(0), tmLibDaysTimes(llCounter).iLen(1), 3, True, slLibLen
                                gAddTimeLength slStartTimeA, slLibLen, "A", 1, slTimePlusLength, "Y"
                                    
                                'bug found in gUnpackTime returning "12M" instead of "12AM/12PM"
                                If StrComp(slTimePlusLength, "12M", vbTextCompare) = 0 Then
                                     slTimePlusLength = "12AM"
                                ElseIf StrComp(slTimePlusLength, "12N", vbTextCompare) = 0 Then
                                    slTimePlusLength = "12PM"
                                End If
                                
                                'Date: 5/29/2019    FYM
                                'Cleaned up old messy codes
                                Select Case slDayOfWeek
                                Case "Mo"
                                    If mCheckDayTimeDuplicateEntry("Mo", slStartTimeA, slTimePlusLength) = False Then
                                        slMonday(UBound(slMonday)) = "Mo " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slMonday(UBound(slMonday) + 1)
                                    End If
                                Case "Tu"
                                    If mCheckDayTimeDuplicateEntry("Tu", slStartTimeA, slTimePlusLength) = False Then
                                        slTuesday(UBound(slTuesday)) = "Tu " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slTuesday(UBound(slTuesday) + 1)
                                    End If
                                Case "We"
                                    If mCheckDayTimeDuplicateEntry("We", slStartTimeA, slTimePlusLength) = False Then
                                        slWednesday(UBound(slWednesday)) = "We " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slWednesday(UBound(slWednesday) + 1)
                                    End If
                                Case "Th"
                                    If mCheckDayTimeDuplicateEntry("Th", slStartTimeA, slTimePlusLength) = False Then
                                        slThursday(UBound(slThursday)) = "Th " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slThursday(UBound(slThursday) + 1)
                                    End If
                                Case "Fr"
                                    If mCheckDayTimeDuplicateEntry("Fr", slStartTimeA, slTimePlusLength) = False Then
                                        slFriday(UBound(slFriday)) = "Fr " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slFriday(UBound(slFriday) + 1)
                                    End If
                                Case "Sa"
                                    If mCheckDayTimeDuplicateEntry("Sa", slStartTimeA, slTimePlusLength) = False Then
                                        slSaturday(UBound(slSaturday)) = "Sa " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slSaturday(UBound(slSaturday) + 1)
                                    End If
                                Case "Su"
                                    If mCheckDayTimeDuplicateEntry("Su", slStartTimeA, slTimePlusLength) = False Then
                                        slSunday(UBound(slSunday)) = "Su " & slStartTimeA & "-" & slTimePlusLength
                                        ReDim Preserve slSunday(UBound(slSunday) + 1)
                                    End If
                                End Select
                                
                            End If
                        Next llCounter
                        
                        'TTP 9944 - Program Library report - Show Library Times or Show Actual Times
                        If RptSel.optShowTimes(1).Value = True And RptSel.ckcShowTimes.Value = vbChecked Then
                            gUnpackTime tmLef.iStartTime(0), tmLef.iStartTime(1), "A", 1, slStartTimeL
                            If StrComp(slStartTimeL, "12M", vbTextCompare) = 0 Then
                                slStartTimeL = "12AM"
                            ElseIf StrComp(slStartTimeL, "12N", vbTextCompare) = 0 Then
                                slStartTimeL = "12PM"
                            End If
                            
                            'Debug.Print slStartTimeL & " => " & slStartTimeA
                            gPackTime Format(DateAdd("s", DateDiff("s", "00:00:00", slStartTimeA), slStartTimeL), "HH:MM:SSAMPM"), tmGrf.iTime(0), tmGrf.iTime(1)
                        End If
                        
                        'Date: 5/26/2019    FYM
                        'routine that creates the day/time string for each lvf code/versions
                        slFinalDTString = "/" & slLibLen & " " & mCreateDaysTimesString
                    
                        tmScr.sScript = slFinalDTString
                        tmScr.lCode = 0
                        tmScr.iGenDate(0) = tmGrf.iGenDate(0)
                        tmScr.iGenDate(1) = tmGrf.iGenDate(1)
                        tmScr.lGenTime = tmGrf.lGenTime
                    
                        'Insert record to SCR_Script_Check.btr -- Multi-step process:
                        ilRet = btrInsert(hmScr, tmScr, imScrRecLen, INDEXKEY0)
                        'Grab the key from the SCR_script_check then update the GRF key; In the report, use that key (SCR.lCode) to link the two tables (GRF and SCR)
                        tmGrf.lChfCode = tmScr.lCode
                        
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        On Error GoTo gCreatePgmLibraryErr
                        gBtrvErrorMsg ilRet, "gCreatePgmLibrary (btrInsert: GRF.BTR)", RptSel
                        On Error GoTo 0
                        ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            Next llLoopOnVer
            
        End If
    Next ilLoopOnVehicle
    
    ilRet = btrClose(hmLvf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmLef)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmLtf)
    ilRet = btrClose(hmScr)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmMcf)
    
    btrDestroy hmLvf
    btrDestroy hmLcf
    btrDestroy hmLef
    btrDestroy hmGrf
    btrDestroy hmLtf
    btrDestroy hmScr
    btrDestroy hmMnf
    btrDestroy hmAnf
    btrDestroy hmCif
    btrDestroy hmCif
    btrDestroy hmCpf
    btrDestroy hmMcf
    
    Erase tmLibVersions
    Erase tmLibDaysTimes
    Exit Sub
gCreatePgmLibraryErr:
    On Error GoTo 0
    Resume Next
End Sub

'Date: 6/12/2019    FYM
'Function to check for duplicate days / times entry for each day of the week
'
Private Function mCheckDayTimeDuplicateEntry(ByVal slDay As String, ByVal slStartTime As String, ByVal slEndTime As String) As Boolean
    Dim blFound As Boolean
    Dim ilCounter As Integer

    mCheckDayTimeDuplicateEntry = False
    Select Case slDay
    Case "Mo"
        For ilCounter = 0 To UBound(slMonday) - 1
            If Mid(slMonday(ilCounter), InStr(1, slMonday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "Tu"
        For ilCounter = 0 To UBound(slTuesday) - 1
            If Mid(slTuesday(ilCounter), InStr(1, slTuesday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "We"
        For ilCounter = 0 To UBound(slWednesday) - 1
            If Mid(slWednesday(ilCounter), InStr(1, slWednesday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "Th"
        For ilCounter = 0 To UBound(slThursday) - 1
            If Mid(slThursday(ilCounter), InStr(1, slThursday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "Fr"
        For ilCounter = 0 To UBound(slFriday) - 1
            If Mid(slFriday(ilCounter), InStr(1, slFriday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "Sa"
        For ilCounter = 0 To UBound(slSaturday) - 1
            If Mid(slSaturday(ilCounter), InStr(1, slSaturday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    Case "Su"
        For ilCounter = 0 To UBound(slSunday) - 1
            If Mid(slSunday(ilCounter), InStr(1, slSunday(ilCounter), " ") + 1) = (slStartTime & "-" & slEndTime) Then
                mCheckDayTimeDuplicateEntry = True
                Exit For
            End If
        Next ilCounter
    End Select
End Function

'********************************************************
'*                                                       *
'*        Procedure Name:  gCRBgtGen                     *
'*                                                       *
'*           Created 04/18/96     D. hosaka              *
'*                                                       *
'*           Generate 52 weeks budget record (office or  *
'*           vehicle into 12 corporate or standard months*
'*           or 13 weekly buckets                        *
'*                                                       *
'*********************************************************
Sub gCRBgtGen()
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim illoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilWkNo As Integer
    Dim llDollar As Long
    Dim slStr As String
    Dim slStart As String
    Dim slEnd As String
    Dim slDate As String
    Dim slPrevStart As String
    Dim ilAdjust As Integer
    Dim ilLoopAdjust As Integer     'loop factor for week(13), quarter(4) or month(12)
    Dim ilWkStart As Integer        'starting week index to accum $ from bvf
    Dim ilWkEnd As Integer          'ending week index to accum $ from bvf
    Dim ilYrIndex As Integer    'index into ilstartwk or ilendwk representing offset from base year
                                'index 1 = base year of budget comparison, 2 = one year ago, 3 = 2 years ago, etc.
    Dim llDate As Long
    Dim llTemp As Long
    Dim llTemp2 As Long
    Dim slBase As String        'Start date of corporate or std base year
    Dim llBase As Long          'start date of corporate or std base year
    Dim llWeekInput As Long     'Start date of week
    Dim ilWkInputInx As Integer 'index into week # for the base year
    Dim ilMaxHit As Integer     'flag for Crystal indicating that more than 1 base and 4 comparisons selected
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmBvf
        btrDestroy hmGrf
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)
    'ReDim tmCompare(1 To 1) As BUDGETWKS
    ReDim tmCompare(0 To 0) As BUDGETWKS
    'build the budget information for the base budget to compare against
    For illoop = 0 To RptSel!lbcSelection(4).ListCount - 1 Step 1
        If RptSel!lbcSelection(4).Selected(illoop) Then    'selected element
            slNameCode = tgRptSelBudgetCode(illoop).sKey   'RptSel!lbcBudgetCode.List(ilLoop)          'pick up office code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)      'obtain budget name for comparisons
            'tmCompare(1).iMnfCode = Val(slCode)
            tmCompare(0).iMnfCode = Val(slCode)
            ilRet = gParseItem(slNameCode, 1, "/", slCode)       'obtain year of budget name
            slCode = gSubStr("9999", slCode)
            'tmCompare(1).iYear = Val(slCode)
            'tmCompare(1).iIndex = 1                             'base year
            'ReDim Preserve tmCompare(1 To 2) As BUDGETWKS
            tmCompare(0).iYear = Val(slCode)
            tmCompare(0).iIndex = 1                             'base year
            ReDim Preserve tmCompare(0 To 1) As BUDGETWKS
            Exit For                        'only 1 base budget allowed
        End If
    Next illoop

    ilMaxHit = False                        'flag to send to Crystal if more than 4 comparisons selected
    'build the budget information for the comparison budgets
    For illoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
        If RptSel!lbcSelection(2).Selected(illoop) Then    'selected element
            ilUpper = UBound(tmCompare)
            'If ilUpper = 6 Then             'have 5 max (1 base plus 4 comparisons, maxed out due to limitations in Crystal)
            If ilUpper = 5 Then             'have 5 max (1 base plus 4 comparisons, maxed out due to limitations in Crystal)
                ilMaxHit = True
                Exit For
            End If
            slNameCode = tgRptSelBudgetCode(illoop).sKey   'RptSel!lbcBudgetCode.List(ilLoop)          'pick up office code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)      'obtain budget name for comparisons
            tmCompare(ilUpper).iMnfCode = Val(slCode)
            ilRet = gParseItem(slNameCode, 1, "/", slCode)       'obtain year of budget name
            slCode = gSubStr("9999", slCode)
            tmCompare(ilUpper).iYear = Val(slCode)
            tmCompare(ilUpper).iIndex = ilUpper
            'ReDim Preserve tmCompare(1 To UBound(tmCompare) + 1) As BUDGETWKS
            ReDim Preserve tmCompare(0 To UBound(tmCompare) + 1) As BUDGETWKS
        End If
    Next illoop
    'build the vehicle or office codes selected for inclusion
    'ReDim ilOfcVeh(1 To 1) As Integer
    ReDim ilOfcVeh(0 To 0) As Integer
    ilUpper = UBound(ilOfcVeh)
    If RptSel!rbcSelCSelect(0).Value Then           'option by office
        For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
            If RptSel!lbcSelection(1).Selected(illoop) Then    'selected element
                slNameCode = tgSOCode(illoop).sKey 'RptSel!lbcSOCode.List(ilLoop)          'pick up office code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilOfcVeh(ilUpper) = Val(slCode)
                ReDim Preserve ilOfcVeh(0 To ilUpper + 1) As Integer
                ilUpper = ilUpper + 1
            End If
        Next illoop
    Else                                            'option by vehicle
        For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
            If RptSel!lbcSelection(0).Selected(illoop) Then    'selected element
                slNameCode = tgCSVNameCode(illoop).sKey    'RptSel!lbcCSVNameCode.List(ilLoop)          'pick up vehicle code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilOfcVeh(ilUpper) = Val(slCode)
                ReDim Preserve ilOfcVeh(0 To ilUpper + 1) As Integer
                ilUpper = ilUpper + 1
            End If
        Next illoop
    End If
    'Loop thru the Budget entries (to get the year to process to determine # weeks in the periods)
    'For ilUpper = 1 To UBound(tmCompare) - 1 Step 1
    For ilUpper = LBound(tmCompare) To UBound(tmCompare) - 1 Step 1
        'get start of year for corporate or standard
        If RptSel!rbcSelC4(0).Value Then            'corporate calendar
            ilYrIndex = gGetCorpCalIndex(tmCompare(ilUpper).iYear)    'find the index to the start of the corp year record requested
            'gUnpackDateLong tgMCof(ilYrIndex).iStartDate(0, 1), tgMCof(ilYrIndex).iStartDate(1, 1), llDate 'get the start of the year
            gUnpackDateLong tgMCof(ilYrIndex).iStartDate(0, 0), tgMCof(ilYrIndex).iStartDate(1, 0), llDate 'get the start of the year
            slBase = Format$(llDate, "m/d/yy")
        Else
            slStr = "1/15/" & Trim$(str$(tmCompare(ilUpper).iYear))
            slBase = gObtainYearStartDate(0, slStr)
        End If
        'Start date of either corporate or standard year is in slBase
        llBase = gDateValue(slBase)
        If ilUpper = 1 Then                     '1st is always the base
            If RptSel!edcSelCTo.Text = "" Then
                slDate = slBase
            Else
                slDate = RptSel!edcSelCTo.Text
            End If
            llWeekInput = gDateValue(slDate)
            'backup start date to Monday
            illoop = gWeekDayLong(llWeekInput)
            Do While illoop <> 0
                llWeekInput = llWeekInput - 1
                illoop = gWeekDayLong(llWeekInput)
            Loop
            ilWkInputInx = (llWeekInput - llBase) / 7
        End If
        If RptSel!rbcSelCInclude(2).Value Then             'week request
            For illoop = 1 To 13
                If (llWeekInput - llBase) / 7 + illoop > 52 Then
                    Exit For
                Else
                    tmCompare(ilUpper).iStartWks(illoop) = ilWkInputInx + illoop  'setup start and end weeks for weekly
                    tmCompare(ilUpper).iEndWks(illoop) = ilWkInputInx + illoop
                End If
            Next illoop
        ElseIf RptSel!rbcSelCInclude(0).Value Then         'qtr
            For illoop = 1 To 4
                tmCompare(ilUpper).iStartWks(illoop) = (illoop - 1) * 13 + 1
                tmCompare(ilUpper).iEndWks(illoop) = illoop * 13
            Next illoop
        Else                                        'monthly
            If RptSel!rbcSelC4(0).Value Then               'corporate
                For illoop = 1 To 12
                    gUnpackDateLong tgMCof(ilYrIndex).iStartDate(0, illoop - 1), tgMCof(ilYrIndex).iStartDate(1, illoop - 1), llTemp
                    gUnpackDateLong tgMCof(ilYrIndex).iEndDate(0, illoop - 1), tgMCof(ilYrIndex).iEndDate(1, illoop - 1), llTemp2
                    tmCompare(ilUpper).iStartWks(illoop) = (llTemp - llBase) / 7 + 1
                    tmCompare(ilUpper).iEndWks(illoop) = (llTemp2 - llBase) / 7
                Next illoop
            Else                                    'standard
                slStart = slBase
                slPrevStart = slBase
                For illoop = 1 To 13 Step 1
                    slEnd = gObtainEndStd(slStart)
                    tmCompare(ilUpper).iStartWks(illoop) = (gDateValue(slStart) - llBase) / 7 + 1
                    tmCompare(ilUpper).iEndWks(illoop) = (gDateValue(slEnd) - llBase) / 7
                    slStart = gIncOneDay(slEnd)
                Next illoop
            End If
        End If
    Next ilUpper
    If RptSel!rbcSelCInclude(0).Value Then          'do quarters, only pass 1st 4 buckets of data
        ilLoopAdjust = 4
    ElseIf RptSel!rbcSelCInclude(1).Value Then
        ilLoopAdjust = 12
    Else
        ilLoopAdjust = 13
    End If
    ilRet = btrGetFirst(hmBvf, tmBvf, imBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        If RptSel!ckcAll.Value = vbChecked Then
            ilFound = True
        Else
            For illoop = LBound(ilOfcVeh) To UBound(ilOfcVeh) - 1
                If RptSel!rbcSelCSelect(0).Value Then           'option by office
                    If ilOfcVeh(illoop) = tmBvf.iSofCode Then   'mustmatch on selling ofc
                        ilFound = True
                        Exit For
                    End If
                Else                                            'option by vehicle
                    If ilOfcVeh(illoop) = tmBvf.iVefCode Then   'must match on vehicle
                         ilFound = True
                        Exit For
                    End If
                End If
            Next illoop
        End If
        If ilFound Then                                     'found a match, build report recd
            'corrct vehicle or office, if comparison report, need to compare on which
            'buckets to compare against
            ilFound = False
            'For ilUpper = 1 To UBound(tmCompare) - 1
            For ilUpper = LBound(tmCompare) To UBound(tmCompare) - 1
                'Must be a budget that was requested
                If (tmCompare(ilUpper).iMnfCode = tmBvf.iMnfBudget) And (tmCompare(ilUpper).iYear = tmBvf.iYear) Then   'must match on budget name & year
                        ilFound = True
                    Exit For
                End If
            Next ilUpper
        End If
        'determine year index for past years # weeks in each period extracted
        'tables ilstartwk & ilendwk
        If ilFound Then     'must match the base year requested
            tmGrf = tmZeroGrf
            For ilAdjust = 1 To ilLoopAdjust Step 1                    'calc where each of the 52 weeks buckets belong
                                                            'by gathering the # of weeks from start to
                                                            'end for each period
                llDollar = 0
                ilWkStart = tmCompare(ilUpper).iStartWks(ilAdjust)
                ilWkEnd = tmCompare(ilUpper).iEndWks(ilAdjust)
                 For ilWkNo = ilWkStart To ilWkEnd
                    llDollar = llDollar + tmBvf.lGross(ilWkNo)
                Next ilWkNo
                'tmGrf.lDollars(ilAdjust) = llDollar
                tmGrf.lDollars(ilAdjust - 1) = llDollar
            Next ilAdjust
            'Build the remainder of the Crystal record for reporting
            tmGrf.iVefCode = tmBvf.iVefCode         'vehicle code
            tmGrf.iSofCode = tmBvf.iSofCode         'selling office code
            tmGrf.iStartDate(0) = 0
            tmGrf.iStartDate(1) = 0
            tmGrf.iYear = tmBvf.iYear
            tmGrf.iCode2 = tmBvf.iMnfBudget         'budget name
            tmGrf.sBktType = Trim(str$(tmCompare(ilUpper).iIndex))          'get relative index of other comparison budgets
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            'tmGrf.iGenTime(0) = igNowTime(0)
            'tmGrf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            'tmGrf.iPerGenl(1) = ilMaxHit        'set to true if more than 5 budgets selected (1 base + 4 comparisons)
            tmGrf.iPerGenl(0) = ilMaxHit        'set to true if more than 5 budgets selected (1 base + 4 comparisons)
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
        ilRet = btrGetNext(hmBvf, tmBvf, imBvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Erase ilOfcVeh, tmCompare
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmGrf)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCRPlayListGen                  *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Generate Play List Data         *
'*                     for Crystal report              *
'             9/2/99 DH :  add option by advt & cntr
'*                                                     *
'*******************************************************
Sub gCRPlayListGen()
    Dim illoop As Integer
    Dim ilLoopAdv As Integer
    Dim ilRet As Integer
    Dim ilVehicle As Integer
    '3/30/13
    Dim ilVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRec As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilAirIndex As Integer
    Dim ilListIndex As Integer
    Dim ilPLByAdvt As Integer           'true if by advt
    ReDim tlPlayList(0 To 0) As PLAYLIST
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imVlfRecLen = Len(tmVlf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmCpr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    ReDim tmCpr(0 To 0) As CPR
    imCprRecLen = Len(tmCpr(0))
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCif
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imCpfRecLen = Len(tmCpf)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imTzfRecLen = Len(tmTzf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imMcfRecLen = Len(tmMcf)


    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmClf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    '12-27-04 open SSF for airing vehicles to test valid airing day
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "SSf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmSsf
        btrDestroy hmClf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)

    '5-13-04
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPlayListErr
    gBtrvErrorMsg ilRet, "mPlayListErr (btrOpen)", RptSel
    On Error GoTo 0

    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPlayListErr
    gBtrvErrorMsg ilRet, "mPlayListErr (btrOpen)", RptSel
    On Error GoTo 0
    imRsfRecLen = Len(tmRsf)

    hmCvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mPlayListErr
    gBtrvErrorMsg ilRet, "mPlayListErr (btrOpen)", RptSel
    On Error GoTo 0

    ilListIndex = RptSel!lbcRptType.ListIndex
    'If tgSpf.sUseCartNo = "N" Then     7-22-04 no need to adjust report index since all reprts are shown
    '                                   regardless of whether the feature is in use
    '    ilListIndex = ilListIndex + 1
    'End If
        
    If ilListIndex = 14 Then            'by advt (vs vehicle or isci)
        ilPLByAdvt = True
    Else
        ilPLByAdvt = False
    End If
    '8-22-19 use csi calendar control vs text box
'    slStartDate = RptSel!edcSelCFrom.Text
'    slEndDate = RptSel!edcSelCTo.Text
    slStartDate = RptSel!CSI_CalFrom.Text
    slEndDate = RptSel!CSI_CalTo.Text

    ReDim tmSATable(0 To 0) As SATABLE
    tmSATable(0).iAirCode = 0
    tmSATable(0).iSellCode = 0
    tmSATable(0).iNextIndex = -1
    ReDim tmSellAirList(0 To 0) As SELLAIRLIST              '7-23-12
    tmVef.iCode = 0
    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        '3/30/13: Added log vehicle test
        'If tmVef.sType = "C" Or tmVef.sType = "G" Then
        If (tmVef.sType = "C" Or tmVef.sType = "G") And ((tmVef.iVefCode = 0) Or (Not mMergeWithLog(tmVef.iCode))) Then
            For ilVehicle = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                If (RptSel!lbcSelection(2).Selected(ilVehicle)) Then
                    slNameCode = tgAirNameCode(ilVehicle).sKey 'RptSel!lbcAirNameCode.List(ilVehicle)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If (tmVef.iCode = ilVefCode) Or (tmVef.iVefCode = ilVefCode) Then
                        ilFound = False
                        For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                            If tlPlayList(ilIndex).iVefCode = tmVef.iVefCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilIndex
                        If Not ilFound Then
                            tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
                            tlPlayList(UBound(tlPlayList)).iVefCode = tmVef.iCode
                            tlPlayList(UBound(tlPlayList)).iLogCode = tmVef.iVefCode
                            ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
                        End If
                        Exit For
                    End If
                End If
            Next ilVehicle
        '3/30/13: Add log vehicle test
        '3/30/13:  this code was fixed to handle airing that is not merged but will use the airing code
        '          that was added
        'ElseIf tmVef.sType = "S" Then
'        ElseIf (tmVef.sType = "S") And (tmVef.iVefCode = 0) Then
'            gBuildLinkArray hmVlf, tmVef, slStartDate, imAVefCode()
'            For ilVehicle = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
'                If (RptSel!lbcSelection(2).Selected(ilVehicle)) Then
'                    slNameCode = tgAirNameCode(ilVehicle).sKey 'RptSel!lbcAirNameCode.List(ilVehicle)
'                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                    ilVefCode = Val(slCode)
'                    For ilLoop = LBound(imAVefCode) To UBound(imAVefCode) - 1 Step 1
'                        If imAVefCode(ilLoop) = ilVefCode Then
'                            ilFound = False
'                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'                                If tgMVef(ilVef).iCode = ilVefCode Then
'                                    If tgMVef(ilVef).iVefCode <= 0 Then
'                                        ilFound = True
'                                    ElseIf (Not mMergeWithLog(ilVefCode)) Then
'                                        ilFound = True
'                                    End If
'                                    Exit For
'                                End If
'                            Next ilVef
'                            If ilFound Then
'                                ilFound = False
'                                For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
'                                    If tlPlayList(ilIndex).iVefCode = tmVef.iVefCode Then
'                                        ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
'                                        Do While ilAirIndex >= 0
'                                            If tmSATable(ilAirIndex).iAirCode = imAVefCode(ilLoop) Then
'                                                ilFound = True
'                                                Exit Do
'                                            End If
'                                            ilAirIndex = tmSATable(ilAirIndex).iNextIndex
'                                        Loop
'                                        If Not ilFound Then
'                                            ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
'                                            If ilAirIndex < 0 Then
'                                                tlPlayList(ilIndex).iSAFirstIndex = UBound(tmSATable)
'                                                tmSATable(UBound(tmSATable)).iAirCode = imAVefCode(ilLoop)
'                                                tmSATable(UBound(tmSATable)).iNextIndex = -1
'                                                ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
'                                            Else
'                                                Do While ilAirIndex >= 0
'                                                    If tmSATable(ilAirIndex).iNextIndex < 0 Then
'                                                        tmSATable(ilAirIndex).iNextIndex = UBound(tmSATable)
'                                                        tmSATable(UBound(tmSATable)).iAirCode = imAVefCode(ilLoop)
'                                                        tmSATable(UBound(tmSATable)).iNextIndex = -1
'                                                        ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
'                                                        Exit Do
'                                                    End If
'                                                    ilAirIndex = tmSATable(ilAirIndex).iNextIndex
'                                                Loop
'                                            End If
'                                        End If
'                                        ilFound = True
'                                        Exit For
'                                    End If
'                                Next ilIndex
'                                If Not ilFound Then
'                                    tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
'                                    tlPlayList(UBound(tlPlayList)).iVefCode = tmVef.iCode
'                                    tlPlayList(UBound(tlPlayList)).iLogCode = 0
'                                    tlPlayList(UBound(tlPlayList)).iSAFirstIndex = UBound(tmSATable)
'                                    tmSATable(UBound(tmSATable)).iAirCode = imAVefCode(ilLoop)
'                                    tmSATable(UBound(tmSATable)).iNextIndex = -1
'                                    ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
'                                    ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
'                                End If
'                            End If
'                            Exit For
'                        End If
'                    Next ilLoop
'                End If
'            Next ilVehicle
        '3/30/13: Add log vehicle test
        ElseIf tmVef.sType = "L" Then
            For ilVehicle = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                If (RptSel!lbcSelection(2).Selected(ilVehicle)) Then
                    slNameCode = tgAirNameCode(ilVehicle).sKey 'RptSel!lbcAirNameCode.List(ilVehicle)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If (tmVef.iCode = ilVefCode) Then
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G")) And (tgMVef(ilVef).iVefCode = ilVefCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                ilFound = False
                                For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                                    If tlPlayList(ilIndex).iVefCode = tgMVef(ilVef).iCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilIndex
                                If Not ilFound Then
                                    tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
                                    tlPlayList(UBound(tlPlayList)).iVefCode = tgMVef(ilVef).iCode
                                    tlPlayList(UBound(tlPlayList)).iLogCode = tmVef.iCode
                                    ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
                                End If
                                'TTP 10974 - Copy Playlist by ISCI report: log vehicle selected, only showing the airing information for one vehicle that makes up the log vehicle, not all
                                'Exit For
                            '3/30/13: Build playlist build for airing vehicle mapped to selling vehicle
                            ElseIf (tgMVef(ilVef).sType = "A") And (tgMVef(ilVef).iVefCode = ilVefCode) And mMergeWithLog(tgMVef(ilVef).iCode) Then
                                gBuildLinkArray hmVLF, tgMVef(ilVef), slStartDate, igSVefCode()
                                For illoop = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                                    ilFound = False
                                    For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                                        If tlPlayList(ilIndex).iVefCode = igSVefCode(illoop) Then
                                            ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
                                            Do While ilAirIndex >= 0
                                                If tmSATable(ilAirIndex).iAirCode = tgMVef(ilVef).iCode Then
                                                    ilFound = True
                                                    Exit Do
                                                End If
                                                ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                                            Loop
                                            If Not ilFound Then
                                                ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
                                                If ilAirIndex < 0 Then
                                                    tlPlayList(ilIndex).iSAFirstIndex = UBound(tmSATable)
                                                    tmSATable(UBound(tmSATable)).iAirCode = tgMVef(ilVef).iCode
                                                    tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                                    tmSATable(UBound(tmSATable)).iNextIndex = -1
                                                    ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                                Else
                                                    Do While ilAirIndex >= 0
                                                        If tmSATable(ilAirIndex).iNextIndex < 0 Then
                                                            tmSATable(ilAirIndex).iNextIndex = UBound(tmSATable)
                                                            tmSATable(UBound(tmSATable)).iAirCode = tgMVef(ilVef).iCode
                                                            tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                                            tmSATable(UBound(tmSATable)).iNextIndex = -1
                                                            ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                                            Exit Do
                                                        End If
                                                        ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                                                    Loop
                                                End If
                                            End If
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilIndex
                                    If Not ilFound Then
                                        tlPlayList(UBound(tlPlayList)).sType = "L"
                                        tlPlayList(UBound(tlPlayList)).iVefCode = igSVefCode(illoop)
                                        tlPlayList(UBound(tlPlayList)).iLogCode = tmVef.iCode
                                        tlPlayList(UBound(tlPlayList)).iSAFirstIndex = UBound(tmSATable)
                                        tmSATable(UBound(tmSATable)).iAirCode = tgMVef(ilVef).iCode
                                        tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                        tmSATable(UBound(tmSATable)).iNextIndex = -1
                                        ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                        ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
                                    End If
                                Next illoop
                            End If
                        Next ilVef
                    End If
                End If
            Next ilVehicle
        '3/30/13: Add airing vehicle test instead of using the Selling code above
        ElseIf (tmVef.sType = "A") And ((tmVef.iVefCode = 0) Or (Not mMergeWithLog(tmVef.iCode))) Then
            For ilVehicle = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                If (RptSel!lbcSelection(2).Selected(ilVehicle)) Then
                    slNameCode = tgAirNameCode(ilVehicle).sKey 'RptSel!lbcAirNameCode.List(ilVehicle)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVefCode = Val(slCode)
                    If (tmVef.iCode = ilVefCode) Then
                        gBuildLinkArray hmVLF, tmVef, slStartDate, igSVefCode()
                        For illoop = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                            ilFound = False
                            For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                                If tlPlayList(ilIndex).iVefCode = igSVefCode(illoop) Then
                                    ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
                                    Do While ilAirIndex >= 0
                                        If tmSATable(ilAirIndex).iAirCode = tmVef.iCode Then
                                            ilFound = True
                                            Exit Do
                                        End If
                                        ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                                    Loop
                                    If Not ilFound Then
                                        ilAirIndex = tlPlayList(ilIndex).iSAFirstIndex
                                        If ilAirIndex < 0 Then
                                            tlPlayList(ilIndex).iSAFirstIndex = UBound(tmSATable)
                                            tmSATable(UBound(tmSATable)).iAirCode = tmVef.iCode
                                            tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                            tmSATable(UBound(tmSATable)).iNextIndex = -1
                                            ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                        Else
                                            Do While ilAirIndex >= 0
                                                If tmSATable(ilAirIndex).iNextIndex < 0 Then
                                                    tmSATable(ilAirIndex).iNextIndex = UBound(tmSATable)
                                                    tmSATable(UBound(tmSATable)).iAirCode = tmVef.iCode
                                                    tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                                    tmSATable(UBound(tmSATable)).iNextIndex = -1
                                                    ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                                    Exit Do
                                                End If
                                                ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                                            Loop
                                        End If
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilIndex
                            If Not ilFound Then
                                tlPlayList(UBound(tlPlayList)).sType = "S"
                                tlPlayList(UBound(tlPlayList)).iVefCode = igSVefCode(illoop)
                                tlPlayList(UBound(tlPlayList)).iLogCode = 0
                                tlPlayList(UBound(tlPlayList)).iSAFirstIndex = UBound(tmSATable)
                                tmSATable(UBound(tmSATable)).iAirCode = tmVef.iCode
                                tmSATable(UBound(tmSATable)).iSellCode = igSVefCode(illoop)
                                tmSATable(UBound(tmSATable)).iNextIndex = -1
                                ReDim Preserve tmSATable(0 To UBound(tmSATable) + 1) As SATABLE
                                ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
                            End If
                        Next illoop
                    End If
                End If
            Next ilVehicle
            'ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
            '        ilVpfIndex = ilLoop
            '        Exit For
            '    End If
            'Next ilLoop
            'If ilVpfIndex >= 0 Then
            '    For ilVehicle = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
            '        If (RptSel!lbcSelection(2).Selected(ilVehicle)) Then
            '            slNameCode = tgAirNameCode(ilVehicle).sKey 'RptSel!lbcAirNameCode.List(ilVehicle)
            '            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            '            ilVefCode = Val(slCode)
            '            For ilLoop = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
            '                If tgVpf(ilVpfIndex).iGLink(ilLoop) > 0 Then
            '                    If tgVpf(ilVpfIndex).iGLink(ilLoop) = ilVefCode Then
            '                        ilFound = False
            '                        For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
            '                            If tlPlayList(ilIndex).iVefCode = tmVef.iVefCode Then
            '                                For ilAirCode = 1 To 6 Step 1
            '                                    If tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop) Then
            '                                        ilFound = True
            '                                        Exit For
            '                                    End If
            '                                Next ilAirCode
            '                                If Not ilFound Then
            '                                    For ilAirCode = 1 To 6 Step 1
            '                                        If tlPlayList(ilIndex).iAirCode(ilAirCode) = 0 Then
            '                                            tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop)
            '                                            Exit For
            '                                        End If
            '                                    Next ilAirCode
            '                                End If
            '                                ilFound = True
            '                                Exit For
            '                            End If
            '                        Next ilIndex
            '                        If Not ilFound Then
            '                            tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
            '                            tlPlayList(UBound(tlPlayList)).iVefCode = tmVef.iCode
            '                            tlPlayList(UBound(tlPlayList)).iLogCode = 0
            '                            tlPlayList(UBound(tlPlayList)).iAirCode(1) = tgVpf(ilVpfIndex).iGLink(ilLoop)
            '                            ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLIST
            '                        End If
            '                        Exit For
            '                    End If
            '                End If
            '            Next ilLoop
            '        End If
            '    Next ilVehicle
            'End If
        End If
        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'Conventional Vehicles without log
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        '3/30/13: Add log vehicle test
        'If (tlPlayList(ilVehicle).sType = "C" Or tlPlayList(ilVehicle).sType = "G") And (tlPlayList(ilVehicle).iLogCode = 0) Then
        If ((tlPlayList(ilVehicle).sType = "C" Or tlPlayList(ilVehicle).sType = "G")) And (tlPlayList(ilVehicle).iLogCode = 0) Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            ReDim tmCpr(0 To 0) As CPR
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle), ilPLByAdvt
            'Output records
            For ilRec = 0 To UBound(tmCpr) - 1 Step 1
                If ilListIndex = 14 Then    'playlist by advertisre
                      For ilLoopAdv = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                        If RptSel!lbcSelection(0).Selected(ilLoopAdv) Then              'selected slsp
                            slNameCode = tgAdvertiser(ilLoopAdv).sKey 'Traffic!lbcAdvertiser.List(ilLoopAdv)         'pick up slsp code
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmCpr(ilRec).iAdfCode Then
                                ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                                Exit For
                            End If
                        End If
                    Next ilLoopAdv
                Else
                    If ilListIndex = 11 Then            'playlist by isci
                        If RptSel!rbcSelC8(2).Value = True Then     'split copy only
                            'eliminate the generic stuff
                            If tmCpr(ilRec).lHd1CefCode > 0 Then
                                ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                            End If
                        Else
                            ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                        End If
                    Else                            'playlist by vehicle
                        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                    End If
                End If
            Next ilRec
        End If
    Next ilVehicle
    ''Conventional Vehicle with Log
    ReDim tmCpr(0 To 0) As CPR
    '3/30/13: Replace with Log vehicle test
    'For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
    '    If (tlPlayList(ilVehicle).sType = "C" Or tlPlayList(ilVehicle).sType = "G") And (tlPlayList(ilVehicle).iLogCode <> 0) Then
    '        ilVefCode = tlPlayList(ilVehicle).iVefCode
    '        mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle), ilPLByAdvt
    '    End If
    'Next ilVehicle
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        If (tlPlayList(ilVehicle).sType = "L") Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle), ilPLByAdvt
        End If
    Next ilVehicle
    'Output records
    For ilRec = 0 To UBound(tmCpr) - 1 Step 1
        If ilListIndex = 14 Then    'playlist by advertisre
                For ilLoopAdv = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                If RptSel!lbcSelection(0).Selected(ilLoopAdv) Then              'selected slsp
                    slNameCode = tgAdvertiser(ilLoopAdv).sKey 'Traffic!lbcAdvertiser.List(ilLoopAdv)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmCpr(ilRec).iAdfCode Then
                        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                        Exit For
                    End If
                End If
            Next ilLoopAdv
        Else
            If ilListIndex = 11 Then                'playlist by ISCI
                If RptSel!rbcSelC8(2).Value = True Then     'split copy only
                    'eliminate the generic stuff
                    If tmCpr(ilRec).lHd1CefCode > 0 Then
                        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                    End If
                Else
                    ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                End If
            Else                                'playlist by vehicle
                ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
            End If
        End If
    Next ilRec
    'Selling Vehicles
    ReDim tmCpr(0 To 0) As CPR
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        '3/30/13:  Add Log vehicle test
        If (tlPlayList(ilVehicle).sType = "S") And (tlPlayList(ilVehicle).iLogCode = 0) Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle), ilPLByAdvt
        End If
    Next ilVehicle
    'Output records
    For ilRec = 0 To UBound(tmCpr) - 1 Step 1
        If ilListIndex = 14 Then    'playlist by advertisre
                For ilLoopAdv = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                If RptSel!lbcSelection(0).Selected(ilLoopAdv) Then              'selected slsp
                    slNameCode = tgAdvertiser(ilLoopAdv).sKey 'Traffic!lbcAdvertiser.List(ilLoopAdv)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmCpr(ilRec).iAdfCode Then
                        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                        Exit For
                    End If
                End If
            Next ilLoopAdv
        Else
            If ilListIndex = 11 Then                'playlist by isci
                If RptSel!rbcSelC8(2).Value = True Then     'split copy only
                    'eliminate the generic stuff
                    If tmCpr(ilRec).lHd1CefCode > 0 Then
                        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                    End If
                Else
                    ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
                End If
            Else                            'playlist by vehicle
                ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
            End If
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                   ilRet = csiHandleValue(0, 7)
                End If
            End If
        End If
    Next ilRec

    
    '7-23-12 Create the subreports for the report by ISCI if requested
    If ilListIndex = 11 And RptSel!ckcSelC7.Value = vbChecked Then
        For ilRec = LBound(tmSellAirList) To UBound(tmSellAirList) - 1
            tmCpr(0).iGenDate(0) = igNowDate(0)
            tmCpr(0).iGenDate(1) = igNowDate(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCpr(0).lGenTime = lgNowTime
            tmCpr(0).iReady = -1               'flag for the filtering of subreport records
            tmCpr(0).iLen = tmSellAirList(ilRec).iVefAirCode
            tmCpr(0).iAdfCode = tmSellAirList(ilRec).iAdfCode
            tmCpr(0).iVefCode = tmSellAirList(ilRec).iVefCode
            ilIndex = gBinarySearchVef(tmSellAirList(ilRec).iVefCode)
            ilRet = btrInsert(hmCpr, tmCpr(0), imCprRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                   ilRet = csiHandleValue(0, 7)
                End If
            End If

        Next ilRec
    End If
    On Error Resume Next
    Erase tmCpr
    Erase tmSellAirList
    Erase tlPlayList
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmMcf)
    ilRet = btrClose(hmTzf)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCrf)
    ilRet = btrClose(hmRsf)
    ilRet = btrClose(hmCvf)
    btrDestroy hmSsf
    btrDestroy hmClf
    btrDestroy hmMcf
    btrDestroy hmTzf
    btrDestroy hmCpf
    btrDestroy hmRcf
    btrDestroy hmCpr
    btrDestroy hmSdf
    btrDestroy hmVLF
    btrDestroy hmVef
    btrDestroy hmCrf
    btrDestroy hmRsf
    btrDestroy hmCvf
    Exit Sub
mPlayListErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
End Sub

Sub gCRRCGen()
'********************************************************
'*                                                       *
'*        Procedure Name:  gCRRCGen                      *
'*                                                       *
'*           Created 04/18/96     D. hosaka              *
'*                                                       *
'*           Generate Rate Card Prices from RIF          *
'*          8-19-02 fix subscript out of range when
'*            year requested starts in previous yr
'*            (i.e. 2002 starts 12/31/01)
'*                                                       *
'*********************************************************
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim illoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilMonth As Integer
    Dim ilWkNo As Integer
    'ReDim ilStartWk(1 To 14) As Integer
    ReDim ilStartWk(0 To 14) As Integer 'Index zero ignored
    'ReDim ilEndWk(1 To 14) As Integer
    ReDim ilEndWk(0 To 14) As Integer   'Index zero ignored
    Dim llDollar As Long
    Dim llAvgDollar As Long
    Dim slStart As String
    Dim slEnd As String
    Dim slDate As String
    Dim slPrevStart As String
    Dim ilAdjust As Integer
    Dim ilCorpWeeks As Integer
    Dim llDate As Long
    Dim ilPeriods As Integer                    'total periods for month average (else 1)
    Dim ilTotalPeriods As Integer               'total periods for years average
    ReDim ilDate(0 To 1) As Integer
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        btrDestroy hmRif
        btrDestroy hmGrf
        Exit Sub
    End If
    imRifRecLen = Len(tmRif)
    'build the vehicle or office codes selected for inclusion
    'ReDim ilVeh(1 To 1) As Integer
    ReDim ilVeh(0 To 0) As Integer
    ilUpper = UBound(ilVeh)
    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
        If RptSel!lbcSelection(0).Selected(illoop) Then    'selected element
            slNameCode = tgCSVNameCode(illoop).sKey    'RptSel!lbcCSVNameCode.List(ilLoop)          'pick up vehicle code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVeh(ilUpper) = Val(slCode)
            'ReDim Preserve ilVeh(1 To ilUpper + 1) As Integer
            ReDim Preserve ilVeh(0 To ilUpper + 1) As Integer
            ilUpper = ilUpper + 1
        End If
    Next illoop
    'For ilLoop = 1 To 14 Step 1
    For illoop = 0 To 13 Step 1
        tmGrf.iDateGenl(0, illoop) = 0
        tmGrf.iDateGenl(1, illoop) = 0
    Next illoop
    If RptSel!rbcSelC4(1).Value Then                    'std
        'determine # weeks in each period (standard  month)
        igYear = Val(RptSel!edcSelCFrom)
        slDate = "1/15/" & Trim$(str$(igYear))
        slStart = gObtainStartStd(slDate)
        slPrevStart = slStart
        If RptSel!rbcSelCInclude(1).Value Then          'std month
            For illoop = 1 To 13 Step 1
                slEnd = gObtainEndStd(slStart)
                If illoop = 1 Then
                    ilStartWk(1) = 1
                    'gPackDate slStart, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1)
                    gPackDate slStart, tmGrf.iDateGenl(0, 0), tmGrf.iDateGenl(1, 0)
                End If
                ilEndWk(illoop) = (gDateValue(slEnd) - gDateValue(slPrevStart) + 1) \ 7
                slStart = gIncOneDay(slEnd)
                If illoop < 13 Then
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1
                    'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop + 1), tmGrf.iDateGenl(1, ilLoop + 1)
                    gPackDate slStart, tmGrf.iDateGenl(0, illoop), tmGrf.iDateGenl(1, illoop)
                End If
            Next illoop
            slStart = gIncOneDay(slEnd)
            'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)   'start date of 13th period
        ElseIf RptSel!rbcSelCInclude(0).Value Then          'std quarter
                For illoop = 1 To 4 Step 1
                    'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
                    gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1)
                    For ilWkNo = 1 To 3 Step 1
                        slEnd = gObtainEndStd(slStart)
                        slStart = gIncOneDay(slEnd)
                    Next ilWkNo
                    If illoop = 1 Then
                        ilStartWk(1) = 1
                    End If
                    ilEndWk(illoop) = (gDateValue(slEnd) - gDateValue(slPrevStart) + 1) \ 7
                    slStart = gIncOneDay(slEnd)
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1
                Next illoop
                'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
                gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1)
        Else                                               'std week
            slDate = RptSel!edcSelCTo.Text                 'user input
            llDate = gDateValue(slDate)
            slDate = Format$(llDate, "m/d/yy")

            'gPackDate slDate, ilDate(0), ilDate(1)          'get month/day and year as a valid day
            'use the year entered for the week
            'ilDate(1) = Val(RptSel!edcSelCFrom)
            'gUnpackDate ilDate(0), ilDate(1), slDate
            ilRet = ((gDateValue(slDate) - gDateValue(slStart)) \ 7 + 1)    'get week index
            For illoop = 1 To 13 Step 1
                If illoop = 1 Then
                    ilStartWk(1) = 1
                    If ilRet <> 1 Then
                        ilStartWk(1) = ilRet
                    End If
                    'gPackDate slDate, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1)
                    gPackDate slDate, tmGrf.iDateGenl(0, 0), tmGrf.iDateGenl(1, 0)
                End If
                slEnd = gObtainNextSunday(slDate)        'obtain end of week
                ilEndWk(illoop) = (((gDateValue(slEnd) - gDateValue(slStart) + 1)) \ 7)
                slDate = gIncOneDay(slEnd)
                If illoop < 13 Then
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1
                    'gPackDate slDate, tmGrf.iDateGenl(0, ilLoop + 1), tmGrf.iDateGenl(1, ilLoop + 1)
                    gPackDate slDate, tmGrf.iDateGenl(0, illoop), tmGrf.iDateGenl(1, illoop)
                End If
            Next illoop
            slStart = gIncOneDay(slEnd)
            'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)   'start date of 13th period
            gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1) 'start date of 13th period
        End If
    End If                                      'endif std
    If RptSel!rbcSelC4(0).Value Then            'corporate calendar
        'determine # weeks in each period corp period
        igYear = Val(RptSel!edcSelCFrom)
        slDate = "1/15/" & Trim$(str$(igYear))
        slStart = gObtainStartCorp(slDate, True)

        slEnd = gObtainEndCorp(slStart, True)
        ilStartWk(1) = 1
        If RptSel!rbcSelCInclude(1).Value Then      'corp month  (vs qtr)
            For illoop = 1 To 12 Step 1
                If illoop = 1 Then
                    ilEndWk(1) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
                    ilAdjust = gWeekDayStr(slStart)
                    If ilAdjust <> 0 Then
                        ilEndWk(1) = ilEndWk(1) + 1   'adjust for week of 1/1 thru sunday, plus the
                                                                'remainder from divide
                    End If
                    'gPackDate slStart, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1) 'start of first corp month
                    gPackDate slStart, tmGrf.iDateGenl(0, 0), tmGrf.iDateGenl(1, 0) 'start of first corp month
                Else
                    slStart = gIncOneDay(slEnd)
                    'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
                    gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1)
                    slEnd = gObtainEndCorp(slStart, True)
                    ilStartWk(illoop) = ilEndWk(illoop - 1) + 1
                    ilEndWk(illoop) = (ilStartWk(illoop) + ((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7)) - 1
                End If
            Next illoop
            slStart = gIncOneDay(slEnd)
            'gPackDate slStart, tmGrf.iDateGenl(0, 13), tmGrf.iDateGenl(1, 13)   'start date of 13th period
            gPackDate slStart, tmGrf.iDateGenl(0, 12), tmGrf.iDateGenl(1, 12)   'start date of 13th period
        ElseIf RptSel!rbcSelCInclude(0).Value Then              'corp qtr
            For illoop = 1 To 4 Step 1
                'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
                gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1)
                For ilWkNo = 1 To 3 Step 1
                    If illoop = 1 And ilWkNo = 1 Then
                        ilCorpWeeks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
                        ilAdjust = gWeekDayStr(slStart)
                        If ilAdjust <> 0 Then
                            ilCorpWeeks = ilCorpWeeks + 1   'adjust for week of 1/1 thru sunday, plus the
                                                                'remainder from divide
                        End If
                        'gPackDate slStart, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1) 'start of first corp month
                    Else
                        slStart = gIncOneDay(slEnd)
                        'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
                        slEnd = gObtainEndCorp(slStart, True)
                        ilCorpWeeks = ilCorpWeeks + ((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7)
                    End If
                Next ilWkNo
                ilEndWk(illoop) = ilCorpWeeks
                ilStartWk(illoop + 1) = ilEndWk(illoop) + 1
                slStart = gIncOneDay(slEnd)
            Next illoop
            'gPackDate slStart, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)
            gPackDate slStart, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1)
        Else                                'corp week
            slDate = RptSel!edcSelCTo.Text                 'user input
            llDate = gDateValue(slDate)
            slDate = Format$(llDate, "m/d/yy")
            gPackDate slDate, ilDate(0), ilDate(1)          'get month/day and year as a valid day
            'use the year entered for the week
            ilDate(1) = Val(RptSel!edcSelCFrom)
            gUnpackDate ilDate(0), ilDate(1), slDate

            ilRet = ((gDateValue(slDate) - gDateValue(slStart)) \ 7 + 1)    'get week index
            For illoop = 1 To 13 Step 1
                    If illoop = 1 And ilRet <> 1 Then       'preset to "1" earlier
                        ilStartWk(1) = ilRet
                    End If
                    'gPackDate slDate, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop) 'start of first corp month
                    gPackDate slDate, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1) 'start of first corp month
                    ilEndWk(illoop) = ilStartWk(illoop)
                    If illoop < 13 Then
                        ilStartWk(illoop + 1) = ilEndWk(illoop) + 1
                    End If
                    slEnd = gObtainNextSunday(slDate)
                    slDate = gIncOneDay(slEnd)
            Next illoop
            slDate = gIncOneDay(slEnd)
            'gPackDate slDate, tmGrf.iDateGenl(0, ilLoop), tmGrf.iDateGenl(1, ilLoop)   'start date of 13th period
            gPackDate slDate, tmGrf.iDateGenl(0, illoop - 1), tmGrf.iDateGenl(1, illoop - 1) 'start date of 13th period
        End If
    End If
    If RptSel!rbcSelCInclude(0).Value Then                'qtr
        ilAdjust = 4
    ElseIf RptSel!rbcSelCInclude(1).Value Then
        ilAdjust = 12                               'month
    Else
        ilAdjust = 13                               'week
    End If
    ilRet = btrGetFirst(hmRif, tmRif, imRifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ilFound = False
        For illoop = LBound(ilVeh) To UBound(ilVeh) - 1
            If ilVeh(illoop) = tmRif.iVefCode Then   'And igYear = tmRif.iYear Then     'must match on vehicle
                'year and vehicle OK, is it a R/C that has been selected
                If RptSel!ckcAllRC.Value = vbChecked Then
                    ilFound = True
                    Exit For
                Else
                    For ilPeriods = 0 To RptSel!lbcSelection(11).ListCount - 1 Step 1
                        If RptSel!lbcSelection(11).Selected(ilPeriods) Then
                            slNameCode = tgRateCardCode(ilPeriods).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tmRif.iRcfCode Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilPeriods
                End If
            End If
        Next illoop
        If ilFound Then                                     'found a match, build report recd
            'For ilLoop = 1 To 14 Step 1
            For illoop = LBound(tmGrf.lDollars) To UBound(tmGrf.lDollars) Step 1
                tmGrf.lDollars(illoop) = 0
            Next illoop
            llAvgDollar = 0
            ilTotalPeriods = 0
            For ilMonth = 1 To ilAdjust Step 1                    'calc where each of the 52 weeks buckets belong
                                                            'by gathering the # of weeks from start to
                                                            'end for each period

                llDollar = 0
                ilPeriods = 0
                For ilWkNo = ilStartWk(ilMonth) To ilEndWk(ilMonth) Step 1
                    If tmRif.lRate(ilWkNo) <> 0 Then
                        ilPeriods = ilPeriods + 1
                        ilTotalPeriods = ilTotalPeriods + 1
                    End If
                    If ilWkNo = 1 And RptSel!rbcSelC4(1).Value Then     'if std, get the 1st part of the std week up to 1/1
                                                                        'from the first bucket
                        llDollar = llDollar + tmRif.lRate(0)
                    End If
                    If ilWkNo < 54 Then
                        llDollar = llDollar + tmRif.lRate(ilWkNo)
                    Else
                        Exit For
                    End If
                Next ilWkNo
                llAvgDollar = llAvgDollar + llDollar        'accum all weeks for average
                If ilPeriods = 0 Then
                    ilPeriods = 1
                    'ilTotalPeriods = 1
                End If
                ''tmGrf.lDollars(ilMonth) = llDollar \ (ilEndWk(ilMonth) - ilStartWk(ilMonth) + 1)
                'tmGrf.lDollars(ilMonth) = llDollar \ ilPeriods
                tmGrf.lDollars(ilMonth - 1) = llDollar \ ilPeriods
            Next ilMonth
            If ilTotalPeriods = 0 Then
                ilTotalPeriods = 1
            End If
            
            If RptSel!rbcSelC6(0).Value Then
                'tmGrf.lDollars(14) = llAvgDollar / ilTotalPeriods
                tmGrf.lDollars(13) = llAvgDollar / ilTotalPeriods
            Else
                'tmGrf.lDollars(14) = tmRif.lAcquisitionCost
                tmGrf.lDollars(13) = tmRif.lAcquisitionCost
            End If
            'Build the remainder of the Crystal record for reporting
            tmGrf.iVefCode = tmRif.iVefCode         'vehicle code
            tmGrf.iCode2 = tmRif.iRcfCode            'rate card #
            tmGrf.iRdfCode = tmRif.iRdfCode         'daypart code
            tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
            tmGrf.iGenDate(1) = igNowDate(1)
            'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
            'tmGrf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
        ilRet = btrGetNext(hmRif, tmRif, imRifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Erase ilStartWk
    Erase ilEndWk
    Erase ilVeh
    ilRet = btrClose(hmRif)
    ilRet = btrClose(hmGrf)
End Sub

Private Function mCreateDaysTimesString() As String
    'Date: 5/28/2019    FYM
    'Process the arrays for each day (Mo-Su) and create the final days/times string
    
    Dim slHourOfDay As String
    Dim slPrevHourOfDay As String
    Dim slDaysTimes As String
    Dim slDaysOfWeek As String
    Dim slFinalDayTimeString As String
    Dim ilCounter As Integer
    Dim blDone As Boolean
    Dim ilSemiColon As Integer
    Dim ilPrevCounter As Integer
    Dim blNewGroup As Boolean
    Dim ilCommaPos As Integer
    
    blDone = False
    Do While Not blDone
        slDaysOfWeek = "": slPrevHourOfDay = "": slDaysTimes = ""
        blNewGroup = False
        ilSemiColon = mGetCharPosition(slFinalDayTimeString, ";")
        For ilCounter = 0 To UBound(slMonday) - 1
            If Mid(slMonday(ilCounter), 1, 2) <> "XX" Then
                slHourOfDay = Mid(slMonday(ilCounter), InStr(1, slMonday(ilCounter), " ") + 1, (InStr(1, slMonday(ilCounter), "-") - 1) - (InStr(1, slMonday(ilCounter), " ")))
                slDaysOfWeek = "Mo"
                If slFinalDayTimeString = "" Then
                    slDaysTimes = slMonday(ilCounter)
                Else
                    If ilSemiColon > 0 Then
                        If InStr(ilSemiColon, slFinalDayTimeString, "-") > 0 Then
                            slPrevHourOfDay = right(slFinalDayTimeString, Len(slFinalDayTimeString) - ((InStr(ilSemiColon, slFinalDayTimeString, "-"))))
                        Else
                            slPrevHourOfDay = Mid(slFinalDayTimeString, ilSemiColon + InStr(ilSemiColon, slFinalDayTimeString, " ") + 1)
                        End If
                    Else
                        If InStr(1, slFinalDayTimeString, "-") > 0 Then
                            slPrevHourOfDay = Mid(slFinalDayTimeString, InStr(1, slFinalDayTimeString, "-") + 1)
                        Else
                            slPrevHourOfDay = Mid(slFinalDayTimeString, InStr(1, slFinalDayTimeString, " ") + 1)
                        End If
                    End If
                    slDaysTimes = slMonday(ilCounter)
                    
                    If Abs(DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay))) > 0 Then blNewGroup = True
                End If
                ilPrevCounter = ilCounter
                slMonday(ilCounter) = "XX" & slMonday(ilCounter)
                Exit For
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slTuesday) - 1
            If Mid(slTuesday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slTuesday(ilCounter)
                    slDaysOfWeek = IIF(InStr(1, slDaysTimes, "Mo") > 0, "Mo,Tu", "Tu")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slTuesday(ilCounter) = "XX" & slTuesday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slTuesday(ilCounter), InStr(1, slTuesday(ilCounter), " ") + 1, (InStr(1, slTuesday(ilCounter), "-") - 1) - (InStr(1, slTuesday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(InStr(1, slDaysTimes, "Mo") > 0, "Mo,Tu", "Tu")
                        slDaysTimes = slDaysOfWeek & Mid(slTuesday(ilCounter), InStr(1, slTuesday(ilCounter), " "))
                        slTuesday(ilCounter) = "XX" & slTuesday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slWednesday) - 1
            If Mid(slWednesday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slWednesday(ilCounter)    'Date: 1/6/2020 used the wrong array --> slThursday(ilCounter)
                    slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",We", "We")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slWednesday(ilCounter) = "XX" & slWednesday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slWednesday(ilCounter), InStr(1, slWednesday(ilCounter), " ") + 1, (InStr(1, slWednesday(ilCounter), "-") - 1) - (InStr(1, slWednesday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        If slDaysTimes = "" Then slDaysTimes = slWednesday(ilCounter)
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",We", "We")
                        slDaysTimes = slDaysOfWeek & Mid(slWednesday(ilCounter), InStr(1, slWednesday(ilCounter), " "))
                        slWednesday(ilCounter) = "XX" & slWednesday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slThursday) - 1
            If Mid(slThursday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slThursday(ilCounter)
                    slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Th", "Th")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slThursday(ilCounter) = "XX" & slThursday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slThursday(ilCounter), InStr(1, slThursday(ilCounter), " ") + 1, (InStr(1, slThursday(ilCounter), "-") - 1) - (InStr(1, slThursday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Th", "Th")
                        slDaysTimes = slDaysOfWeek & Mid(slThursday(ilCounter), InStr(1, slThursday(ilCounter), " "))
                        slThursday(ilCounter) = "XX" & slThursday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slFriday) - 1
            If Mid(slFriday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slFriday(ilCounter)
                    slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Fr", "Fr")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slFriday(ilCounter) = "XX" & slFriday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slFriday(ilCounter), InStr(1, slFriday(ilCounter), " ") + 1, (InStr(1, slFriday(ilCounter), "-") - 1) - (InStr(1, slFriday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Fr", "Fr")
                        slDaysTimes = slDaysOfWeek & Mid(slFriday(ilCounter), InStr(1, slFriday(ilCounter), " "))
                        slFriday(ilCounter) = "XX" & slFriday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slSaturday) - 1
            If Mid(slSaturday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slSaturday(ilCounter)
                    slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Sa", "Sa")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slSaturday(ilCounter) = "XX" & slSaturday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slSaturday(ilCounter), InStr(1, slSaturday(ilCounter), " ") + 1, (InStr(1, slSaturday(ilCounter), "-") - 1) - (InStr(1, slSaturday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Sa", "Sa")
                        slDaysTimes = slDaysOfWeek & Mid(slSaturday(ilCounter), InStr(1, slSaturday(ilCounter), " "))
                        slSaturday(ilCounter) = "XX" & slSaturday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes = "" Then ilPrevCounter = 0
        For ilCounter = 0 To UBound(slSunday) - 1
            If Mid(slSunday(ilCounter), 1, 2) <> "XX" Then
                If slDaysTimes = "" Then
                    slDaysTimes = slSunday(ilCounter)
                    slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Su", "Su")
                    ilPrevCounter = ilCounter
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon)) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    Else
                        If (Mid(slFinalDayTimeString, 1, (InStr(1, slFinalDayTimeString, " ")))) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) Then
                            blNewGroup = False
                        Else
                            blNewGroup = True
                        End If
                    End If
                    slSunday(ilCounter) = "XX" & slSunday(ilCounter)
                    Exit For
                Else
                    slHourOfDay = Mid(slSunday(ilCounter), InStr(1, slSunday(ilCounter), " ") + 1, (InStr(1, slSunday(ilCounter), "-") - 1) - (InStr(1, slSunday(ilCounter), " ")))
                    If InStr(1, slDaysTimes, "-") > 0 Then
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1, (InStr(1, slDaysTimes, "-") - 1) - (InStr(1, slDaysTimes, " ")))
                    Else
                        slPrevHourOfDay = Mid(slDaysTimes, InStr(1, slDaysTimes, " ") + 1)
                    End If
                    If (DateDiff("h", CDate(slPrevHourOfDay), CDate(slHourOfDay)) = 0) And (ilPrevCounter = ilCounter) Then
                        slDaysOfWeek = IIF(slDaysOfWeek <> "", slDaysOfWeek & ",Su", "Su")
                        slDaysTimes = slDaysOfWeek & Mid(slSunday(ilCounter), InStr(1, slSunday(ilCounter), " "))
                        slSunday(ilCounter) = "XX" & slSunday(ilCounter)
                        ilPrevCounter = ilCounter
                        Exit For
                    End If
                End If
            End If
        Next ilCounter
        
        If slDaysTimes <> "" Then
            If slFinalDayTimeString = "" Then
                slFinalDayTimeString = slDaysTimes
            Else
                ilCommaPos = mGetCharPosition(slFinalDayTimeString, ",")
                If InStr(1, slFinalDayTimeString, slDaysOfWeek) > 0 Then
                    If ilSemiColon > 0 Then
                        If (Mid(slFinalDayTimeString, ilSemiColon + 1, (InStr(ilSemiColon, slFinalDayTimeString, " ")) - ilSemiColon) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ")) And Not blNewGroup) Then
                            If InStr(ilSemiColon, slFinalDayTimeString, ",") > 0 Then
                                'there are multiple commas in the string after the last semi-colon
                                If (Abs(DateDiff("h", CDate(right(slFinalDayTimeString, Len(slFinalDayTimeString) - InStr(InStr(ilCommaPos, slFinalDayTimeString, ","), slFinalDayTimeString, "-"))), CDate(right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, "-"))))) <= 1) Or (Abs(DateDiff("h", CDate(right(slFinalDayTimeString, Len(slFinalDayTimeString) - InStr(InStr(ilCommaPos, slFinalDayTimeString, ","), slFinalDayTimeString, "-"))), CDate(right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, "-"))))) = 23) Then
                                    slFinalDayTimeString = Mid(slFinalDayTimeString, 1, (InStr(ilCommaPos, slFinalDayTimeString, "-"))) & right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, "-"))
                                Else
                                    slFinalDayTimeString = slFinalDayTimeString & "," & right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, " "))
                                End If
                            Else
                                If Abs(DateDiff("h", CDate(right(slFinalDayTimeString, Len(slFinalDayTimeString) - InStr(ilSemiColon, slFinalDayTimeString, "-"))), CDate(right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, "-"))))) <= 1 Then
                                    slFinalDayTimeString = Mid(slFinalDayTimeString, 1, (InStr(ilSemiColon, slFinalDayTimeString, "-"))) & right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, "-"))
                                Else
                                    slFinalDayTimeString = slFinalDayTimeString & "," & right(slDaysTimes, Len(slDaysTimes) - InStr(1, slDaysTimes, " "))
                                End If
                            End If
                        Else
                            slFinalDayTimeString = slFinalDayTimeString & ";" & slDaysTimes
                        End If
                    Else
                        If InStr(1, slFinalDayTimeString, "-") > 0 Then
                            If (Mid(slFinalDayTimeString, 1, InStr(1, slFinalDayTimeString, " ") - 1) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ") - 1) And (Not blNewGroup)) Then
                                slFinalDayTimeString = Mid(slFinalDayTimeString, 1, InStr(1, slFinalDayTimeString, "-")) & Mid(slDaysTimes, InStr(1, slDaysTimes, "-") + 1)
                            Else
                                slFinalDayTimeString = slFinalDayTimeString & ";" & slDaysTimes
                            End If
                        Else
                            If Mid(slFinalDayTimeString, 1, InStr(1, slFinalDayTimeString, " ") - 1) = Mid(slDaysTimes, 1, InStr(1, slDaysTimes, " ") - 1) Then
                                slFinalDayTimeString = slFinalDayTimeString & "-" & Mid(slDaysTimes, InStr(1, slDaysTimes, "-") + 1)
                            Else
                                'to do
                                'Stop
                            End If
                        End If
                    End If
                Else
                    slFinalDayTimeString = slFinalDayTimeString & ";" & slDaysTimes
                End If
            End If
        Else
            blDone = True
        End If
    Loop
    
    If slFinalDayTimeString <> "" Then
        'call to reformat the final string
        mReformatFinalDayTimeString slFinalDayTimeString
    End If
    mCreateDaysTimesString = slFinalDayTimeString
End Function

Private Sub mReformatFinalDayTimeString(slFinalDayTimeString As String)
    'Date: 5/29/2019    FYM
    'This routine cleans up the final days/times string -- converts "Mo,Tu,We,Th,Fr 12AM-6AM" to "Mo-Fr 12AM-6AM"
    
    Dim slDayTimeSplit() As String
    Dim ilCounter As Integer
    Dim slTemp As String
    Dim ilSplitArrayCount As Integer
    
    slDayTimeSplit() = Split(slFinalDayTimeString, ";")
    
    ilSplitArrayCount = UBound(slDayTimeSplit)
    If slDayTimeSplit(UBound(slDayTimeSplit)) = "" Then ilSplitArrayCount = UBound(slDayTimeSplit) - 1
    
    For ilCounter = 0 To ilSplitArrayCount
        slTemp = Mid(slDayTimeSplit(ilCounter), 1, InStr(1, slDayTimeSplit(ilCounter), " ") - 1)
        If mGetCharPosition(slTemp, ",") > 3 Then
            slDayTimeSplit(ilCounter) = Replace(slDayTimeSplit(ilCounter), Mid(slDayTimeSplit(ilCounter), InStr(1, slDayTimeSplit(ilCounter), ","), (InStr(1, slDayTimeSplit(ilCounter), " ") - InStr(1, slDayTimeSplit(ilCounter), ",") - 2)), "-")
        End If
        
        If ilCounter = 0 Then
            slFinalDayTimeString = slDayTimeSplit(0)
        Else
            slFinalDayTimeString = slFinalDayTimeString & "," & slDayTimeSplit(ilCounter)
        End If
    Next ilCounter
    
End Sub

Private Function mGetCharPosition(ByVal slDaysTimes As String, slAsciiChar As String) As Integer
    Dim ilCharPos() As Integer
    Dim ilCharPostCounter As Integer
    Dim ilSemiPos As Integer
    Dim blSemiFound As Boolean
    Dim i As Integer
    'count semi-colon character
    Dim asciiToSearchFor As Integer
    
    ReDim ilCharPos(0)

    blSemiFound = False
    asciiToSearchFor = Asc(slAsciiChar)
    For i = 1 To Len(slDaysTimes)
        If Asc(Mid$(slDaysTimes, i, 1)) = asciiToSearchFor Then
            blSemiFound = True                  'found a semi-colon; need to loop through slDaysTimes (e.g. Mo,Tu,We 12AM-3AM;Su 8AM-)
            ilCharPos(UBound(ilCharPos)) = i
            i = i + 1
            ilCharPostCounter = ilCharPostCounter + 1
            'return the last position of the asc character being searched for
            mGetCharPosition = ilCharPos(UBound(ilCharPos))
            ReDim Preserve ilCharPos(0 To UBound(ilCharPos) + 1)
        End If
    Next
End Function

Private Function mGetDayOfWeek(iDOW As Integer) As String
    Select Case iDOW
    Case 1
        mGetDayOfWeek = "Mo"
    Case 2
        mGetDayOfWeek = "Tu"
    Case 3
        mGetDayOfWeek = "We"
    Case 4
        mGetDayOfWeek = "Th"
    Case 5
        mGetDayOfWeek = "Fr"
    Case 6
        mGetDayOfWeek = "Sa"
    Case 7
        mGetDayOfWeek = "Su"
    End Select
End Function

'************************************************************
'*                                                          *
'*      Procedure Name:mGetPlayList                         *
'*                                                          *
'*             Created:10/09/93      By:D. LeVine           *
'*            Modified:              By:                    *
'*                                                          *
'*            Comments:Obtain the Sdf records to be         *
'*                     reported                             *
'*          12/23/98 dh : if not using cart #, show the     *
'*          reel # field and show in cart # column          *
'*      12/27/04    Test for valid airing day for an airing *
'                   vehicle.  If a airing vehicles library  *
'                   isnt M-F (i.e. Tuesday/Thursday not     *
'                   defined), the spots were still included *
'                   from the selling vehicles Tu/Th day     *
'*                                                          *
'************************************************************
Sub mGetPlayList(ilFdVefCode As Integer, slStartDate As String, slEndDate As String, tlPlayList As PLAYLIST, ilPLByAdvt As Integer)
'
'
'   Where
'   ilFdVefCode = vehicle code to process
'   slstartDate - user entered start date
'   slenddate - user entered end date
'   tlPlayList - structure of vehicles to process
'   ilPLByAdvt - true if by advertiser (build records based on unique advt & cntr), vs
'               by vehicle or isci (build records based on unique advt & vehicle)
'
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim slProduct As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slCart As String
    Dim slZone As String
    Dim slDate As String
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilAirIndex As Integer
    Dim ilVefIndex As Integer           '5-13-04
    Dim ilVpfIndex As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim blTestAirTimeUnits As Boolean       '10-12-16           Test honoring zero units feature
    Dim blAirExistsToHonorZeroUnits As Boolean           '10-14-16   true if the vehicle processing is Selling and airing exists to honor zero units
    Dim ilVff As Integer

    'ReDim ilVlfStartDate0(1 To 6) As Integer
    'ReDim ilVlfStartDate1(1 To 6) As Integer
    ReDim ilCurrDate(0 To 1) As Integer
    'ReDim ilVefCode(1 To 6) As Integer
    Dim ilVef As Integer
    Dim ilTerminated As Integer
    
    
    '10-13-16 Need to see if an airing vehicle, and to ignore avails defined as 0 units (HonorZeroUnits feature)
    blTestAirTimeUnits = False
    blAirExistsToHonorZeroUnits = False         '10-14-16 this is a flag to indicate that for a selling vehicle, at least one of the airing vehicles selected has feature
                                                'set to ignore avails with 0 units and or seconds defined (HonorZeroUnits in vehicle options)
    If tlPlayList.sType = "S" Then
        For ilAirIndex = LBound(tmSATable) To UBound(tmSATable) - 1
            ilVff = gBinarySearchVff(tmSATable(ilAirIndex).iAirCode)
            If ilVff <> -1 Then
                If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                    blAirExistsToHonorZeroUnits = True
                    Exit For
                End If
            End If
        Next ilAirIndex
    End If
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilFdVefCode
    gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilFdVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    'If vehicle is selling- obtain start date of Vlf
                    'ReDim ilVefCode(1 To 1) As Integer
                    ReDim ilVefCode(0 To 0) As Integer
                    If tlPlayList.sType = "S" Then
                        If (ilCurrDate(0) <> tmSdf.iDate(0)) Or (ilCurrDate(1) <> tmSdf.iDate(1)) Or (blAirExistsToHonorZeroUnits) Then '10-14-16 processing all the airing vehicles for a single selling vehicles
                                                                                                                    'fall thru here if at least 1 of the airing vehicles is honoring zero units, or it was a different ssf read
                            ilCurrDate(0) = tmSdf.iDate(0)
                            ilCurrDate(1) = tmSdf.iDate(1)
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                            ilDay = gWeekDayStr(slDate)
                            If (ilDay >= 0) And (ilDay <= 4) Then
                                slDay = "0"
                            ElseIf ilDay = 5 Then
                                slDay = "6"
                            Else
                                slDay = "7"
                            End If
                            'For ilVef = 1 To 6 Step 1
                            '    If tlPlayList.iAirCode(ilVef) > 0 Then
                            ilAirIndex = tlPlayList.iSAFirstIndex
                            'find the true start date that encompasses the start date of the link
                            Do While ilAirIndex >= 0
                                    'ilVlfStartDate0(ilVef) = 0
                                    'ilVlfStartDate1(ilVef) = 0
                                     'if airing vehicle, see if zero units should be excluded
                                    tmVlfSrchKey0.iSellTime(0) = 0
                                    tmVlfSrchKey0.iSellTime(1) = 6144    '24*256
                                    blTestAirTimeUnits = False
                                    ilVff = gBinarySearchVff(tmSATable(ilAirIndex).iAirCode)
                                    If ilVff <> -1 Then
                                        If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                                            blTestAirTimeUnits = True
                                            tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
                                            tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)
                                        End If
                                    End If
'                                    If blTestAirTimeUnits Then          '10-13-16 use time of selling vehicle to search the links if honoring zero units, need to test the avail
'                                        tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
'                                        tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)
'                                    End If
                                    tmSATable(ilAirIndex).iStartDate(0) = 0
                                    tmSATable(ilAirIndex).iStartDate(1) = 0
                                    tmVlfSrchKey0.iSellCode = ilFdVefCode
                                    tmVlfSrchKey0.iSellDay = Val(slDay)
                                    tmVlfSrchKey0.iEffDate(0) = tmSdf.iDate(0)
                                    tmVlfSrchKey0.iEffDate(1) = tmSdf.iDate(1)
'                                    tmVlfSrchKey0.iSellTime(0) = 0
'                                    tmVlfSrchKey0.iSellTime(1) = 6144    '24*256
                                    tmVlfSrchKey0.iSellPosNo = 32000
                                    tmVlfSrchKey0.iSellSeq = 32000
                                    ilRet = btrGetLessOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilFdVefCode) And (tmVlf.iSellDay = Val(slDay))
                                        ilTerminated = False
                                        If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                            If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                        If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                            'ilVlfStartDate0(ilVef) = tmVlf.iEffDate(0)
                                            'ilVlfStartDate1(ilVef) = tmVlf.iEffDate(1)
                                            If tmVlf.iAirCode = tmSATable(ilAirIndex).iAirCode Then
                                                '10-14-16 processing airing vehicle- see if honoring zero units (vehicleoptions feature)
'                                                blTestAirTimeUnits = False
'                                                ilVff = gBinarySearchVff(tmSATable(ilAirIndex).iAirCode)
'                                                If ilVff <> -1 Then
'                                                    If tgVff(ilVff).sHonorZeroUnits = "Y" Then
'                                                        blTestAirTimeUnits = True
'                                                    End If
'                                                End If
                                                If blTestAirTimeUnits Then
                                                    If tmSdf.iTime(0) <> tmVlf.iSellTime(0) Or tmSdf.iTime(1) <> tmVlf.iSellTime(1) Then
                                                        ilRet = False
                                                    Else
                                                        ilRet = gTestAirVefValidDay(hmSsf, slDate, tmVlf.iAirCode, tmVlf, blTestAirTimeUnits)
                                                    End If
                                                Else
                                                    ilRet = gTestAirVefValidDay(hmSsf, slDate, tmVlf.iAirCode, tmVlf, blTestAirTimeUnits)
                                                End If
'                                                '12-27-04   determine if this day is a valid air vehicle air date
'                                                ilRet = gTestAirVefValidDay(hmSsf, slDate, tmVlf.iAirCode, tmVlf, blTestAirTimeUnits)
                                                If ilRet Then       'found valid airing day
                                                    tmSATable(ilAirIndex).iStartDate(0) = tmVlf.iEffDate(0)
                                                    tmSATable(ilAirIndex).iStartDate(1) = tmVlf.iEffDate(1)
                                                    Exit Do
                                                Else
                                                    tmSATable(ilAirIndex).iStartDate(0) = 0
                                                    tmSATable(ilAirIndex).iStartDate(1) = 0
                                                End If
                                            End If
                                        End If
                                        ilRet = btrGetPrevious(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                'Else
                                '    tmSATable(ilAirIndex).iStartDate(0) = 0
                                '    tmSATable(ilAirIndex).iStartDate(1) = 0
                                'End If
                            'Next ilVef
                                ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                            Loop
                        End If
                        'For ilVef = 1 To 6 Step 1
                        '    ilVefCode(ilVef) = 0
                        'Next ilVef
                        'For ilVef = 1 To 6 Step 1
                        '    If tlPlayList.iAirCode(ilVef) > 0 Then
                        ilAirIndex = tlPlayList.iSAFirstIndex
                        'gather each airing vehicle that links to the selling vehicle; include multiple entries for the same airing vehicle
                        Do While ilAirIndex >= 0
                                'ilVefCode(ilVef) = 0
                            If (tmSATable(ilAirIndex).iStartDate(0) <> 0 Or tmSATable(ilAirIndex).iStartDate(1) <> 0) Then      '12-28-04
                                tmVlfSrchKey0.iSellCode = ilFdVefCode
                                tmVlfSrchKey0.iSellDay = Val(slDay)
                                tmVlfSrchKey0.iEffDate(0) = tmSATable(ilAirIndex).iStartDate(0) 'ilVlfStartDate0(ilVef)
                                tmVlfSrchKey0.iEffDate(1) = tmSATable(ilAirIndex).iStartDate(1) '
                                tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
                                tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)    '24*256
                                tmVlfSrchKey0.iSellPosNo = 0
                                tmVlfSrchKey0.iSellSeq = 0
                                ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilFdVefCode) And (tmVlf.iSellDay = Val(slDay)) And (tmVlf.iSellTime(0) = tmSdf.iTime(0)) And (tmVlf.iSellTime(1) = tmSdf.iTime(1))
                                    ilTerminated = False
                                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                            ilTerminated = True
                                        End If
                                    End If
                                    If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                        If tmVlf.iAirCode = tmSATable(ilAirIndex).iAirCode Then
                                            ilVefCode(UBound(ilVefCode)) = tmSATable(ilAirIndex).iAirCode
                                            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                                            'Exit Do            '5-13-04 get all avails for the airing
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            'Else
                            '    ilVefCode(ilVef) = 0
                            'End If
                        'Next ilVef
                            End If
                            ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                        Loop
                    Else
                        'For ilVef = 1 To 6 Step 1
                        '    ilVefCode(ilVef) = 0
                        'Next ilVef
                        'ReDim ilVefCode(1 To 2) As Integer
                        ReDim ilVefCode(0 To 1) As Integer
                        If tlPlayList.iLogCode > 0 Then
                            'ilVefCode(1) = tlPlayList.iLogCode
                            ilVefCode(0) = tlPlayList.iLogCode
                        Else
                            'ilVefCode(1) = tmSdf.iVefCode
                            ilVefCode(0) = tmSdf.iVefCode
                        End If
                    End If
                    For ilVef = LBound(ilVefCode) To UBound(ilVefCode) - 1 Step 1       'each link to the airing has its own entry
                        If ilVefCode(ilVef) > 0 Then
                            '5-14-04 need to vehicle info to get airing copy if applicable
                            ilVefIndex = gBinarySearchVef(ilVefCode(ilVef))
                            If ilVefIndex = -1 Then

                                Exit Sub            'vehicle doesn't exist
                            End If
                            ilVpfIndex = gBinarySearchVpf(ilVefCode(ilVef))
                            If ilVpfIndex = -1 Then
                                Exit Sub            'vhicle doesn't exist
                            End If


                            '5-14-04 Airing vehicles do not handle copy across zones
                            'For airing copy, we dont create a TZF, and individual
                            'rsf records are created for each zone.  Either we have to create a phoney TZF record, or we have
                            'to create an array by zone of the airing copy.  The problem is to know that we really have
                            'airing copy by a given zone.  In ggetaircopy, we should return back the copy that the zone
                            'was found for.  It could be blank, which means that it doesnt have that zone.  Need to make 4 calls
                            'to ggetaircopy prior to getting copy.
                            'Step 1:  save current values of pttype & sdfcopycode;
                            'step 2: call ggetaircopy for each of the possible zones of the vehicle (vpf table)
                            'step 3:  for each zone, check to seee if unique copy (ggetaircopy) for a zone. On the
                            '           return of ggetaircopy, check to see if generic copy or zone copy
                            'Step 4:  If zone copy, either build TZF or an array.
                            'gGetAirCopy tgMVef(ilVefIndex).sType, ilVefCode(ilVef), ilVpfIndex, tmSdf, hmCrf, hmRsf, slZone

                            If tmSdf.sPtType = "1" Then  '  Single Copy

                                slZone = ""
                                gGetAirCopy tgMVef(ilVefIndex).sType, ilVefCode(ilVef), ilVpfIndex, tmSdf, hmCrf, hmRsf, hmCvf, slZone

                                ' Read CIF using lCopyCode from SDF
                                'tmCifSrchKey.lCode = tmSdf.lCopyCode
                                'ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)

                                mGetCIFInfoForPlayList tmSdf.lCopyCode, ilPLByAdvt, ilVefCode(ilVef), 0
                                
                                mPlayListSellAirList tgMVef(ilVefIndex).sType, tgMVef(ilVefIndex).iCode, tmSdf.iVefCode, tmSdf.iAdfCode     '7-23-12 determine the associated vehicles with the advt

                                'ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                               4-16-07 make into subroutine to implement split copy
'                                If ilRet = BTRV_ERR_NONE Then
'                                    'initialize all fields incase prod/isci not defined 5/3/99
'                                    slZone = ""
'                                    slISCI = ""
'                                    slProduct = ""
'                                    slCreative = ""
'                                    If tmCif.lCpfCode > 0 Then
'                                        tmCpfSrchKey.lCode = tmCif.lCpfCode
'                                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        If ilRet <> BTRV_ERR_NONE Then
'                                            tmCpf.sISCI = ""
'                                            tmCpf.sName = ""
'                                            tmCpf.sCreative = ""
'                                        End If
'                                        slISCI = Trim$(tmCpf.sISCI)
'                                        slProduct = Trim$(tmCpf.sName)
'                                        slCreative = Trim$(tmCpf.sCreative)
'                                    End If
'                                    If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
'                                        If tmCif.iMcfCode <> tmMcf.iCode Then
'                                            tmMcfSrchKey.iCode = tmCif.iMcfCode
'                                            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                            If ilRet <> BTRV_ERR_NONE Then
'                                                tmMcf.sName = ""
'                                            End If
'                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
'                                        Else
'                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
'                                        End If
'                                    Else
'                                        'slCart = ""
'                                        slCart = Trim$(tmCif.sReel)
'                                    End If
'                                    ilFound = False
'                                    If ilPLByAdvt Then          '9-2-99
'                                        For ilLoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
'                                            If (tmCpr(ilLoop).lCntrNo = tmSdf.lChfCode) And (tmCpr(ilLoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(ilLoop).iLen = tmSdf.iLen) Then
'                                                If (Trim$(tmCpr(ilLoop).sProduct) = slProduct) And (Trim$(tmCpr(ilLoop).sZone) = slZone) And (Trim$(tmCpr(ilLoop).sCartNo) = slCart) And (Trim$(tmCpr(ilLoop).sISCI) = slISCI) And (Trim$(tmCpr(ilLoop).sCreative) = slCreative) Then
'                                                    tmCpr(ilLoop).iLineNo = tmCpr(ilLoop).iLineNo + 1
'                                                    ilFound = True
'                                                    Exit For
'                                                End If
'                                            End If
'                                        Next ilLoop
'                                    Else
'                                        For ilLoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
'                                            If (tmCpr(ilLoop).iVefCode = ilVefCode(ilVef)) And (tmCpr(ilLoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(ilLoop).iLen = tmSdf.iLen) Then
'                                                If (Trim$(tmCpr(ilLoop).sProduct) = slProduct) And (Trim$(tmCpr(ilLoop).sZone) = slZone) And (Trim$(tmCpr(ilLoop).sCartNo) = slCart) And (Trim$(tmCpr(ilLoop).sISCI) = slISCI) And (Trim$(tmCpr(ilLoop).sCreative) = slCreative) Then
'                                                    tmCpr(ilLoop).iLineNo = tmCpr(ilLoop).iLineNo + 1
'                                                    ilFound = True
'                                                    Exit For
'                                                End If
'                                            End If
'                                        Next ilLoop
'                                    End If
'                                    If Not ilFound Then
'                                        ilUpper = UBound(tmCpr)
'                                        tmCpr(ilUpper).iGenDate(0) = igNowDate(0)
'                                        tmCpr(ilUpper).iGenDate(1) = igNowDate(1)
'                                        'tmCpr(ilUpper).iGenTime(0) = igNowTime(0)
'                                        'tmCpr(ilUpper).iGenTime(1) = igNowTime(1)
'                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
'                                        tmCpr(ilUpper).lGenTime = lgNowTime
'                                        tmCpr(ilUpper).iVefCode = ilVefCode(ilVef)
'                                        tmCpr(ilUpper).iAdfCode = tmSdf.iAdfCode
'                                        tmCpr(ilUpper).lCntrNo = tmSdf.lChfCode     '9-2-99
'                                        tmCpr(ilUpper).iLen = tmSdf.iLen
'                                        tmCpr(ilUpper).sProduct = slProduct
'                                        tmCpr(ilUpper).sZone = slZone
'                                        tmCpr(ilUpper).sCartNo = slCart
'                                        tmCpr(ilUpper).sISCI = slISCI
'                                        tmCpr(ilUpper).sCreative = slCreative
'                                        tmCpr(ilUpper).iLineNo = 1
'                                        ReDim Preserve tmCpr(0 To ilUpper + 1) As CPR
'                                    End If
'                                End If

                                'process the split copy if applicable
                                'If RptSel!ckcSelC5(0).Value = vbChecked Then
                                '3-4-10 changed to show either generic, split copy for only split copy
                                If Not RptSel!rbcSelC8(0).Value Then       'not generic only, show split copy or split copy only
                                    mFindSplitForPlayList ilPLByAdvt, ilVefCode(ilVef), tlPlayList.iSAFirstIndex
                                End If
                            ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
                            ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy

                                ' Read TZF using lCopyCode from SDF
                                tmTzfSrchKey.lCode = tmSdf.lCopyCode
                                ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    ' Look for the first positive lZone value
                                    For ilIndex = 1 To 6 Step 1
                                        If (tmTzf.lCifZone(ilIndex - 1) > 0) Then ' Process just the first positive Zone
                                            ' Read CIF using lCopyCode from SDF
                                            tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                slZone = Trim$(tmTzf.sZone(ilIndex - 1))
                                                'initialize all fields incase prod/isci not defined 5/3/99
                                                slISCI = ""
                                                slProduct = ""
                                                slCreative = ""
                                                If tmCif.lcpfCode > 0 Then
                                                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                                                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        tmCpf.sISCI = ""
                                                        tmCpf.sName = ""
                                                        tmCpf.sCreative = ""
                                                    End If
                                                    slISCI = Trim$(tmCpf.sISCI)
                                                    slProduct = Trim$(tmCpf.sName)
                                                    slCreative = Trim$(tmCpf.sCreative)
                                                    mPlayListSellAirList tgMVef(ilVefIndex).sType, tgMVef(ilVefIndex).iCode, tmSdf.iVefCode, tmSdf.iAdfCode     '7-23-12 determine the associated vehicles with the advt
                                                End If
                                                If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                                                    If tmCif.iMcfCode <> tmMcf.iCode Then
                                                        tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tmMcf.sName = ""
                                                        End If
                                                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                                    Else
                                                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                                    End If
                                                Else
                                                    'slCart = ""
                                                    slCart = Trim$(tmCif.sReel)
                                                End If
                                                ilFound = False
                                                If ilPLByAdvt Then              'by advt , build by contr & advt
                                                    For illoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                                                        If (tmCpr(illoop).lCntrNo = tmSdf.lChfCode) And (tmCpr(illoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(illoop).iLen = tmSdf.iLen) Then
                                                            If (Trim$(tmCpr(illoop).sProduct) = slProduct) And (Trim$(tmCpr(illoop).sZone) = slZone) And (Trim$(tmCpr(illoop).sCartNo) = slCart) And (Trim$(tmCpr(illoop).sISCI) = slISCI) And (Trim$(tmCpr(illoop).sCreative) = slCreative) Then
                                                                tmCpr(illoop).iLineNo = tmCpr(illoop).iLineNo + 1
                                                                ilFound = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next illoop
                                                Else
                                                    For illoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                                                        If (tmCpr(illoop).iVefCode = ilVefCode(ilVef)) And (tmCpr(illoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(illoop).iLen = tmSdf.iLen) Then
                                                            If (Trim$(tmCpr(illoop).sProduct) = slProduct) And (Trim$(tmCpr(illoop).sZone) = slZone) And (Trim$(tmCpr(illoop).sCartNo) = slCart) And (Trim$(tmCpr(illoop).sISCI) = slISCI) And (Trim$(tmCpr(illoop).sCreative) = slCreative) Then
                                                                tmCpr(illoop).iLineNo = tmCpr(illoop).iLineNo + 1
                                                                ilFound = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next illoop
                                                End If
                                                If Not ilFound Then
                                                    ilUpper = UBound(tmCpr)
                                                    tmCpr(ilUpper).iGenDate(0) = igNowDate(0)
                                                    tmCpr(ilUpper).iGenDate(1) = igNowDate(1)
                                                    'tmCpr(ilUpper).iGenTime(0) = igNowTime(0)
                                                    'tmCpr(ilUpper).iGenTime(1) = igNowTime(1)
                                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                                    tmCpr(ilUpper).lGenTime = lgNowTime
                                                    tmCpr(ilUpper).iVefCode = ilVefCode(ilVef)
                                                    tmCpr(ilUpper).iAdfCode = tmSdf.iAdfCode
                                                    tmCpr(ilUpper).lCntrNo = tmSdf.lChfCode     'use chf code instead of contr #
                                                    tmCpr(ilUpper).iLen = tmSdf.iLen
                                                    tmCpr(ilUpper).sProduct = slProduct
                                                    tmCpr(ilUpper).sZone = slZone
                                                    tmCpr(ilUpper).sCartNo = slCart
                                                    tmCpr(ilUpper).sISCI = slISCI
                                                    tmCpr(ilUpper).sCreative = slCreative
                                                    tmCpr(ilUpper).iLineNo = 1
                                                    tmCpr(ilUpper).lFt2CefCode = tmCif.lCode    '4-1-13
                                                    ReDim Preserve tmCpr(0 To ilUpper + 1) As CPR
                                                End If
                                            End If
                                        End If
                                    Next ilIndex
                                End If
                            End If
                        End If
                    Next ilVef
                End If
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
End Sub

'           Generate prepass for Copy Rotation Report by Advertiser.
'           Selectivity is based on rotation headers Active dates.
'           Selective advertisers, contracts and vehicles also included.
'
'           Cycle thru CRF (Rotation Header) and gather instructions of
'           rotations whose start/end dates are within effective user
'           entered dates, and whose advertisers/contract and vehicles match.
'
Public Sub gGenCopyRotRpt(Optional blExport As Boolean = False)
    Dim ilRet As Integer
    Dim llActiveDate As Long
    Dim slActiveDate As String
    Dim ilActiveDate(0 To 1) As Integer
    Dim llActiveEndDate As Long              '8-3-10 add date span
    Dim slActiveEndDate As String
    Dim ilActiveEndDate(0 To 1) As Integer
    Dim illoop As Integer
    ReDim llChfCodes(0 To 0) As Long
    Dim ilUpper As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilInclVefCodes As Integer            'true = include codes stored in ilusecode array,
                                                'false = exclude codes store din ilusecode array
    ReDim ilUseVefCodes(0 To 0) As Integer       'valid  vehicles codes to process--
    Dim ilInclAdfCodes As Integer            'true = include codes stored in ilusecode array,
                                                'false = exclude codes store din ilusecode array
    ReDim ilUseAdfCodes(0 To 0) As Integer       'valid  advertiser codes to process--
    ReDim tlCrf(0 To 0) As CRF
    Dim llRots As Long
    Dim ilInclude As Integer
    Dim ilLoopOnCnt As Integer
    Dim ilFoundVef As Integer
    
    Dim llEnterDate As Long
    Dim slEnterDate As String
    Dim ilEnterDate(0 To 1) As Integer
    Dim llEnterEndDate As Long              '8-3-10 add date span
    Dim slEnterEndDate As String
    Dim ilEnterEndDate(0 To 1) As Integer
    Dim llRotDateEntered As Long
    
    'TTP 10791 - Copy Rotations by Advertiser report: add special export option
    Dim lmExportCount As Long
    Dim slRepeat As String
    Dim smClientName As String
    Dim tmMnfSrchKey As INTKEY0
    Dim slFileName As String
    Dim slStr As String
    
    'Open btrieve files
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: VEF.BTR)", RptSel
    On Error GoTo 0

    imCHFRecLen = Len(tmChf)    'Save Contr header record length
    hmCHF = CBtrvTable(ONEHANDLE)          'Save CHF handle
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CHF.BTR)", RptSel
    On Error GoTo 0

    imClfRecLen = Len(tmClf)    'Save Contr header record length
    hmClf = CBtrvTable(ONEHANDLE)          'Save CHF handle
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CLF.BTR)", RptSel
    On Error GoTo 0

    imAdfRecLen = Len(tmAdf)    'Save Advertiser record length
    hmAdf = CBtrvTable(ONEHANDLE)          'Save ADF handle
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: AdF.BTR)", RptSel
    On Error GoTo 0

    imCprRecLen = Len(tmTCpr)    'Save prepass CPR record length
    hmCpr = CBtrvTable(ONEHANDLE)          'Save CPR handle
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CPR.BTR)", RptSel
    On Error GoTo 0

    imCrfRecLen = Len(tmCrf)    'Save Copy Rotation header record length
    hmCrf = CBtrvTable(ONEHANDLE)          'Save CRF handle
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CRF.BTR)", RptSel
    On Error GoTo 0

    imCnfRecLen = Len(tmCnf)    'Save Copy instruction record length
    hmCnf = CBtrvTable(ONEHANDLE)          'Save CNF handle
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CNF.BTR)", RptSel
    On Error GoTo 0
    
    imCafRecLen = Len(tmCaf)    'Save Rotation by Game or Team record length
    hmCaf = CBtrvTable(ONEHANDLE)          'Save CAF handle
    ilRet = btrOpen(hmCaf, "", sgDBPath & "Caf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CAF.BTR)", RptSel
    On Error GoTo 0
    
    imVlfRecLen = Len(tmVlf)    'Save Rotation by Game or Team record length
    hmVLF = CBtrvTable(ONEHANDLE)          'Save CAF handle
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: VLF.BTR)", RptSel
    On Error GoTo 0

    imCvfRecLen = Len(tmCvf)        'vehicle list if rotation consists of more than 1 vehicle
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CVF.BTR)", RptSel
    On Error GoTo 0
    
    imTxrRecLen = Len(tmTxr)        'record for subreport to show vehicle list if rotation consists of more than 1 vehicle
    hmTxr = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: TXR.BTR)", RptSel
    On Error GoTo 0
    
    'TTP 10791 - Copy Rotations by Advertiser report: add special export option
    imMnfRecLen = Len(tmMnf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: MNF.BTR)", RptSel
    On Error GoTo 0
    
    imAnfRecLen = Len(tmAnf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: ANF.BTR)", RptSel
    On Error GoTo 0

    imCifRecLen = Len(tmCif)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CIF.BTR)", RptSel
    On Error GoTo 0
    
    imCpfRecLen = Len(tmCpf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: CPF.BTR)", RptSel
    On Error GoTo 0

    imMcfRecLen = Len(tmMcf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenCopyRotErr
    gBtrvErrorMsg ilRet, "gGenCopyRot (btrOpen: MCF.BTR)", RptSel
    On Error GoTo 0

    'TTP 10791 - Copy Rotations by Advertiser report: add special export option
    If blExport = True Then
        RptSel.lacExport.Caption = "Exporting..."
        RptSel.lacExport.Refresh
        lmExportCount = 0
        slRepeat = "A"
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
        
        'Generate Export Filename
        Do
            ilRet = 0
            slFileName = "CopyRot-"
'            'DateRange
'            slFileName = slFileName & Format(RptSel.CSI_CalFrom.Text, "mmddyy")
'            slFileName = slFileName & "To"
'            slFileName = slFileName & Format(RptSel.CSI_CalTo.Text, "mmddyy")
'            slFileName = slFileName & " - "
            'Todays Date
            slFileName = slFileName & Format(gNow, "mmddyy")
            slFileName = slFileName & slRepeat & " " & gFileNameFilter2(Trim$(smClientName))
            slFileName = slFileName & ".csv"
            'Check if exists, make new character
            ilRet = gFileExist(sgExportPath & slFileName)
            If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
                slRepeat = Chr(Asc(slRepeat) + 1)
            End If
        Loop While ilRet = 0
        
        'Create File
        ilRet = gFileOpen(sgExportPath & slFileName, "OUTPUT", hmExport)
        If ilRet <> 0 Then
            MsgBox "Error writing file:" & sgExportPath & slFileName & vbCrLf & "Error:" & ilRet & " - " & Error(ilRet)
            Close #hmExport
            Exit Sub
        End If
        
        'Generate Header
        slStr = "Advertiser, Contract, Rotation number, Vehicle, Restrictions, Entry Date, Feed Date, Feed Status, Earliest Date Spot, Latest Assign, Last Assign Done, Contract Product, Cart Number, ISCI, Creative Title, Inventory Product"
        'Write Header
        Print #hmExport, slStr
        
        'Pop Anf names
        gPopAnf hmAnf, tmAnfTable()
    End If
 
    'get the vehicle codes selected
    'build array of vehicles to include or exclude
    gObtainCodesForMultipleLists 6, tgVehicle(), ilInclVefCodes, ilUseVefCodes(), RptSel
    'get the Advt codes selected
    gObtainCodesForMultipleLists 0, tgAdvertiser(), ilInclAdfCodes, ilUseAdfCodes(), RptSel

    '8-23-19 input dates changed to use csi calendar control vs edit box
'    slActiveDate = RptSel!edcSelCFrom.Text      'determine rotations to gather
    slActiveDate = RptSel!CSI_CalFrom.Text      'determine rotations to gather
    If slActiveDate = "" Then
        slActiveDate = "1/1/1970"
    End If
    llActiveDate = gDateValue(slActiveDate)
    gPackDate slActiveDate, ilActiveDate(0), ilActiveDate(1)

'    slActiveEndDate = RptSel!edcSelCFrom1.Text      'determine active end date rotations to gather
    slActiveEndDate = RptSel!CSI_CalTo.Text      'determine active end date rotations to gather
    If slActiveEndDate = "" Then
        slActiveEndDate = "12/31/2026"
    End If
    llActiveEndDate = gDateValue(slActiveEndDate)
    gPackDate slActiveEndDate, ilActiveEndDate(0), ilActiveEndDate(1)
    
    '11-30-10 implement date entered span for filter
'    slEnterDate = RptSel!edcSelCTo.Text      'determine rotations to gather for date entered
     slEnterDate = RptSel!CSI_CalFrom2.Text      'determine rotations to gather for date entered
    If slEnterDate = "" Then
        slEnterDate = "1/1/1970"
    End If
    llEnterDate = gDateValue(slEnterDate)
    gPackDate slEnterDate, ilEnterDate(0), ilEnterDate(1)

'    slEnterEndDate = RptSel!edcSelCTo1.Text      'determine date entered end date rotations to gather
    slEnterEndDate = RptSel!CSI_CalTo2.Text      'determine date entered end date rotations to gather
    If slEnterEndDate = "" Then
        slEnterEndDate = "12/31/2026"
    End If
    llEnterEndDate = gDateValue(slEnterEndDate)
    gPackDate slEnterEndDate, ilEnterEndDate(0), ilEnterEndDate(1)
    
    ilUpper = LBound(llChfCodes)
    For illoop = 0 To RptSel!lbcSelection(5).ListCount - 1 Step 1
        If RptSel!lbcSelection(5).Selected(illoop) Then
            slNameCode = RptSel!lbcSelection(3).List(illoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            llChfCodes(ilUpper) = Val(slCode)
            ReDim Preserve llChfCodes(0 To ilUpper + 1) As Long
            ilUpper = ilUpper + 1
        End If
    Next illoop

    ilRet = gObtainCrfByDate(RptSel, hmCrf, tlCrf(), slActiveDate, slActiveEndDate)     'gather the active rotation headers,
                                                                        'if no date, get them all
    For llRots = LBound(tlCrf) To UBound(tlCrf) - 1
        gUnpackDateLong tlCrf(llRots).iEntryDate(0), tlCrf(llRots).iEntryDate(1), llRotDateEntered
        '8-11-10 test for inclusion of dormant rotations
        If (RptSel!ckcSelC11(0).Value = vbChecked And tlCrf(llRots).sState = "D") Or tlCrf(llRots).sState <> "D" Then        'include the dormants
            'filter the advertisers and contracts & vehicles
            ilInclude = True
            If Not RptSel!ckcAll.Value = vbChecked Then           'include all advertisers?
                '12-14-05 chg parms to send
                ilInclude = mFilterLists(tlCrf(llRots).iAdfCode, ilInclAdfCodes, ilUseAdfCodes())
                If ilInclude Then
                    'which selective contracts for this advt?
                    For ilLoopOnCnt = LBound(llChfCodes) To UBound(llChfCodes) - 1
                        If tlCrf(llRots).lChfCode = llChfCodes(ilLoopOnCnt) Then
                            'check for valid vehicles, too.
                            '12-14-05 chg parms to send
                            '4/22/14: Handle packages
                            'ilFoundVef = mFilterLists(tlCrf(llRots).iVefCode, ilInclVefCodes, ilUseVefCodes())
                            ilFoundVef = mCheckPackageVehicles(tlCrf(llRots).lChfCode, tlCrf(llRots).iVefCode, slActiveDate, ilInclVefCodes, ilUseVefCodes())
                            If ilFoundVef And llRotDateEntered >= llEnterDate And llRotDateEntered <= llEnterEndDate Then
                                mCreateCnfRecords tlCrf(llRots), ilInclVefCodes, ilUseVefCodes(), blExport
                            End If
                        End If
                    Next ilLoopOnCnt
                End If
            Else
                'check for valid vehicles
                '4/22/14: Handle packages
                'ilFoundVef = mFilterLists(tlCrf(llRots).iVefCode, ilInclVefCodes, ilUseVefCodes())
                ilFoundVef = mCheckPackageVehicles(tlCrf(llRots).lChfCode, tlCrf(llRots).iVefCode, slActiveDate, ilInclVefCodes, ilUseVefCodes())
                If ilFoundVef And llRotDateEntered >= llEnterDate And llRotDateEntered <= llEnterEndDate Then
                    mCreateCnfRecords tlCrf(llRots), ilInclVefCodes, ilUseVefCodes(), blExport
                End If
            End If
        End If          'test for dormant
    Next llRots
    
    'TTP 10791 - Copy Rotations by Advertiser report: add special export option
    If blExport = True Then
        lmExportCount = 0
        Close #hmExport
        If InStr(1, smExportStatus, "Error") > 0 Then
            RptSel.lacExport.Caption = "Export Failed:" & smExportStatus
        Else
            RptSel.lacExport.Caption = "Export Stored in- " & sgExportPath & slFileName
        End If
    End If
    
    Erase ilUseVefCodes, ilUseAdfCodes, llChfCodes
     'close all files
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmCpr)
    btrDestroy hmCpr
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmCaf)
    btrDestroy hmCaf
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmTxr)
    btrDestroy hmTxr
    
    Exit Sub

gGenCopyRotErr:
    On Error GoTo 0
    Resume Next
End Sub

'       mFilterLists - check the option and which list boxes to test
'       for inclusion/exclusion
'
'       <input>
'               'ilWhichField - 0 = advt, 1 = vehicle
'               ilWhichField = 12-14-05 value of field to compare for inclusion/exclusion
'               ilIncludeCodes = true to include codes in array;
'                                false to exclude codes in array
'               ilUseCodes()- array of codes to include/exclude
'       <return> true = include transaction, else false to exclude
'
'       12-14-05 change the parameters.  Send the field to compare rathern
'       than sending a flag and the buffer for record to retrieve item to compare
'
Private Function mFilterLists(ilWhichField As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************
    Dim ilCompare As Integer
    Dim ilTemp As Integer
    Dim ilFoundOption As Integer

    ilFoundOption = False
    ilCompare = ilWhichField            '12-14-05
    'If ilWhichField = 0 Then            'test advt
     '   ilCompare = tlCrf.iAdfCode
    'ElseIf ilWhichField = 1 Then
    '    ilCompare = tlCrf.iVefCode      'test vehicle
    'End If

    If ilIncludeCodes Then
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilCompare Then
                ilFoundOption = True
                Exit For
            End If
        Next ilTemp
    Else
        ilFoundOption = True
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilCompare Then
                ilFoundOption = False
                Exit For
            End If
        Next ilTemp
    End If
    mFilterLists = ilFoundOption
End Function

'       Create the instruction records for the rotation pattern
'       2-2-15  Create records in TXR for subreport to list the vehicle names within a rotation pattern
'       <input> tlCrf - Rotation header buffer
'               ilIncludeCodes - include vehicle codes  or exclude codes (based on selectivity)
'               ilUseCodes() - array of vehicle codes to include or exclude
Private Sub mCreateCnfRecords(tlCrf As CRF, ilIncludeCodes As Integer, ilUseVefCodes() As Integer, Optional blExport As Boolean = False)
    Dim ilRet As Integer
    Dim llCvfCode As Long
    Dim illoop As Integer
    Dim slVehicleNameString As String
    Dim ilVefInx As Integer
    Dim ilLen As Integer
    Dim ilStartPos As Integer
    Dim ilMaxFieldLen As Integer
    Dim blFoundVef As Boolean
    Dim tlCaf As CAF
    tmTCpr.iGenDate(0) = igNowDate(0)
    tmTCpr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmTCpr.lGenTime = lgNowTime
    
    '2-2-15 determine of subrecords need to be created for the vehicle list when more than 1 vehicle is associated with a rotation pattern
    'first determine if rotation entered for more than 1 vehicle, and if filtering of vehicles by user before writing records for subreport
    slVehicleNameString = ""
    llCvfCode = tlCrf.lCvfCode
    tmCvfSrchKey0.lCode = tlCrf.lCvfCode
    ilRet = btrGetGreaterOrEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    tmTxr.lSeqNo = 0
    tmTxr.iGenDate(0) = igNowDate(0)
    tmTxr.iGenDate(1) = igNowDate(1)
    tmTxr.lGenTime = lgNowTime
    Do While tmCvf.lCode = llCvfCode And ilRet = BTRV_ERR_NONE
        tmTxr.lCsfCode = tmCvf.lCrfCode         'rotation code
        tmTxr.lSeqNo = tmTxr.lSeqNo + 1
        For illoop = 0 To 99
            blFoundVef = False
            If tmCvf.iVefCode(illoop) > 0 Then
                blFoundVef = mFilterLists(tmCvf.iVefCode(illoop), ilIncludeCodes, ilUseVefCodes())
                If blFoundVef Then
                    If Trim$(slVehicleNameString) = "" Then
                        ilVefInx = gBinarySearchVef(tmCvf.iVefCode(illoop))
                        If ilVefInx = -1 Then
                            tmVef.sName = ""
                        End If
                        slVehicleNameString = Trim$(tgMVef(ilVefInx).sName)
                    Else
                        ilVefInx = gBinarySearchVef(tmCvf.iVefCode(illoop))
                        If ilVefInx = -1 Then
                            tmVef.sName = ""
                        End If
                        slVehicleNameString = slVehicleNameString & "," & Trim$(tgMVef(ilVefInx).sName)
                    End If
                End If
            Else
                Exit For
            End If
        Next illoop
        llCvfCode = tmCvf.lLkCvfCode
        If llCvfCode > 0 Then
            tmCvfSrchKey0.lCode = tmCvf.lLkCvfCode
            ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            Exit Do
        End If
    Loop
    
    'determine if any rotations by game or team.  Need to set a flag in prepass so that if there
    'isnt any CAF records, that section can be suppressed in Crystal (due to Crystal bug in not suppressing a blank section)
    tmTCpr.sStatus = ""
    tmCafSrchKey1.lCode = tlCrf.lCode
    ilRet = btrGetGreaterOrEqual(hmCaf, tlCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE And tlCaf.lCrfCode = tlCrf.lCode Then
        tmTCpr.sStatus = "S"        'sports copy by game or team
    End If
    
    If tlCrf.iBkoutInstAdfCode = POOLROTATION Then         'blackout pool exists?
        'create a fake record to print because the instructions of inventory do not exist
        tmTCpr.lFt1CefCode = tlCrf.lCode         'rotation code
        tmTCpr.lFt2CefCode = 0
        tmTCpr.iLineNo = 0            'instruction #
        ilRet = btrInsert(hmCpr, tmTCpr, imCprRecLen, INDEXKEY0)
    Else
        tmCnfSrchKey.lCrfCode = tlCrf.lCode
        tmCnfSrchKey.iInstrNo = 0
        ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While tmCnf.lCrfCode = tlCrf.lCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
            If tlCrf.iVefCode > 0 Or (tlCrf.iVefCode = 0 And Trim$(slVehicleNameString) <> "") Then
                tmTCpr.lFt1CefCode = tmCnf.lCrfCode 'rotation code
                tmTCpr.lFt2CefCode = tmCnf.lCifCode 'inventory code
                tmTCpr.iLineNo = tmCnf.iInstrNo     'instruction #
                'TTP 10791 - Copy Rotations by Advertiser report: add special export option
                If blExport = True Then
                    'Fix v81 TTP 10791 - copy rotations by advertiser report - test results
                    smExportStatus = mExportCopyRotRecord(tmTCpr, tlCrf, slVehicleNameString)
                    If InStr(1, smExportStatus, "Error") > 0 Then
                        Exit Do
                    End If
                Else
                    ilRet = btrInsert(hmCpr, tmTCpr, imCprRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                End If
            End If
            ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    
    '2-2-15  If any multi-vehicles defined for the rotation pattern, create records in txr for subreport to print them
    If Trim$(slVehicleNameString) <> "" And blExport = False Then
        ilStartPos = 1
        ilLen = Len(Trim$(slVehicleNameString))
        Do While ilLen > 0
            ilMaxFieldLen = 200
            If ilLen < 200 Then
                ilMaxFieldLen = ilLen
            End If
            
            tmTxr.sText = Mid$(slVehicleNameString, ilStartPos, ilMaxFieldLen)
            ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            Else
                ilLen = ilLen - 200             '200 is max field length that has been written
                ilStartPos = ilStartPos + ilMaxFieldLen
            End If
        Loop
    End If
    Exit Sub
End Sub

'           Generate prepass for the Live Log Actitivity report  which details
'           all events posted for a live program or sports event.
'           12-9-05
'
Public Sub gGenLiveLogRpt()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************
    Dim ilInclVefCodes As Integer            'true = include codes stored in ilusecode array,
                                                'false = exclude codes store din ilusecode array
    ReDim ilUseVefCodes(0 To 0) As Integer       'valid  vehicles codes to process--
    Dim ilRet As Integer
    Dim llLlfLoop As Long
    ReDim tmLlfList(0 To 0) As LLF
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilInclude As Integer
    Dim ilLcf As Integer
    Dim llAirDate As Long
    Dim llUpper As Long
    Dim llLoopOnEvent As Long
    Dim tlEventComplete() As EVENTCOMPLETE

   'Open btrieve files
    imLlfRecLen = Len(tmLlf)    'Save Live Log record length
    hmLlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLlf, "", sgDBPath & "Llf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: LLF.BTR)", RptSel
    On Error GoTo 0

    imLvfRecLen = Len(tmLvf)    'Save library Version record length
    hmLvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: LLF.BTR)", RptSel
    On Error GoTo 0

    imLcfRecLen = Len(tmLcf)    'Save Library calendar
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: LCF.BTR)", RptSel
    On Error GoTo 0

    imGrfRecLen = Len(tmGrf)    'Save prepass GRF record length
    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: GRF.BTR)", RptSel
    On Error GoTo 0

    imLtfRecLen = Len(tmLtf)    'Save library title record length
    hmLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLtf, "", sgDBPath & "Ltf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: LTF.BTR)", RptSel
    On Error GoTo 0

    imLdfRecLen = Len(tmLdf)    'Save Livelog detail record length
    hmLdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLdf, "", sgDBPath & "Ldf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenLiveLogErr
    gBtrvErrorMsg ilRet, "gGenLiveLog (btrOpen: LDF.BTR)", RptSel
    On Error GoTo 0

    
    '8-20-19 use csi calendar control vs edit box
'    slStartDate = RptSel!edcSelCFrom.Text      'determine rotations to gather
    slStartDate = RptSel!CSI_CalFrom.Text      'determine rotations to gather
    llStartDate = gDateValue(slStartDate)
'    slEndDate = RptSel!edcSelCTo.Text
    slEndDate = RptSel!CSI_CalTo.Text
    llEndDate = gDateValue(slEndDate)

    'build array of vehicles to include or exclude
    gObtainCodesForMultipleLists 0, tgVehicle(), ilInclVefCodes, ilUseVefCodes(), RptSel
    'gather the llf records to print
    ilRet = gObtainLLFbyDate(RptSel, hmLlf, tmLlfList(), slStartDate, slEndDate)
    If ilRet Then
        'set generation date and time filter
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        
        ReDim Preserve tlEventComplete(0 To 0) As EVENTCOMPLETE
        llUpper = UBound(tlEventComplete)
        For llLoopOnEvent = LBound(tmLlfList) To UBound(tmLlfList) - 1
            ilInclude = False
            gUnpackDateLong tmLlfList(llLoopOnEvent).iAirDate(0), tmLlfList(llLoopOnEvent).iAirDate(1), llAirDate
            If (tlEventComplete(llLoopOnEvent).iVefCode = tmLlfList(llLoopOnEvent).iVefCode) And (tlEventComplete(llLoopOnEvent).iGameNo = tmLlfList(llLoopOnEvent).iGameNo) And (tlEventComplete(llLoopOnEvent).lAirDate = llAirDate) Then
                If tlEventComplete(llLoopOnEvent).sLogComplete <> "Y" Then      'not flagged as fully posted yet, see if this current one has it flagged
                    tlEventComplete(llLoopOnEvent).sLogComplete = tmLlfList(llLoopOnEvent).sLogCompleted
                    ilInclude = True
                    Exit For
                End If
            End If
            If Not ilInclude Then
                tlEventComplete(llUpper).iVefCode = tmLlfList(llLoopOnEvent).iVefCode
                tlEventComplete(llUpper).iGameNo = tmLlfList(llLoopOnEvent).iGameNo
                tlEventComplete(llUpper).lAirDate = llAirDate
                tlEventComplete(llUpper).sLogComplete = tmLlfList(llLoopOnEvent).sLogCompleted
                llUpper = llUpper + 1
                ReDim Preserve tlEventComplete(0 To llUpper) As EVENTCOMPLETE
            End If
        Next llLoopOnEvent
        
        For llLlfLoop = LBound(tmLlfList) To UBound(tmLlfList) - 1
            'test for vehicle selectivity
            ilInclude = True
            If Not RptSel!ckcAll.Value = vbChecked Then           'include all advertisers?
                ilInclude = mFilterLists(tmLlfList(llLlfLoop).iVefCode, ilInclVefCodes, ilUseVefCodes())
            End If

            If ilInclude Then           'vehicles match; only get the records that donthave replacements

                tmGrf.sGenDesc = ""         'intitalize in case library  not found
                tmLcfSrchKey.iType = tmLlfList(llLlfLoop).iGameNo
                tmLcfSrchKey.sStatus = "C"
                tmLcfSrchKey.iVefCode = tmLlfList(llLlfLoop).iVefCode
                tmLcfSrchKey.iLogDate(0) = tmLlfList(llLlfLoop).iAirDate(0)
                tmLcfSrchKey.iLogDate(1) = tmLlfList(llLlfLoop).iAirDate(1)
                tmLcfSrchKey.iSeqNo = 1
                ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    For ilLcf = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                        If tmLcf.lLvfCode(ilLcf) > 0 And (tmLcf.iTime(0, ilLcf) = tmLlfList(llLlfLoop).iLcfStartTime(0)) And (tmLcf.iTime(1, ilLcf) = tmLlfList(llLlfLoop).iLcfStartTime(1)) Then
                            tmLvfSrchKey.lCode = tmLcf.lLvfCode(ilLcf)
                            ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilRet = BTRV_ERR_NONE Then
                                tmLtfSrchKey.iCode = tmLvf.iLtfCode
                                ilRet = btrGetEqual(hmLtf, tmLtf, imLtfRecLen, tmLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                If ilRet = BTRV_ERR_NONE Then
                                    tmGrf.sGenDesc = Trim$(tmLtf.sName)
                                End If
                                Exit For
                            End If
                        End If
                    Next ilLcf
                End If

                tmLDFSrchKey1.lCode = tmLlfList(llLlfLoop).lCode
                ilRet = btrGetGreaterOrEqual(hmLdf, tmLdf, imLdfRecLen, tmLDFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

                Do While (ilRet = BTRV_ERR_NONE) And (tmLdf.lLlfCode = tmLlfList(llLlfLoop).lCode)
                    'grf fields
                    'grfgendate - generation date
                    'grfgentime - generatin time
                    'grfLong - live log activity detail rcd
                    'grfBktType - status (activity status, F=feed, S=sign on/off, P=pgm,  G=game)
                    'grfDateType - Feed status
                    'grfStartTime - start or time of activity
                    'grfCode4 - Live log actitivty header
                    'grfPer1Genl - show start(1)  or end time(2) activity
                    'grfDollars(1) = time of sign on (for orig logkeeper or replacement logkeeper)
                    'grfDollars(2) = date of sign on (for orig logkeeper or replacement logkeeper)
                    'grfcode2 = game #
                    'grfvefcode = vehicle code
                    'grfStartDate - air date
                    'grfMissedTime - library start time
                    'sort major to minor in crystal :  air date, vehicle, game #, start Time (of library), Entered time of logkeeper, detail time of event
                    tmGrf.lLong = tmLdf.lCode          'detail code
                    tmGrf.sBktType = tmLdf.sType        'S = signon/off, P = pgm, G = game, F = feed change
                    tmGrf.sDateType = tmLdf.sCurrentFeed
                    tmGrf.iTime(0) = tmLdf.iTime(0)
                    tmGrf.iTime(1) = tmLdf.iTime(1)

                    gPackDateLong tmLdf.lEnteredDate, tmGrf.iDate(0), tmGrf.iDate(1)        '10-9-19
                    
                    If tmLdf.sSubType = "S" Then                'start time record
                        'tmGrf.iPerGenl(1) = 1                   'flag to show start time activity
                        tmGrf.iPerGenl(0) = 1                   'flag to show start time activity
                    Else
                        'tmGrf.iPerGenl(1) = 2                   'flag to show end time activity
                        tmGrf.iPerGenl(0) = 2                   'flag to show end time activity
                    End If
                    tmGrf.iStartDate(0) = tmLlfList(llLlfLoop).iAirDate(0)
                    tmGrf.iStartDate(1) = tmLlfList(llLlfLoop).iAirDate(1)
                    tmGrf.iCode2 = tmLlfList(llLlfLoop).iGameNo
                    tmGrf.iVefCode = tmLlfList(llLlfLoop).iVefCode
                    tmGrf.iMissedTime(0) = tmLlfList(llLlfLoop).iLcfStartTime(0)        'start time of game or library
                    tmGrf.iMissedTime(1) = tmLlfList(llLlfLoop).iLcfStartTime(1)
                    'tmGrf.lDollars(1) = tmLlfList(llLlfLoop).lEnteredTime  'log keeper entered time
                    'tmGrf.lDollars(2) = tmLlfList(llLlfLoop).lEnteredDate  'log keeper entered Date
                    tmGrf.lDollars(0) = tmLlfList(llLlfLoop).lEnteredTime  'log keeper entered time
                    tmGrf.lDollars(1) = tmLlfList(llLlfLoop).lEnteredDate  'log keeper entered Date
                    tmGrf.lCode4 = tmLlfList(llLlfLoop).lCode
                    gUnpackDateLong tmLlfList(llLlfLoop).iAirDate(0), tmLlfList(llLlfLoop).iAirDate(1), llAirDate
                    'determine if this game/airdate/vehicle has been completed and posted
                    For llUpper = LBound(tlEventComplete) To UBound(tlEventComplete) - 1
                        tmGrf.iPerGenl(1) = 0           'not complete
                        If tlEventComplete(llUpper).lAirDate = llAirDate And tlEventComplete(llUpper).iVefCode = tmGrf.iVefCode And tlEventComplete(llUpper).iGameNo = tmGrf.iCode2 Then
                            If tlEventComplete(llUpper).sLogComplete = "Y" Then
                                tmGrf.iPerGenl(1) = 1           'flag as complete
                                Exit For
                            End If
                        End If
                    Next llUpper
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

                    ilRet = btrGetNext(hmLdf, tmLdf, imLdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

            End If
        Next llLlfLoop
    End If

    Erase tmLlfList, ilUseVefCodes
    ilRet = btrClose(hmLlf)
    ilRet = btrClose(hmLvf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmLtf)
    ilRet = btrClose(hmLdf)
    btrDestroy hmLlf
    btrDestroy hmLvf
    btrDestroy hmLcf
    btrDestroy hmGrf
    btrDestroy hmLtf
    btrDestroy hmLdf

    Exit Sub
gGenLiveLogErr:
    On Error GoTo 0
    Resume Next
End Sub

'           Generate Copy Inventory report from CIF
'           <return>  true if OK, else false
'
Public Function gGenCopyInv() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slLowValue                    slHiValue                                               *
'******************************************************************************************
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE
    Dim ilListIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim tlCharTypeBuff As POPCHARTYPE
    Dim slStr As String
    Dim illoop As Integer
    Dim slCode As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim blSent As Boolean
    Dim blNotSent As Boolean
    Dim blProduced As Boolean
    Dim blHeld As Boolean

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        gGenCopyInv = False
        Exit Function
    End If
    imVefRecLen = Len(tmVef)

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmVef)
        btrDestroy hmCif
        btrDestroy hmVef
        gGenCopyInv = False
        Exit Function
    End If
    imCifRecLen = Len(tmCif)

    hmCpr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpr
        btrDestroy hmCif
        btrDestroy hmVef
        gGenCopyInv = False
        Exit Function
    End If
    imCprRecLen = Len(tmTCpr)

    ReDim ilcifList(0 To 0) As Integer

    ilListIndex = RptSel!lbcRptType.ListIndex       'determine which inventory report
    tmTCpr.iGenDate(0) = igNowDate(0)                'prepass filter
    tmTCpr.iGenDate(1) = igNowDate(1)
    tmTCpr.lGenTime = lgNowTime

    If ilListIndex = COPY_INVBYSTARTDATE Or ilListIndex = COPY_INVPRODUCER Then         '6-10-13 option to selective media codes on Producer report
        ilUpper = 0
        For illoop = 0 To RptSel!lbcSelection(10).ListCount - 1 Step 1
            If RptSel!lbcSelection(10).Selected(illoop) Then
                slStr = tgMcfCode(illoop).sKey
                ilRet = gParseItem(slStr, 2, "\", slCode)
                ilcifList(ilUpper) = Val(slCode)
                ReDim Preserve ilcifList(0 To ilUpper + 1) As Integer
                ilUpper = ilUpper + 1
            End If
        Next illoop
    End If
    
    blSent = True
    blNotSent = True
    blProduced = True
    blHeld = True
    If ilListIndex = COPY_INVPRODUCER Then              '4-10-13
        If RptSel!ckcSelC3(0).Value = vbUnchecked Then
            blNotSent = False
        End If
        If RptSel!ckcSelC3(1).Value = vbUnchecked Then
            blSent = False
        End If
        If RptSel!ckcSelC3(2).Value = vbUnchecked Then
            blProduced = False
        End If
        If RptSel!ckcSelC3(3).Value = vbUnchecked Then
            blHeld = False
        End If
'        If Trim$(RptSel!edcSelCFrom) = "" Then
'            RptSel!edcSelCFrom = "01/01/1970"
'        End If
'        If Trim$(RptSel!edcSelCTo) = "" Then
'            RptSel!edcSelCTo = "12/31/2069"
'        End If
'       8-22-19 use csi calendar control vs text box
        If Trim$(RptSel!CSI_CalFrom.Text) = "" Then
            RptSel!CSI_CalFrom.Text = "01/01/1970"
        End If
        If Trim$(RptSel!CSI_CalTo.Text) = "" Then
            RptSel!CSI_CalTo.Text = "12/31/2069"
        End If
        
        '7-9-13 Get the latest status changes from vCreative
        If (Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT And RptSel!ckcSelC7.Value = vbChecked Then
            ilRet = gGetVCreativeCompCopy(True)
            If ilRet = 1 Then
                gMsgBox "Get vCreative Authorization Failed." & vbCrLf & "No New Completed Copy Will be Imported." & vbCrLf & "Please Refer to VCreativeErrors.txt", vbOK
            End If
            If ilRet = 2 Then
                gMsgBox "vCreative Get Newly Completed Copy Failed." & vbCrLf & "No Newly Completed Copy Will be Imported." & vbCrLf & "Please Refer to VCreativeErrors.txt", vbOK
            End If
            If ilRet = 3 Then
                gMsgBox "Could not Interpret vCreative Return." & vbCrLf & "No New Completed Copy Will be Imported." & vbCrLf & "Please Refer to VCreativeErrors.txt", vbOK
            End If
        End If
    End If

    If ilListIndex = COPY_INVBYNUMBER Or ilListIndex = COPY_INVBYISCI Then

    ElseIf ilListIndex = COPY_INVBYSTARTDATE Or ilListIndex = COPY_INVBYEXPDATE Or ilListIndex = COPY_INVBYPURGE Or ilListIndex = COPY_INVBYENTRYDATE Or ilListIndex = COPY_INVPRODUCER Then
'        slStartDate = Trim$(RptSel!edcSelCFrom)
'        slEndDate = Trim$(RptSel!edcSelCTo)
        slStartDate = Trim$(RptSel!CSI_CalFrom.Text)
        slEndDate = Trim$(RptSel!CSI_CalTo.Text)

    ElseIf ilListIndex = COPY_INVUNAPPROVED Then
    
    Else                                                'advt
    End If
    btrExtClear hmCif   'Clear any previous extend operation
    ilExtLen = Len(tmCif)  'Extract operation record size
    imCifRecLen = Len(tmCif)

    ilRet = btrGetFirst(hmCif, tmCif, imCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmCif, llNoRec, -1, "UC", "Cif", "") '"EG") 'Set extract limits (all records)

        If ilListIndex = COPY_INVBYNUMBER Then
        ElseIf ilListIndex = COPY_INVBYISCI Then
        ElseIf ilListIndex = COPY_INVBYADVT Then
        ElseIf ilListIndex = COPY_INVBYSTARTDATE Or ilListIndex = COPY_INVBYEXPDATE Or ilListIndex = COPY_INVPRODUCER Then      '4-10-13
            tlCharTypeBuff.sType = "H"
            ilOffSet = gFieldOffset("Cif", "CifPurged")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

            tlCharTypeBuff.sType = "P"
            ilOffSet = gFieldOffset("Cif", "CifPurged")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            If slStartDate = "" Then            'no start date entered, end date mandatory
                If ilListIndex = COPY_INVBYEXPDATE Then

                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                ElseIf ilListIndex = COPY_INVBYSTARTDATE Then
                    '4-17-07 start/end dates entered, test for carted/uncarted and date spans
                    If RptSel!rbcSelC6(0).Value = True Then         'carted
                        tlCharTypeBuff.sType = "Y"
                        ilOffSet = gFieldOffset("Cif", "CifCleared")
                        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    ElseIf RptSel!rbcSelC6(1).Value = True Then     'uncarted
                        tlCharTypeBuff.sType = "Y"
                        ilOffSet = gFieldOffset("Cif", "CifCleared")
                        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If

                    gPackDate "1/1/1970", tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

                    'gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    'ilOffset = gFieldOffset("Cif", "CifRotEndDate")
                    'ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
              Else                'start and end date entered
                If ilListIndex = COPY_INVBYEXPDATE Then
                    
                    gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                ElseIf ilListIndex = COPY_INVPRODUCER Then              'test rotation dates along with anything that doesnt have a date at all
                
                
                    gPackDate "1/1/1970", tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_OR, tlDateTypeBuff, 4)

                    gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                Else
                    If RptSel!rbcSelC6(0).Value = True Then         'carted
                        tlCharTypeBuff.sType = "Y"
                        ilOffSet = gFieldOffset("Cif", "CifCleared")
                        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    ElseIf RptSel!rbcSelC6(1).Value = True Then     'uncarted
                        tlCharTypeBuff.sType = "Y"
                        ilOffSet = gFieldOffset("Cif", "CifCleared")
                        ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If

                    gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
                    ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
            End If
        ElseIf ilListIndex = COPY_INVBYPURGE Then
        ElseIf ilListIndex = COPY_INVBYENTRYDATE Then
        ElseIf ilListIndex = COPY_INVUNAPPROVED Then
        Else
            gGenCopyInv = False
            Exit Function
        End If

        ilRet = btrExtAddField(hmCif, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainCifErr
        gBtrvErrorMsg ilRet, "gGenCopyInv (btrExtAddField):" & "Cif.Btr", RptSel
        On Error GoTo 0
        ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainCifErr
            gBtrvErrorMsg ilRet, "gGenCopyInv (btrExtGetNextExt):" & "Cif.Btr", RptSel
            On Error GoTo 0
            ilExtLen = Len(tmCif)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = True
                '4-28-07 test the selectivity for the media code
                If (ilListIndex = COPY_INVBYSTARTDATE Or ilListIndex = COPY_INVPRODUCER) And UBound(ilcifList) > 0 Then
                    ilFound = False
                    For illoop = 0 To UBound(ilcifList) - 1
                        If tmCif.iMcfCode = ilcifList(illoop) Then
                            ilFound = True
                            Exit For
                        End If
                    Next illoop
                End If
                If ilListIndex = COPY_INVPRODUCER Then       '4-10-13
                    If ilFound = True Then                  'continue only if a valid media code was already found
                        ilFound = False
                        If blNotSent = True And tmCif.sCleared = "N" Then
                            ilFound = True
                        End If
                        If blSent = True And tmCif.sCleared = "S" Then
                            ilFound = True
                        End If
                        If blProduced = True And tmCif.sCleared = "Y" Then
                            ilFound = True
                        End If
                        If blHeld = True And tmCif.sCleared = "H" Then
                            ilFound = True
                        End If
                       ' tmTCpr.sCreative = tgUrf(0).sName           'show the user name on the Inv Producer reprt
                        tmTCpr.iReady = tmCif.iUrfCode              'user that added/changed the inventory
                     End If
                 End If
                 
                If ilFound Then
                  tmTCpr.lCntrNo = tmCif.lCode
                  ilRet = btrInsert(hmCpr, tmTCpr, imCprRecLen, INDEXKEY0)
                  On Error GoTo mObtainCifErr
                  gBtrvErrorMsg ilRet, "gGenCopyInv (btrInsert CPR):" & "Cif.Btr", RptSel
                  On Error GoTo 0
                End If
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    
    If ilListIndex = COPY_INVPRODUCER Then              '4-10-13
    '8-22-19 use cal control vs edit box
    'reset values to show nothing if the max dates were set for filtering purposes
'        If Trim$(RptSel!edcSelCFrom) = "01/01/1970" Then
'            RptSel!edcSelCFrom = ""
'        End If
'        If Trim$(RptSel!edcSelCTo) = "12/31/2069" Then
'            RptSel!edcSelCTo = ""
'        End If
        If Trim$(RptSel!CSI_CalFrom.Text) = "01/01/1970" Then
            RptSel!CSI_CalFrom.Text = ""
        End If
        If Trim$(RptSel!CSI_CalTo.Text) = "12/31/2069" Then
            RptSel!CSI_CalTo.Text = ""
        End If

    End If

    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmVef)
    btrDestroy hmCpr
    btrDestroy hmCif
    btrDestroy hmVef
    gGenCopyInv = True                  'valid generation
    Exit Function
mObtainCifErr:
    On Error GoTo 0
    MsgBox "RptCr: gGenCopyInv error", vbCritical + vbOKOnly, "Cif I/O Error"
    gGenCopyInv = False
    Exit Function
End Function

'       mFindsplitForPlayList - find all the split copy for a spot and show
'           on the PlayList by ISCI
'       <input> llSdfCode - spot internal code
'
Sub mFindSplitForPlayList(ilPLByAdvt As Integer, ilVefCode As Integer, ilAirIndex As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                                                                               *
'******************************************************************************************
    Dim ilValue As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim blMatchFound As Boolean
    Dim illoop As Integer
    Dim ilStartIndex As Integer
    '5/22/15: Handle generic copy assigned to airing vehicle
    Dim blStopRsf As Boolean
    Dim ilOffSet As Integer
    Dim ilRsf As Integer
    Dim slVefType As String
    Dim blAddRec As Boolean
    ReDim tlRsfSort(0 To 0) As RSFSORT

    ilValue = Asc(tgSpf.sUsingFeatures2)  'Option Fields in Orders/Proposals
    ilStartIndex = ilAirIndex
    If UBound(tmSATable) <= LBound(tmSATable) Then
        ilStartIndex = -1
    End If

    '9/29/16: Bypass airing rotation copy if vehicle is an airing vehicle and the rotation copy is not for that vehicle
    slVefType = ""
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        slVefType = tgMVef(ilVef).sType
    End If

    'test to see if split copy is used
    If ((ilValue And SPLITCOPY) = SPLITCOPY) Or ((ilValue And REGIONALCOPY) = REGIONALCOPY) Then
        '5/22/15: Handle generic copy assigned to airing vehicle
        'tmRsfSrchKey1.lCode = tmSdf.lCode
        'ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
        tmRsfSrchKey1.lCode = tmSdf.lCode
        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
            If (tmRsf.sType <> "B") Then
                '9/29/16: Bypass airing rotation copy if vehicle is an airing vehicle and the rotation copy is not for that vehicle
                blAddRec = True
                If slVefType = "A" Then
                    ilVef = gBinarySearchVef(tmRsf.iBVefCode)
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType = "A" Then
                            If ilVefCode <> tmRsf.iBVefCode Then
                                blAddRec = False
                            End If
                        End If
                    End If
                End If
                If blAddRec Then
                    ilRsf = UBound(tlRsfSort)
                    tlRsfSort(ilRsf).iRotNo = tmRsf.iRotNo
                    tlRsfSort(ilRsf).tRsf = tmRsf
                    ReDim Preserve tlRsfSort(0 To ilRsf + 1) As RSFSORT
                End If
            End If
            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If UBound(tlRsfSort) > 1 Then
            ArraySortTyp fnAV(tlRsfSort(), 0), UBound(tlRsfSort), 1, LenB(tlRsfSort(0)), 0, -1, 0
        End If
        blStopRsf = False
        For ilRsf = 0 To UBound(tlRsfSort) - 1 Step 1
            tmRsf = tlRsfSort(ilRsf).tRsf
            
            '7/25/13: Vehicle test missing.
            '         A = Rotation defined for Airing vehicle without split copy.
            '         R = Split Copy
            '         B = Blackout copy via Log blackout
            'If (tmRsf.sType <> "B") And (tmRsf.sType <> "A") Then
            If (tmRsf.sType <> "B") Then
                ilVef = gBinarySearchVef(tmRsf.iBVefCode)
                If ilVef <> -1 Then
                    blMatchFound = False
                    'If tgMVef(ilVef).sType = "A" Or tgMVef(ilVef).sType = "C" Then
                    If tgMVef(ilVef).sType <> "S" Then
                        
                        '5/22/15: Handle generic copy assigned to airing vehicle
                        If (tmRsf.lRafCode = 0) And (tmRsf.iBVefCode = ilVefCode) And (tgMVef(ilVef).sType = "A") Then
                            blStopRsf = True
                        End If
                        
                        '4/4/14: Process package vehicles
                        'If tmRsf.iBVefCode = ilVefCode Then
                        '    blMatchFound = True
                        'End If
                        If tgMVef(ilVef).sType <> "P" Then
                            If tmRsf.iBVefCode = ilVefCode Then
                                blMatchFound = True
                                '4/29/14: Test if allowed split copy
                                ilVpf = gBinarySearchVpf(ilVefCode)
                                If ilVpf <> -1 Then
                                    If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                        blMatchFound = False
                                    End If
                                End If
                            End If
                        Else
                            '4/4/14: Scan the contract lines to see if any hidden line of the package match the spot vehicle
                            ReDim ilPkLineNo(0 To 0) As Integer
                            tmClfSrchKey1.lChfCode = tmRsf.lRChfCode
                            tmClfSrchKey1.iVefCode = tgMVef(ilVef).iCode
                            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            Do While tmClf.lChfCode = tmRsf.lRChfCode And tmClf.iVefCode = tgMVef(ilVef).iCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
                                If (tmClf.sType = "O") Or (tmClf.sType = "A") Then
                                    ilPkLineNo(UBound(ilPkLineNo)) = tmClf.iLine
                                    ReDim Preserve ilPkLineNo(0 To UBound(ilPkLineNo) + 1) As Integer
                                End If
                                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            tmClfSrchKey1.lChfCode = tmRsf.lRChfCode
                            tmClfSrchKey1.iVefCode = 0
                            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While tmClf.lChfCode = tmRsf.lRChfCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
                                For illoop = 0 To UBound(ilPkLineNo) - 1 Step 1
                                    If (tmClf.iPkLineNo > 0) And (tmClf.iPkLineNo = ilPkLineNo(illoop)) Then
                                        ilVef = gBinarySearchVef(tmClf.iVefCode)
                                        If ilVef <> -1 Then
                                            If tgMVef(ilVef).sType <> "S" Then
                                                If tmClf.iVefCode = ilVefCode Then
                                                    blMatchFound = True
                                                    '4/29/14: Test if allowed split copy
                                                    ilVpf = gBinarySearchVpf(ilVefCode)
                                                    If ilVpf <> -1 Then
                                                        If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                                            blMatchFound = False
                                                        End If
                                                    End If
                                                    Exit Do
                                                End If
                                            Else
                                                ilAirIndex = ilStartIndex
                                                Do While ilAirIndex >= 0
                                                    If tmSATable(ilAirIndex).iSellCode = tmClf.iVefCode Then
                                                        If tmSATable(ilAirIndex).iAirCode = ilVefCode Then
                                                            blMatchFound = True
                                                            Exit Do
                                                        End If
                                                    End If
                                                    ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                                                Loop
                                                If blMatchFound Then
                                                    '4/21/14: Test if allowed split copy
                                                    ilVpf = gBinarySearchVpf(ilVefCode)
                                                    If ilVpf <> -1 Then
                                                        If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                                            blMatchFound = False
                                                        End If
                                                    End If
                                                    Exit Do
                                                End If
                                            End If
                                        End If
                                    End If
                                Next illoop
                                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                    Else                            'selling only
                        '4/4/14
                        ilAirIndex = ilStartIndex
                        Do While ilAirIndex >= 0
                            If tmSATable(ilAirIndex).iAirCode = ilVefCode Then
                                blMatchFound = True
                                Exit Do
                            End If
                            ilAirIndex = tmSATable(ilAirIndex).iNextIndex
                        Loop
                        If blMatchFound Then
                            '4/23/14: Test if allowed split copy
                            ilVpf = gBinarySearchVpf(ilVefCode)
                            If ilVpf <> -1 Then
                                If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                    blMatchFound = False
                                End If
                            End If
                        End If
                    End If
                    If blMatchFound Then
                        mGetCIFInfoForPlayList tmRsf.lCopyCode, ilPLByAdvt, ilVefCode, tmRsf.lRafCode
                    End If
                End If
            End If
        '5/22/15: Handle generic copy assigned to airing vehicle
        '    ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
                                        
            If blStopRsf Then
                Exit For
            End If
        Next ilRsf
    End If
    ilAirIndex = ilStartIndex
End Sub

'           mGetCIFInfoForPlayList - get all the info from the inventory (product, isci, creative title,
'               cart #) to place in array to prepare for writing to prepass file (CPR)
'           <input> llCopyCode - split copy cif code or spot cif code
'                    ilPLByAdvt - 0 = by advt, otherwise by ISCI or vehicle
'                    ilVefCode = vehicle code
'                    llRafCode = region area code for split copy, else 0
'
'       Playlist by ISCI, Playlist by Advt, Playlist by Vehicle (using carts:L09, not using carts, use reel #: L17)
Public Sub mGetCIFInfoForPlayList(llCopyCode As Long, ilPLByAdvt As Integer, ilVefCode As Integer, llRafCode As Long)
    Dim slZone As String
    Dim slISCI As String
    Dim slProduct As String
    Dim slCreative As String
    Dim ilRet As Integer
    Dim slCart As String
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilUpper As Integer
    Dim llCsfCode As Long       '3-2-10

    llCsfCode = 0
    tmCifSrchKey.lCode = llCopyCode
    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        'date the copy script into prepass so the .rpt can show it if present
        'by clearing the field if not requested, no formula needs to be sent to .rpt for the option
        If RptSel!lbcRptType.ListIndex = 11 And RptSel!ckcSelC6Add(0).Value = vbChecked Then        'playlist by ISCI
            llCsfCode = tmCif.lCsfCode          '3-2-10  copy script
        End If
        'initialize all fields incase prod/isci not defined 5/3/99
        slZone = ""
        slISCI = ""
        slProduct = ""
        slCreative = ""
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCpf.sISCI = ""
                tmCpf.sName = ""
                tmCpf.sCreative = ""
            End If
            slISCI = Trim$(tmCpf.sISCI)
            slProduct = Trim$(tmCpf.sName)
            slCreative = Trim$(tmCpf.sCreative)
        End If
        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
            If tmCif.iMcfCode <> tmMcf.iCode Then
                tmMcfSrchKey.iCode = tmCif.iMcfCode
                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmMcf.sName = ""
                End If
                slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
            Else
                slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
            End If
        Else
            'slCart = ""
            slCart = Trim$(tmCif.sReel)
        End If
        ilFound = False
        If ilPLByAdvt Then          '9-2-99
            For illoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                If (tmCpr(illoop).lCntrNo = tmSdf.lChfCode) And (tmCpr(illoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(illoop).iLen = tmSdf.iLen) Then
                    If (Trim$(tmCpr(illoop).sProduct) = slProduct) And (Trim$(tmCpr(illoop).sZone) = slZone) And (Trim$(tmCpr(illoop).sCartNo) = slCart) And (Trim$(tmCpr(illoop).sISCI) = slISCI) And (Trim$(tmCpr(illoop).sCreative) = slCreative) Then
                        tmCpr(illoop).iLineNo = tmCpr(illoop).iLineNo + 1
                        ilFound = True
                        Exit For
                    End If
                End If
            Next illoop
        Else
            For illoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                If (tmCpr(illoop).iVefCode = ilVefCode) And (tmCpr(illoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(illoop).iLen = tmSdf.iLen) Then
                    If (Trim$(tmCpr(illoop).sProduct) = slProduct) And (Trim$(tmCpr(illoop).sZone) = slZone) And (Trim$(tmCpr(illoop).sCartNo) = slCart) And (Trim$(tmCpr(illoop).sISCI) = slISCI) And (Trim$(tmCpr(illoop).sCreative) = slCreative) Then
                        tmCpr(illoop).iLineNo = tmCpr(illoop).iLineNo + 1
                        ilFound = True
                        Exit For
                    End If
                End If
            Next illoop
        End If
        If Not ilFound Then
            ilUpper = UBound(tmCpr)
            tmCpr(ilUpper).iGenDate(0) = igNowDate(0)
            tmCpr(ilUpper).iGenDate(1) = igNowDate(1)
            'tmCpr(ilUpper).iGenTime(0) = igNowTime(0)
            'tmCpr(ilUpper).iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCpr(ilUpper).lGenTime = lgNowTime
            tmCpr(ilUpper).iVefCode = ilVefCode
            tmCpr(ilUpper).iAdfCode = tmSdf.iAdfCode
            tmCpr(ilUpper).lCntrNo = tmSdf.lChfCode     '9-2-99
            tmCpr(ilUpper).iLen = tmSdf.iLen
            tmCpr(ilUpper).sProduct = slProduct
            tmCpr(ilUpper).sZone = slZone
            tmCpr(ilUpper).sCartNo = slCart
            tmCpr(ilUpper).sISCI = slISCI
            tmCpr(ilUpper).sCreative = slCreative
            tmCpr(ilUpper).lHd1CefCode = llRafCode      '4-17-07
            tmCpr(ilUpper).iLineNo = 1
            tmCpr(ilUpper).lFt1CefCode = llCsfCode      '3-2-10
            tmCpr(ilUpper).lFt2CefCode = tmCif.lCode    '4-1-13
            ReDim Preserve tmCpr(0 To ilUpper + 1) As CPR
        Else
            '10-8-10 entry already exists, but if regional copy are same for generic copy, it needs to be updated
            If (llRafCode > 0 And tmCpr(illoop).lHd1CefCode = 0) Then
                tmCpr(illoop).lHd1CefCode = llRafCode
            End If
        End If
    End If
    Exit Sub
End Sub

'           Generate prepass for report that will show Advertisers that are either
'           split copy or blackout advertisers.  Select by active rotation start/
'           end dates and Advertisers.
'
Public Sub gGenSplitBlackOutRpt()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        llChfCodes                    ilUpper                   *
'*  slNameCode                    slCode                        ilLoopOnCnt               *
'*                                                                                        *
'******************************************************************************************
    Dim ilRet As Integer
    Dim llActiveStart As Long
    Dim llActiveEnd As Long
    Dim slActiveStart As String
    Dim slActiveEnd As String
    Dim ilActiveStart(0 To 1) As Integer
    Dim ilActiveEnd(0 To 1) As Integer
    Dim ilInclVefCodes As Integer            'true = include codes stored in ilusecode array,
                                                'false = exclude codes store din ilusecode array
    ReDim ilUseVefCodes(0 To 0) As Integer       'valid  vehicles codes to process--
    Dim ilInclAdfCodes As Integer            'true = include codes stored in ilusecode array,
                                                'false = exclude codes store din ilusecode array
    ReDim ilUseAdfCodes(0 To 0) As Integer       'valid  advertiser codes to process--
    ReDim tlCrf(0 To 0) As CRF
    Dim llRots As Long
    Dim ilInclude As Integer
    Dim ilFoundVef As Integer
    Dim ilIncludeSC As Integer      'include split copy
    Dim ilIncludeBO As Integer      'include blackout
    Dim ilIncludeDormant As Integer 'include Dormant rotations
    Dim llCvfCode As Long
    Dim blFoundVef As Boolean
    Dim illoop As Integer
    Dim slVehicleNameString As String
    Dim ilStartPos As Integer
    Dim ilLen As Integer
    Dim ilMaxFieldLen As Integer
    Dim ilVefInx As Integer


    'Open btrieve files
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: VEF.BTR)", RptSel
    On Error GoTo 0

    imAdfRecLen = Len(tmAdf)    'Save Advertiser record length
    hmAdf = CBtrvTable(ONEHANDLE)          'Save ADF handle
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: AdF.BTR)", RptSel
    On Error GoTo 0

    imCprRecLen = Len(tmTCpr)    'Save prepass CPR record length
    hmCpr = CBtrvTable(ONEHANDLE)          'Save CPR handle
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: CPR.BTR)", RptSel
    On Error GoTo 0

    imCrfRecLen = Len(tmCrf)    'Save Copy Rotation header record length
    hmCrf = CBtrvTable(ONEHANDLE)          'Save CRF handle
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: CRF.BTR)", RptSel
    On Error GoTo 0
    
    imCvfRecLen = Len(tmCvf)
    hmCvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: CVF.BTR)", RptSel
    On Error GoTo 0


    imSefRecLen = Len(tmSef)    'Save SplitEntry record length
    hmSef = CBtrvTable(ONEHANDLE)          'Save SEF handle
    ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: SEF.BTR)", RptSel
    On Error GoTo 0
    
    imTxrRecLen = Len(tmTxr)
    hmTxr = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenSplitBlackOutRptErr
    gBtrvErrorMsg ilRet, "gGenSplitBlackOutRpt (btrOpen: TXR.BTR)", RptSel
    On Error GoTo 0

    'get the vehicle codes selected
    'build array of vehicles to include or exclude
    gObtainCodesForMultipleLists 6, tgVehicle(), ilInclVefCodes, ilUseVefCodes(), RptSel
    'get the Advt codes selected
    gObtainCodesForMultipleLists 0, tgAdvertiser(), ilInclAdfCodes, ilUseAdfCodes(), RptSel

    ilIncludeSC = False
    ilIncludeBO = False
    If RptSel!rbcSelCSelect(2).Value Then       'include split copy and blackouts
        ilIncludeSC = True
        ilIncludeBO = True
    ElseIf RptSel!rbcSelCSelect(0).Value Then      'include split copy only
        ilIncludeSC = True
    Else                                    'include blackouts only
        ilIncludeBO = True
    End If
    '8-22-19 use csi calendar control vs edit box
'    slActiveStart = RptSel!edcSelCFrom.Text      'rotation start date filter: determine rotations to gather
    slActiveStart = RptSel!CSI_CalFrom.Text      'rotation start date filter: determine rotations to gather
    If slActiveStart = "" Then
        slActiveStart = "1/1/1970"
    End If
    llActiveStart = gDateValue(slActiveStart)
    gPackDate slActiveStart, ilActiveStart(0), ilActiveStart(1)

'    slActiveEnd = RptSel!edcSelCTo.Text      'rotation end date filter: determine rotations to gather
    slActiveEnd = RptSel!CSI_CalTo.Text      'rotation end date filter: determine rotations to gather
    If slActiveEnd = "" Then
        slActiveEnd = "12/31/2069"
    End If
    llActiveEnd = gDateValue(slActiveEnd)
    gPackDate slActiveEnd, ilActiveEnd(0), ilActiveEnd(1)

    ilIncludeDormant = gSetCheck(RptSel!ckcSelC7.Value)

    ilRet = gObtainCrfByRegionDateSpan(RptSel, hmCrf, tlCrf(), slActiveStart, slActiveEnd)     'gather the active rotation headers,
                                                                        'if no date, get them all
    For llRots = LBound(tlCrf) To UBound(tlCrf) - 1
        'test for blackout only, split copy only
        '8/15/16: Fixed parathesis
        'If ((ilIncludeBO) And (tlCrf(llRots).iBkoutInstAdfCode > 0)) Or ((ilIncludeSC) And ((tlCrf(llRots).iBkoutInstAdfCode = 0) And (tlCrf(llRots).lRafCode > 0) And (ilIncludeDormant = True And tlCrf(llRots).sState = "D") Or (tlCrf(llRots).sState = "A"))) Then
        '7-8-19 test for PoolRotation type (crfbkoutinstadfcode = -1234)
        If ((((ilIncludeBO) And (tlCrf(llRots).iBkoutInstAdfCode > 0)) Or ((ilIncludeBO) And (tlCrf(llRots).iBkoutInstAdfCode = POOLROTATION))) Or ((ilIncludeSC) And (tlCrf(llRots).iBkoutInstAdfCode = 0) And (tlCrf(llRots).lRafCode > 0))) And (((ilIncludeDormant = True) And (tlCrf(llRots).sState = "D")) Or (tlCrf(llRots).sState = "A")) Then
            'filter the advertisers and contracts & vehicles
            ilInclude = False
            ilInclude = mFilterLists(tlCrf(llRots).iAdfCode, ilInclAdfCodes, ilUseAdfCodes())
            If ilInclude Then
                If tlCrf(llRots).iVefCode = 0 Then
                    '2-2-15 determine of subrecords need to be created for the vehicle list when more than 1 vehicle is associated with a rotation pattern
                    'first determine if rotation entered for more than 1 vehicle, and if filtering of vehicles by user before writing records for subreport
                    slVehicleNameString = ""

                    llCvfCode = tlCrf(llRots).lCvfCode
                    tmCvfSrchKey0.lCode = tlCrf(llRots).lCvfCode
                    ilRet = btrGetGreaterOrEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    tmTxr.lSeqNo = 0
                    tmTxr.iGenDate(0) = igNowDate(0)
                    tmTxr.iGenDate(1) = igNowDate(1)
                    tmTxr.lGenTime = lgNowTime
                    Do While tmCvf.lCode = llCvfCode And ilRet = BTRV_ERR_NONE
                        tmTxr.lCsfCode = tmCvf.lCrfCode         'rotation code
                        tmTxr.lSeqNo = tmTxr.lSeqNo + 1
                        For illoop = 0 To 99
                            blFoundVef = False
                            If tmCvf.iVefCode(illoop) > 0 Then
                                blFoundVef = mFilterLists(tmCvf.iVefCode(illoop), ilInclVefCodes, ilUseVefCodes())
                                If Trim$(slVehicleNameString) = "" Then
                                    ilVefInx = gBinarySearchVef(tmCvf.iVefCode(illoop))
                                    If ilVefInx = -1 Then
                                        tmVef.sName = ""
                                    End If
                                    slVehicleNameString = Trim$(tgMVef(ilVefInx).sName)
                                Else
                                    ilVefInx = gBinarySearchVef(tmCvf.iVefCode(illoop))
                                    If ilVefInx = -1 Then
                                        tmVef.sName = ""
                                    End If
                                    slVehicleNameString = slVehicleNameString & "," & Trim$(tgMVef(ilVefInx).sName)
                                End If
                            Else
                                Exit For
                            End If
                        Next illoop
                        llCvfCode = tmCvf.lLkCvfCode
                        If llCvfCode > 0 Then
                            tmCvfSrchKey0.lCode = tmCvf.lLkCvfCode
                            ilRet = btrGetEqual(hmCvf, tmCvf, imCvfRecLen, tmCvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    '2-2-15  If any multi-vehicles defined for the rotation pattern, create records in txr for subreport to print them
                    If Trim$(slVehicleNameString) <> "" Then
                        ilStartPos = 1
                        ilLen = Len(Trim$(slVehicleNameString))
                        Do While ilLen > 0
                            ilMaxFieldLen = 200
                            If ilLen < 200 Then
                                ilMaxFieldLen = ilLen
                            End If
                            
                            tmTxr.sText = Mid$(slVehicleNameString, ilStartPos, ilMaxFieldLen)
                            ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                Exit Do
                            Else
                                ilLen = ilLen - 200             '200 is max field length that has been written
                                ilStartPos = ilStartPos + ilMaxFieldLen
                            End If
                        Loop
                        
                        mCreateSplitBlackOutRecd tlCrf(llRots)
                    End If
                Else
                    'check for valid vehicles
                    ilFoundVef = False
                    ilFoundVef = mFilterLists(tlCrf(llRots).iVefCode, ilInclVefCodes, ilUseVefCodes())
                    If ilFoundVef Then
                        mCreateSplitBlackOutRecd tlCrf(llRots)
                    End If
                End If
            End If
        End If
    Next llRots

    Erase ilUseVefCodes, ilUseAdfCodes
     'close all files
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmCpr)
    btrDestroy hmCpr
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmSef)
    btrDestroy hmSef
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmTxr)
    btrDestroy hmTxr
    Exit Sub

gGenSplitBlackOutRptErr:
    On Error GoTo 0
    Resume Next
End Sub

'               Create prepass records for the Blackout / Split Copy Rotation report
'
Public Sub mCreateSplitBlackOutRecd(tlCrf As CRF)
    Dim ilRet As Integer

    tmTCpr.iGenDate(0) = igNowDate(0)
    tmTCpr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmTCpr.lGenTime = lgNowTime
    tmTCpr.iLineNo = 0           'seq #

    tmTCpr.lFt1CefCode = tlCrf.lCode     'rotation header code
    tmSefSrchKey1.lRafCode = tlCrf.lRafCode
    tmSefSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hmSef, tmSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While tmSef.lRafCode = tlCrf.lRafCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
        'file design changed, handle orig design where the Incl/Excl code is stored in RAF
        'If tmSef.sInclExcl = "I" Or tmSef.sInclExcl = "E" Then
            tmTCpr.iLineNo = tmTCpr.iLineNo + 1
            tmTCpr.lFt2CefCode = tmSef.lCode         'split entry code
            ilRet = btrInsert(hmCpr, tmTCpr, imCprRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
        'End If
        ilRet = btrGetNext(hmSef, tmSef, imSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
End Sub

'           Generate Copy Script Affidavits report from CIF
'           <return>  true if OK, else false
'
Public Function gGenCopyScriptAffs() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slLowValue                    slHiValue                                               *
'******************************************************************************************
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE
    Dim ilListIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim tlCharTypeBuff As POPCHARTYPE
    Dim slStr As String
    Dim illoop As Integer
    Dim slCode As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPLCODE

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmVef)
        btrDestroy hmCif
        btrDestroy hmVef
        gGenCopyScriptAffs = False
        Exit Function
    End If
    imCifRecLen = Len(tmCif)

    hmCpr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpr
        btrDestroy hmCif
        btrDestroy hmVef
        gGenCopyScriptAffs = False
        Exit Function
    End If
    imCprRecLen = Len(tmTCpr)

    ReDim ilcifList(0 To 0) As Integer

    ilListIndex = RptSel!lbcRptType.ListIndex       'determine which inventory report
    tmTCpr.iGenDate(0) = igNowDate(0)                'prepass filter
    tmTCpr.iGenDate(1) = igNowDate(1)
    tmTCpr.lGenTime = lgNowTime

    ilUpper = 0
    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
        If RptSel!lbcSelection(0).Selected(illoop) Then
            slStr = tgAdvertiser(illoop).sKey
            ilRet = gParseItem(slStr, 2, "\", slCode)
            ilcifList(ilUpper) = Val(slCode)
            ReDim Preserve ilcifList(0 To ilUpper + 1) As Integer
            ilUpper = ilUpper + 1
        End If
    Next illoop

'   8-22-19 use csi calendar control vs text box
'    slStartDate = Trim$(RptSel!edcSelCFrom)
'    slEndDate = Trim$(RptSel!edcSelCTo)
    slStartDate = Trim$(RptSel!CSI_CalFrom.Text)
    slEndDate = Trim$(RptSel!CSI_CalTo.Text)

    btrExtClear hmCif   'Clear any previous extend operation
    ilExtLen = Len(tmCif)  'Extract operation record size
    imCifRecLen = Len(tmCif)

    ilRet = btrGetFirst(hmCif, tmCif, imCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmCif, llNoRec, -1, "UC", "Cif", "") '"EG") 'Set extract limits (all records)
            'ignore Purged or History items
            tlCharTypeBuff.sType = "H"
            ilOffSet = gFieldOffset("Cif", "CifPurged")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

            tlCharTypeBuff.sType = "P"
            ilOffSet = gFieldOffset("Cif", "CifPurged")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            
            ilOffSet = gFieldOffset("Cif", "CifCsfCode")
            tlIntTypeBuff.lCode = 0
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 4)
            
            'test to see that user entered dates span the Inventory rotation start/end dates
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
            ilRet = btrExtAddLogicConst(hmCif, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
           
        ilRet = btrExtAddField(hmCif, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainCifErr
        gBtrvErrorMsg ilRet, "gGenCopyScriptAffs (btrExtAddField):" & "Cif.Btr", RptSel
        On Error GoTo 0
        ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainCifErr
            gBtrvErrorMsg ilRet, "gGenCopyScriptAffs (btrExtGetNextExt):" & "Cif.Btr", RptSel
            On Error GoTo 0
            ilExtLen = Len(tmCif)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                For illoop = 0 To UBound(ilcifList) - 1
                    If tmCif.iAdfCode = ilcifList(illoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next illoop
                If ilFound Then
                  tmTCpr.lCntrNo = tmCif.lCode

                  ilRet = btrInsert(hmCpr, tmTCpr, imCprRecLen, INDEXKEY0)
                  On Error GoTo mObtainCifErr
                  gBtrvErrorMsg ilRet, "gGenCopyScriptAffs (btrInsert CPR):" & "Cif.Btr", RptSel
                  On Error GoTo 0
                End If
                ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCif, tmCif, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If

    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmCif)
    btrDestroy hmCpr
    btrDestroy hmCif
    btrDestroy hmVef
    gGenCopyScriptAffs = True                  'valid generation
    Exit Function
mObtainCifErr:
    On Error GoTo 0
    MsgBox "RptCr: gGenCopyScriptAffs error", vbCritical + vbOKOnly, "Cif I/O Error"
    gGenCopyScriptAffs = False
    Exit Function
End Function

'                           mPlayListSellAirList - create a list of the advertisers and their vehicles associated with them.
'                            Playlist by ISCI (by option), will show the vehicles that make up the ISCI copy
'                           <input> ilAirVefCode - vehicle to process for copy (could be airing)
'                                   ilVefCode - vehicle that the spot is determined from (selling, conventional, game0
'                                   iladfcode - advertiser code of spot
'
Public Sub mPlayListSellAirList(slType As String, ilVefAirCode As Integer, ilVefCode As Integer, ilAdfCode As Integer)
    Dim ilLoopOnSA As Integer
    Dim ilFound As Integer

    ilFound = False
    For ilLoopOnSA = LBound(tmSellAirList) To UBound(tmSellAirList) - 1
        If tmSellAirList(ilLoopOnSA).iAdfCode = ilAdfCode Then
            If tmSellAirList(ilLoopOnSA).iVefCode = ilVefCode And tmSellAirList(ilLoopOnSA).iVefAirCode = ilVefAirCode Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoopOnSA
    
    If Not ilFound Then
        tmSellAirList(UBound(tmSellAirList)).iAdfCode = ilAdfCode
        tmSellAirList(UBound(tmSellAirList)).iVefCode = ilVefCode
        tmSellAirList(UBound(tmSellAirList)).iVefAirCode = ilVefAirCode
        tmSellAirList(UBound(tmSellAirList)).sType = slType
        ReDim Preserve tmSellAirList(0 To UBound(tmSellAirList) + 1) As SELLAIRLIST
    End If
    Exit Sub
End Sub

Private Function mMergeWithLog(ilVefCode As Integer) As Integer
    Dim ilVff As Integer
    
    mMergeWithLog = True
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            If tgVff(ilVff).sMergeTraffic = "S" Then
                mMergeWithLog = False
            End If
            Exit For
        End If
    Next ilVff

End Function

Private Function mCheckPackageVehicles(llChfCode As Long, ilVefCode As Integer, slStartDate As String, ilIncludeCodes As Integer, ilUseVefCodes() As Integer) As Integer
    Dim ilVef As Integer
    Dim ilAVef As Integer
    Dim ilVpf As Integer
    Dim ilFoundVef As Integer
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim tlVef As VEF
    
    If ilVefCode = 0 Then           'more than 1 vehicle associated with this rotation, and placed in special file
        mCheckPackageVehicles = True
        Exit Function
    End If
    ilFoundVef = mFilterLists(ilVefCode, ilIncludeCodes, ilUseVefCodes())
    If ilFoundVef Then
            mCheckPackageVehicles = True
        Exit Function
    End If
    If RptSel!ckcOption.Value = vbUnchecked Then
        mCheckPackageVehicles = False           'forget about pkg references, just show the rotations for vehicles selected
        Exit Function
    End If
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "P" Then
            'Find hidden lines of the package
            ReDim ilPkLineNo(0 To 0) As Integer
            tmClfSrchKey1.lChfCode = llChfCode
            tmClfSrchKey1.iVefCode = ilVefCode
            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While tmClf.lChfCode = llChfCode And tmClf.iVefCode = ilVefCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
                If (tmClf.sType = "O") Or (tmClf.sType = "A") Then
                    ilPkLineNo(UBound(ilPkLineNo)) = tmClf.iLine
                    ReDim Preserve ilPkLineNo(0 To UBound(ilPkLineNo) + 1) As Integer
                End If
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            tmClfSrchKey1.lChfCode = llChfCode
            tmClfSrchKey1.iVefCode = 0
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While tmClf.lChfCode = llChfCode And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
                For illoop = 0 To UBound(ilPkLineNo) - 1 Step 1
                    If (tmClf.iPkLineNo > 0) And (tmClf.iPkLineNo = ilPkLineNo(illoop)) Then
                        ilFoundVef = mFilterLists(tmClf.iVefCode, ilIncludeCodes, ilUseVefCodes())
                        If ilFoundVef Then
                            mCheckPackageVehicles = True
                            Exit Function
                        End If
                        ilVef = gBinarySearchVef(tmClf.iVefCode)
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType = "S" Then
                                ReDim ilAVefCode(0 To 0) As Integer
                                tlVef = tgMVef(ilVef)
                                gBuildLinkArray hmVLF, tlVef, slStartDate, ilAVefCode()
                                For ilAVef = 0 To UBound(ilAVefCode) - 1 Step 1
                                    ilFoundVef = mFilterLists(ilAVefCode(ilAVef), ilIncludeCodes, ilUseVefCodes())
                                    If ilFoundVef Then
                                        mCheckPackageVehicles = True
                                        Exit Function
                                    End If
                                Next ilAVef
                            End If
                        End If
                    End If
                Next illoop
                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
    mCheckPackageVehicles = False
End Function

'           gCreateLinksWithAvaillen - Selling to Airing or Airing to Selling list of links
'           Each link with have its associated selling avails length shown.
'           Monday - friday gathers just the Monday SSF record, Sat & Sun separate passes
'           Use SSF for speed rather than going thru the library
'
'           Outer loops makes 3 passes:  M-f, Sa & Sun
'           Loop on the Selling or Airing vehicles selected and process 1 vehicle at a time
'           Obtain links for associated selling or airing vehicle (tmVlf)
'              i.e. gObtainVlf "S", hmVlf, ilVefCode, llTZAdjDate, tmVlf()        'get the associated airing vehicles links with this selling  vehicle
'           Build array of selling vehicles (if sell to air, only the 1 vehicle is in table (tmSellList)
'                   if air to sell, loop thru the links table and get unique selling vehicles
'           Sort the array of links by selling time (array contains key (sell time) and vlf record
'           Loop on the selling vehicles created (tmSellList)
'           Get SSF for the vehicle
'           Loop thru SSF and pick up avails only
'           Find matching entry in tmVlf
'
Public Sub gCreateLinksWithAvailLen()
    Dim ilRet As Integer
    Dim ilVehicle As Integer            'loop on vehicles to process
    Dim blFound As Boolean
    Dim slNameCode As String
    Dim slName As String
    Dim ilVefCode As Integer            'vehicle code
    Dim slCode As String
    Dim ilVefIndex As Integer           ' index into vehicle array
    Dim ilIndex As Integer              'list box to use (sell or air)
    Dim ilListIndex As Integer          'selling to air or air to sell request
    Dim slWhichWay As String * 1        'S = selling option, A = airing option
    Dim llEffectiveMF As Long           'effective MF date
    Dim llEffectiveSa As Long           'effective Sa date
    Dim llEffectiveSU As Long           'effective Su date
    Dim llEffectiveDate As Long
    Dim llTemp As Long
    Dim ilTemp As Integer
    Dim ilLoopOnDay As Integer          '3 passes for each vehicle:  MF, Sa & Su
    Dim ilDayOfWeek As Integer
    Dim ilSellList() As Integer         'list of selling vehicles to process at a time:  only 1 if sell to air, list of selling if airing selection
    Dim ilLoopOnSellList As Integer
    Dim ilLoopOnVlf As Integer
    Dim ilDate(0 To 1) As Integer
    Dim ilEvt As Integer
    Dim ilLoopOnLink As Integer
    Dim llTime As Long
    Dim ilNextLink As Integer
    Dim ilLen As Integer
    Dim blFoundZeroLen As Boolean
    Dim ilAdjustMF As Integer
    Dim ilAdjustStart As Integer
    Dim ilAdjustEnd As Integer
    Dim llZeroLen() As Long
    Dim ilLoopOnZeroLen As Integer
    Dim blAnyZeroLenOnDay As Boolean
    Dim blInclMF As Boolean
    Dim blInclSa As Boolean
    Dim blInclSu As Boolean
    Dim blIgnorePass As Boolean
    Dim ilSelectedVefCode As Integer
    Dim ilAnfCode As Integer
    Dim llInputDate As Long
    Dim ilInputDayOfWeek As Integer
    Dim llLatestDate As Long
    Dim slDate As String
    ReDim ilEvtType(0 To 14) As Integer
    Dim ilType As Integer
    Dim ilWeekDay As Integer

    imVlfRecLen = Len(tmVlf)
    hmVLF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateLinksWithAvailLenErr
    gBtrvErrorMsg ilRet, "gCreateLinksWithAVaillen (btrOpen: VLF.BTR)", RptSel
    On Error GoTo 0
    
    imGrfRecLen = Len(tmGrf)
    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateLinksWithAvailLenErr
    gBtrvErrorMsg ilRet, "gCreateLinksWithAVaillen (btrOpen: GRF.BTR)", RptSel
    On Error GoTo 0
    
    imSsfRecLen = Len(tmSsf)
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateLinksWithAvailLenErr
    gBtrvErrorMsg ilRet, "gCreateLinksWithAVaillen (btrOpen: SSF.BTR)", RptSel
    On Error GoTo 0
    
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateLinksWithAvailLenErr
    gBtrvErrorMsg ilRet, "gCreateLinksWithAVaillen (btrOpen: lcf.BTR)", RptSel
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    
    'set the type of events to get fro the day (only  avails)
    For ilTemp = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilTemp) = False
    Next ilTemp
    ilEvtType(2) = True
    
    ilListIndex = RptSel!lbcRptType.ListIndex
    If ilListIndex = 0 Then         'sell to air
        ilIndex = 0                 'index for selling vehicle list box selection
        slWhichWay = "S"
    Else
        ilIndex = 1                 'index for airing vehicle list box selections
        slWhichWay = "A"
    End If
    
    blInclMF = True
    blInclSa = True
    blInclSu = True
    If RptSel!ckcSel2(0).Value = vbUnchecked Then
        blInclMF = False
    End If
    If RptSel!ckcSel2(1).Value = vbUnchecked Then
        blInclSa = False
    End If
    If RptSel!ckcSel2(2).Value = vbUnchecked Then
        blInclSu = False
    End If
    
    '2-6-14 determine the effective date for m-f (backup date entered to monday), sa (backup or increment to sa), su (increment to sun)
    'all valid transactions need to show
    '        llInputDate = gDateValue(RptSel!edcSelA.Text)
    llInputDate = gDateValue(RptSel!CSI_CalDateA.Text)      '12-11-19 change to use csi calendar control
    llTemp = llInputDate
    ilInputDayOfWeek = gWeekDayLong(llInputDate)
    ilDayOfWeek = ilInputDayOfWeek
    
    Do While ilDayOfWeek <> 0           'backup MF to monday
        llTemp = llTemp - 1
        ilDayOfWeek = gWeekDayLong(llTemp)
    Loop
    llEffectiveMF = llTemp
    
    'Sat
    llTemp = llInputDate
    ilDayOfWeek = ilInputDayOfWeek
    If ilDayOfWeek = 5 Then
        llEffectiveSa = llTemp
    Else
        If ilDayOfWeek = 6 Then     'its sunday, backup to sat
            llTemp = llTemp - 1
        Else
            Do While ilDayOfWeek <> 5           'its a m-f date, increment to sa
                llTemp = llTemp + 1
                ilDayOfWeek = gWeekDayLong(llTemp)
            Loop
        End If
        llEffectiveSa = llTemp
    End If
    
    'Sun
    llTemp = llInputDate
    ilDayOfWeek = ilInputDayOfWeek
    If ilDayOfWeek = 6 Then
        llEffectiveSU = llTemp
    Else
        Do While ilDayOfWeek <> 6
            llTemp = llTemp + 1
            ilDayOfWeek = gWeekDayLong(llTemp)
        Loop
        llEffectiveSU = llTemp
    End If
    
    'process each vehicle 3 times.  Process each vehicle for MF links, Sa links & Su links
    For ilVehicle = 0 To RptSel!lbcSelection(ilIndex).ListCount - 1 Step 1
        If (RptSel!lbcSelection(ilIndex).Selected(ilVehicle)) Then
            If ilIndex = 0 Then
                slNameCode = tgSellNameCode(ilVehicle).sKey
            Else
                slNameCode = tgAirNameCode(ilVehicle).sKey
            End If
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSelectedVefCode = Val(slCode)
            ilVefIndex = gBinarySearchVef(ilSelectedVefCode)
            
            For ilLoopOnDay = 1 To 3        'pass 1 = M-F, pass 2 = SA, pass 3 = su
                blIgnorePass = False
                If ilLoopOnDay = 1 Then     'm-f
                    llEffectiveDate = llEffectiveMF
                    'looping factor in case the first days of the week are just placeholders.  different libraries across the week
                    ilAdjustStart = 1
                    ilAdjustEnd = 5
                    If Not blInclMF Then            'dont include MF pass
                        blIgnorePass = True         'ignore the processing
                    End If
                ElseIf ilLoopOnDay = 2 Then     'sa
                    llEffectiveDate = llEffectiveSa
                    ilAdjustStart = 1
                    ilAdjustEnd = 1
                    If Not blInclSa Then            'dont includeSA pass
                        blIgnorePass = True         'ignore the processing
                    End If
                Else                            'su
                    llEffectiveDate = llEffectiveSU
                    ilAdjustStart = 1
                    ilAdjustEnd = 1
                    If Not blInclSu Then            'dont includeSU pass
                        blIgnorePass = True         'ignore the processing
                    End If
                End If
                'ReDim tmVlfSort(1 To 1) As VLFSORT
                ReDim tmVlfSort(0 To 0) As VLFSORT
                ReDim ilSellList(0 To 0) As Integer
                
                If Not blIgnorePass Then
                    mObtainVlf slWhichWay, hmVLF, ilSelectedVefCode, llEffectiveDate, tmVlfSort()        'get the associated airing vehicles links with this selling  vehicle
                    If slWhichWay = "S" Then        'sell to air, only 1 vehicle at a time
                        ilSellList(0) = ilSelectedVefCode
                        ReDim Preserve ilSellList(0 To 1) As Integer
                    Else                            'air to sell, get the list of selling vehicles associated with the airing.  List to be used to get the avail lengths from SSF
                        For ilLoopOnVlf = LBound(tmVlfSort) To UBound(tmVlfSort) - 1
                            ilTemp = tmVlfSort(ilLoopOnVlf).tVlf.iSellCode
                            blFound = False
                            For ilLoopOnSellList = 0 To UBound(ilSellList) - 1
                                If ilSellList(ilLoopOnSellList) = ilTemp Then
                                    blFound = True
                                    Exit For
                                End If
                            Next ilLoopOnSellList
                            If Not blFound Then
                                ilSellList(UBound(ilSellList)) = ilTemp
                                ReDim Preserve ilSellList(0 To UBound(ilSellList) + 1) As Integer
                            End If
                        Next ilLoopOnVlf
                    End If
                End If
                
                'list of selling vehicles are now created so that the SSF can be read to get the avail lengths
                For ilLoopOnSellList = LBound(ilSellList) To UBound(ilSellList) - 1
                    ilVefCode = ilSellList(ilLoopOnSellList)
                    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
                    ilType = 0
    
                    ReDim llZeroLen(0 To 0) As Long
                    blAnyZeroLenOnDay = False
                    For ilAdjustMF = ilAdjustStart To ilAdjustEnd       'if pass 1 (MF), may need to go thru entire week to find first day that has non-zero avail lengths
                        gPackDateLong llEffectiveDate, ilDate(0), ilDate(1)
                        slDate = Format$(llEffectiveDate, "m/d/yy")
    
                        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        tmSsfSrchKey.iType = 0 'slType
                        tmSsfSrchKey.iVefCode = ilVefCode
                        tmSsfSrchKey.iDate(0) = ilDate(0)
                        tmSsfSrchKey.iDate(1) = ilDate(1)
                        tmSsfSrchKey.iStartTime(0) = 0
                        tmSsfSrchKey.iStartTime(1) = 0
                        ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                       
                        'if no ssf found, build one from the calendar
                        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iVefCode <> ilVefCode) Or ((tmSsf.iDate(0) <> ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))) Then
                            If (llEffectiveDate + ilAdjustMF - 1 > llLatestDate) Then
                                
                                ReDim tlLLC(0 To 0) As LLC  'Merged library names
                                If tgMVef(ilVefIndex).sType <> "G" Then
                                    ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                                    If ilWeekDay = 1 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 2 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 3 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 4 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 5 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 6 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                                    ElseIf ilWeekDay = 7 Then
                                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                                    End If
                                End If
                 
                                tmSsf.iType = 0
                                tmSsf.iVefCode = ilVefCode
                                tmSsf.iDate(0) = ilDate(0)
                                tmSsf.iDate(1) = ilDate(1)
                                gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                                tmSsf.iCount = 0
                
                                For ilTemp = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                
                                    tmAvail.iRecType = Val(tlLLC(ilTemp).sType)
                                    gPackTime tlLLC(ilTemp).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                                    tmAvail.iLtfCode = tlLLC(ilTemp).iLtfCode
                                    tmAvail.iAvInfo = tlLLC(ilTemp).iAvailInfo Or tlLLC(ilTemp).iUnits
                                    tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilTemp).sLength))
                                    tmAvail.ianfCode = Val(tlLLC(ilTemp).sName)
                                    tmAvail.iNoSpotsThis = 0
                                    tmAvail.iOrigUnit = 0
                                    tmAvail.iOrigLen = 0
                                    tmSsf.iCount = tmSsf.iCount + 1
                                    tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                                Next ilTemp
                                ilRet = BTRV_ERR_NONE
                            End If
                        End If
                                                  
                       Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))
                           'loop thru the SSF records and pick up the selling avail time and length.  Find all links in tmVlfSort that have the matching sell vehicle to create
                           'a record for the associated airiing vehicle
                            ilEvt = 1
                            ilNextLink = LBound(tmVlfSort)
                            Do While ilEvt <= tmSsf.iCount
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then 'Contract Avails only
                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                    'ilNextLink = LBound(tmVlfSort)
                                    For ilLoopOnLink = ilNextLink To UBound(tmVlfSort) - 1
                                        If tmVlfSort(ilLoopOnLink).lAvailTime = llTime And tmVlfSort(ilLoopOnLink).tVlf.iSellCode = ilVefCode Then
                                            'get the avail length & create an entry to output
                                            ilLen = tmAvail.iLen
                                            ilAnfCode = tmAvail.ianfCode                '10-30-14
                                            'is this time in the list of zero length avails?
                                            blFoundZeroLen = False
                                            'If ilLen = 0 Then
                                                For ilLoopOnZeroLen = LBound(llZeroLen) To UBound(llZeroLen) - 1
                                                    If llZeroLen(ilLoopOnZeroLen) = llTime Then         'matching zero length avail time
                                                        blFoundZeroLen = True
                                                        Exit For
                                                    End If
                                                Next ilLoopOnZeroLen
    
                                            'if pass 2 or 3 (sa or sun), always continue
                                            'if pass 1: ok to process if non-zero avail length and first day of mon-fri (iladjust = 1), or if processing MF and an avail time that had zero length has been found with non-zero avai length
                                            If (ilLoopOnDay = 2 Or ilLoopOnDay = 3) Or (ilLoopOnDay = 1 And ilLen > 0 And blFoundZeroLen = True) Or (ilLoopOnDay = 1 And ilLen > 0 And blFoundZeroLen = True) Or (ilLoopOnDay = 1 And ilLen > 0 And ilAdjustMF = 1) Then
                                                tmGrf.lGenTime = lgNowTime
                                                tmGrf.iGenDate(0) = igNowDate(0)
                                                tmGrf.iGenDate(1) = igNowDate(1)
                                                tmGrf.iVefCode = ilVefCode          'selling vehicle code
                                                tmGrf.iCode2 = tmVlfSort(ilLoopOnLink).tVlf.iAirCode        'airing vehicle
                                                'tmGrf.iPerGenl(1) = tmVlfSort(ilLoopOnLink).tVlf.iSellDay           'selling vehicle day of week (0=mf , 6 = sa, 7 =su)
                                                'tmGrf.iPerGenl(2) = tmVlfSort(ilLoopOnLink).tVlf.iAirDay           'selling vehicle day of week (0=mf , 6 = sa, 7 =su)
                                                tmGrf.iPerGenl(0) = tmVlfSort(ilLoopOnLink).tVlf.iSellDay           'selling vehicle day of week (0=mf , 6 = sa, 7 =su)
                                                tmGrf.iPerGenl(1) = tmVlfSort(ilLoopOnLink).tVlf.iAirDay           'selling vehicle day of week (0=mf , 6 = sa, 7 =su)
                                                tmGrf.iTime(0) = tmVlfSort(ilLoopOnLink).tVlf.iSellTime(0)      'sell avail time
                                                tmGrf.iTime(1) = tmVlfSort(ilLoopOnLink).tVlf.iSellTime(1)
                                                tmGrf.iMissedTime(0) = tmVlfSort(ilLoopOnLink).tVlf.iAirTime(0)      'air time
                                                tmGrf.iMissedTime(1) = tmVlfSort(ilLoopOnLink).tVlf.iAirTime(1)
                                                tmGrf.iStartDate(0) = tmVlfSort(ilLoopOnLink).tVlf.iEffDate(0)      'effective date
                                                tmGrf.iStartDate(1) = tmVlfSort(ilLoopOnLink).tVlf.iEffDate(1)      'effective date
                                                tmGrf.iDate(0) = tmVlfSort(ilLoopOnLink).tVlf.iTermDate(0)          'end/termination date
                                                tmGrf.iDate(1) = tmVlfSort(ilLoopOnLink).tVlf.iTermDate(1)
                                                tmGrf.sBktType = tmVlfSort(ilLoopOnLink).tVlf.sStatus
                                                tmGrf.lCode4 = ilLen
                                                tmGrf.iRdfCode = ilAnfCode              '10-30-14
                                                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                                On Error GoTo gCreateLinksWithAvailLenErr
                                                gBtrvErrorMsg ilRet, "gCreateLinksWithAvailLen (btrInsert: GRF.BTR)", RptSel
                                                On Error GoTo 0
                                            Else
                                                'see if this avail time has been saved already, keep track only for mf links
                                                If ilLoopOnDay = 1 And ilLen = 0 Then
                                                    blAnyZeroLenOnDay = True
                                                    blFoundZeroLen = False
                                                    For ilLoopOnZeroLen = LBound(llZeroLen) To UBound(llZeroLen) - 1
                                                        If llZeroLen(ilLoopOnZeroLen) = llTime Then
                                                            blFoundZeroLen = True
                                                            Exit For
                                                        End If
                                                    Next ilLoopOnZeroLen
                                                    If Not blFoundZeroLen Then
                                                        llZeroLen(ilLoopOnZeroLen) = llTime
                                                        ReDim Preserve llZeroLen(0 To UBound(llZeroLen) + 1)
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If tmVlfSort(ilLoopOnLink).lAvailTime > llTime Then      'if this time is past the time just finished, stop and save the current index.
                                                                                                    'use that index as a starting point in table
                                                ilNextLink = ilLoopOnLink
                                                Exit For
                                            End If
                                        End If
                                        
                                    Next ilLoopOnLink
                                    ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                                End If
                                ilEvt = ilEvt + 1                                       'get next event
                            Loop                                            'Do While ilEvt <= tmSsf.iCount
                            'read next ssf to see if theres a continuation of same day and vehicle
                            imSsfRecLen = Len(tmSsf)
                            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop            'go see if the next ssf is matching day and vehicle
                        If ilLoopOnDay = 1 Then     'if on pass 1 (MF), was there a non-zero avail found.  If so, no need to search any further m-f days
                            If UBound(llZeroLen) = 0 Then       'no zero lengths, continue to next pass (sat)
                                Exit For
                            Else
    '                                    'increment the day and do next day only if more non zero times remain
    '                                    blFoundZeroLen = False
    '                                    For ilLoopOnZeroLen = LBound(llZeroLen) To UBound(llZeroLen) - 1
    '                                        If llZeroLen(ilLoopOnZeroLen) >= 0 Then     'found avail time that is zero len
    '                                            blFoundZeroLen = True
    '                                            Exit For
    '                                        End If
    '                                    Next ilLoopOnZeroLen
    
                                If blAnyZeroLenOnDay Then
                                    llEffectiveDate = llEffectiveDate + ilAdjustMF
                                    blAnyZeroLenOnDay = False           'reset for next day, might have all avails with length on next day
                                Else
                                    Exit For
                                End If
                            End If
                        End If                  'If ilLoopOnDay = 1
                    Next ilAdjustMF             'loop on m-f if pass 1 (only looping on multiple days if 0 avail lengths found
                Next ilLoopOnSellList           'process next selling vehicle (could be multiple selling vehicles if airing to selling option was selected)
                                                'if sell to air selected, will only be one in the table
            Next ilLoopOnDay                    'process next pass for MF, Sa, Su
        End If
    Next ilVehicle                              'next sell or air vehicle in selection list
    
    Erase ilSellList
    Erase tmVlfSort
    Erase llZeroLen
    Erase tlLLC
    
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmGrf)
    btrDestroy hmSsf
    btrDestroy hmVLF
    btrDestroy hmLcf
    btrDestroy hmGrf
    Exit Sub
    
gCreateLinksWithAvailLenErr:
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainVlf                      *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicle link         *
'*                     records for date specified      *
'*
'*  This a copy of gObtainVlf.  A sort key is appended *
'   to each link so it can be sorted by selling time
'*******************************************************
Sub mObtainVlf(slType As String, hlVlf As Integer, ilVefCode As Integer, llDate As Long, tlVlfSort() As VLFSORT)
'
'   mObtainVcf llDate
'   Where:
'       slType(I) "S" = Selling; "A" = Airing
'       hlVlf(I)- Vcf handle
'       ilVefCode(I)- Vehicle code
'       llDate(I)- Date within week to obtain Vcf records
'       tlVlfSort(O)- Array of VLF records
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim llEffDate As Long
    Dim llTermDate As Long
    Dim ilDay As Integer
    Dim ilUpperBound As Integer
    Dim ilVlfRecLen As Integer
    Dim ilVlfDefined As Integer
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilTerminated As Integer
    Dim tlSrchKey0 As VLFKEY0
    Dim tlSrchKey1 As VLFKEY1
    Dim llTime As Long
    Dim slStr As String
    'Convert to Monday date so tlVlfSort for monday can only have an Monday date
    'Convert to Saturday date so tlVlfSort for saturday can only have an Monday date
    'Convert to Sunday date so tlVlfSort for sunday can only have an Monday date
    slDate = Format$(llDate, "m/d/yy")
    gPackDate slDate, ilEffDate0, ilEffDate1
    ilVlfDefined = False
    ilDay = gWeekDayLong(llDate)
    If ilDay = 6 Then
        ilDay = 7
    ElseIf ilDay = 5 Then   'Saturady
        ilDay = 6
    Else
        ilDay = 0
    End If
'    If UBound(tlVlfSort) > 1 Then
'        If slType = "S" Then
'            If (tlVlfSort(1).tVlf.iSellCode = ilVefCode) And (tlVlfSort(1).tVlf.iSellDay = ilDay) Then
'                gUnpackDate tlVlfSort(1).tVlf.iEffDate(0), tlVlfSort(1).tVlf.iEffDate(1), slDate
'                llEffDate = gDateValue(slDate)
'                gUnpackDate tlVlfSort(1).tVlf.iTermDate(0), tlVlfSort(1).tVlf.iTermDate(1), slDate
'                If slDate = "" Then
'                    slDate = "12/31/2060"
'                End If
'                llTermDate = gDateValue(slDate)
'                If (llDate >= llEffDate) And (llDate <= llTermDate) Then
'                    ilVlfDefined = True
'                End If
'            End If
'        Else
'            If (tlVlfSort(1).tVlf.iAirCode = ilVefCode) And (tlVlfSort(1).tVlf.iAirDay = ilDay) Then
'                gUnpackDate tlVlfSort(1).tVlf.iEffDate(0), tlVlfSort(1).tVlf.iEffDate(1), slDate
'                llEffDate = gDateValue(slDate)
'                gUnpackDate tlVlfSort(1).tVlf.iTermDate(0), tlVlfSort(1).tVlf.iTermDate(1), slDate
'                If slDate = "" Then
'                    slDate = "12/31/2060"
'                End If
'                llTermDate = gDateValue(slDate)
'                If (llDate >= llEffDate) And (llDate <= llTermDate) Then
'                    ilVlfDefined = True
'                End If
'            End If
'        End If
'    End If
    If Not ilVlfDefined Then
        'ReDim tlVlfSort(1 To 1) As VLFSORT
        ReDim tlVlfSort(0 To 0) As VLFSORT
        ilUpperBound = UBound(tlVlfSort)
        ilVlfRecLen = Len(tmVlf)
        'Determine effective date
        If slType = "S" Then
            tlSrchKey0.iSellCode = ilVefCode
            tlSrchKey0.iSellDay = ilDay
            tlSrchKey0.iEffDate(0) = ilEffDate0
            tlSrchKey0.iEffDate(1) = ilEffDate1
            ilEffDate0 = 0
            ilEffDate1 = 0
            tlSrchKey0.iSellTime(0) = 0
            tlSrchKey0.iSellTime(1) = 6144  '24*256
            tlSrchKey0.iSellPosNo = 32000
            ilRet = btrGetLessOrEqual(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlfSort(ilUpperBound).tVlf.iSellCode = ilVefCode)
                ilTerminated = False
                'Check for CBS
                If (tlVlfSort(ilUpperBound).tVlf.iTermDate(1) <> 0) Or (tlVlfSort(ilUpperBound).tVlf.iTermDate(0) <> 0) Then
                    If (tlVlfSort(ilUpperBound).tVlf.iTermDate(1) < tlVlfSort(ilUpperBound).tVlf.iEffDate(1)) Or ((tlVlfSort(ilUpperBound).tVlf.iEffDate(1) = tlVlfSort(ilUpperBound).tVlf.iTermDate(1)) And (tlVlfSort(ilUpperBound).tVlf.iTermDate(0) < tlVlfSort(ilUpperBound).tVlf.iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (tlVlfSort(ilUpperBound).tVlf.sStatus <> "P") And (tlVlfSort(ilUpperBound).tVlf.iSellDay = ilDay) And (Not ilTerminated) Then
                    ilEffDate0 = tlVlfSort(ilUpperBound).tVlf.iEffDate(0)
                    ilEffDate1 = tlVlfSort(ilUpperBound).tVlf.iEffDate(1)
                    Exit Do
                End If
                ilRet = btrGetPrevious(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'If (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode) And (tlVlf(ilUpperBound).iSellDay = ilDay) And (tlVlf(ilUpperBound).sStatus = "C") Then
            '    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
            '    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
            'Else
            '    ilEffDate0 = 0
            '    ilEffDate1 = 0
            'End If
        Else
            tlSrchKey1.iAirCode = ilVefCode
            tlSrchKey1.iAirDay = ilDay
            tlSrchKey1.iEffDate(0) = ilEffDate0
            tlSrchKey1.iEffDate(1) = ilEffDate1
            ilEffDate0 = 0
            ilEffDate1 = 0
            tlSrchKey1.iAirTime(0) = 0
            tlSrchKey1.iAirTime(1) = 6144   '24*256
            tlSrchKey1.iAirPosNo = 32000
            ilRet = btrGetLessOrEqual(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlfSort(ilUpperBound).tVlf.iAirCode = ilVefCode)
                ilTerminated = False
                'Check for CBS
                If (tlVlfSort(ilUpperBound).tVlf.iTermDate(1) <> 0) Or (tlVlfSort(ilUpperBound).tVlf.iTermDate(0) <> 0) Then
                    If (tlVlfSort(ilUpperBound).tVlf.iTermDate(1) < tlVlfSort(ilUpperBound).tVlf.iEffDate(1)) Or ((tlVlfSort(ilUpperBound).tVlf.iEffDate(1) = tlVlfSort(ilUpperBound).tVlf.iTermDate(1)) And (tlVlfSort(ilUpperBound).tVlf.iTermDate(0) < tlVlfSort(ilUpperBound).tVlf.iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (tlVlfSort(ilUpperBound).tVlf.sStatus <> "P") And (tlVlfSort(ilUpperBound).tVlf.iAirDay = ilDay) And (Not ilTerminated) Then
                    ilEffDate0 = tlVlfSort(ilUpperBound).tVlf.iEffDate(0)
                    ilEffDate1 = tlVlfSort(ilUpperBound).tVlf.iEffDate(1)
                    Exit Do
                End If
                ilRet = btrGetPrevious(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'If (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode) And (tlVlf(ilUpperBound).iSellDay = ilDay) And (tlVlf(ilUpperBound).sStatus = "C") Then
            '    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
            '    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
            'Else
            '    ilEffDate0 = 0
            '    ilEffDate1 = 0
            'End If
        End If
        If slType = "S" Then
            tlSrchKey0.iSellCode = ilVefCode
            tlSrchKey0.iSellDay = ilDay
            tlSrchKey0.iEffDate(0) = ilEffDate0
            tlSrchKey0.iEffDate(1) = ilEffDate1
            tlSrchKey0.iSellTime(0) = 0
            tlSrchKey0.iSellTime(1) = 0
            tlSrchKey0.iSellPosNo = 0
            ilRet = btrGetGreaterOrEqual(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlfSort(ilUpperBound).tVlf.iSellCode = ilVefCode) And (tlVlfSort(ilUpperBound).tVlf.iSellDay = ilDay)
                If tlVlfSort(ilUpperBound).tVlf.sStatus = "C" Then
                    gUnpackDate tlVlfSort(ilUpperBound).tVlf.iEffDate(0), tlVlfSort(ilUpperBound).tVlf.iEffDate(1), slDate
                    llEffDate = gDateValue(slDate)
                    gUnpackDate tlVlfSort(ilUpperBound).tVlf.iTermDate(0), tlVlfSort(ilUpperBound).tVlf.iTermDate(1), slDate
                    If slDate = "" Then
                        slDate = "12/31/2060"
                    End If
                    llTermDate = gDateValue(slDate)
                    
                    If (llDate >= llEffDate) And (llDate <= llTermDate) Or (llTermDate = gDateValue("12/31/2060")) Then   '2-6-15 all tfn links need to be included
                        gUnpackTimeLong tlVlfSort(ilUpperBound).tVlf.iSellTime(0), tlVlfSort(ilUpperBound).tVlf.iSellTime(1), False, llTime
                        slStr = Trim$(str$(llTime))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        tlVlfSort(ilUpperBound).sKey = slStr            'used for sorting by time
                        tlVlfSort(ilUpperBound).lAvailTime = llTime     'used for the binary search to get to
                        ilUpperBound = ilUpperBound + 1
                        'ReDim Preserve tlVlfSort(1 To ilUpperBound) As VLFSORT
                        ReDim Preserve tlVlfSort(0 To ilUpperBound) As VLFSORT
                    Else
                        If llDate < llEffDate Then
                            Exit Do
                        End If
                    End If
                End If
                ilRet = btrGetNext(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Else
            tlSrchKey1.iAirCode = ilVefCode
            tlSrchKey1.iAirDay = ilDay
            tlSrchKey1.iEffDate(0) = ilEffDate0
            tlSrchKey1.iEffDate(1) = ilEffDate1
            tlSrchKey1.iAirTime(0) = 0
            tlSrchKey1.iAirTime(1) = 0
            tlSrchKey1.iAirPosNo = 0
            ilRet = btrGetGreaterOrEqual(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlfSort(ilUpperBound).tVlf.iAirCode = ilVefCode) And (tlVlfSort(ilUpperBound).tVlf.iAirDay = ilDay)
                If tlVlfSort(ilUpperBound).tVlf.sStatus = "C" Then
                    gUnpackDate tlVlfSort(ilUpperBound).tVlf.iEffDate(0), tlVlfSort(ilUpperBound).tVlf.iEffDate(1), slDate
                    llEffDate = gDateValue(slDate)
                    gUnpackDate tlVlfSort(ilUpperBound).tVlf.iTermDate(0), tlVlfSort(ilUpperBound).tVlf.iTermDate(1), slDate
                    If slDate = "" Then
                        slDate = "12/31/2060"
                    End If
                    llTermDate = gDateValue(slDate)
                    If (llDate >= llEffDate) And (llDate <= llTermDate) Or (llTermDate = gDateValue("12/31/2060")) Then '2-6-15 all tfn need to be included
                        gUnpackTimeLong tlVlfSort(ilUpperBound).tVlf.iSellTime(0), tlVlfSort(ilUpperBound).tVlf.iSellTime(1), False, llTime
                        slStr = Trim$(str$(llTime))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        tlVlfSort(ilUpperBound).sKey = slStr
                        tlVlfSort(ilUpperBound).lAvailTime = llTime     'used for the binary search to get to
                        ilUpperBound = ilUpperBound + 1
                        'ReDim Preserve tlVlfSort(1 To ilUpperBound) As VLFSORT
                        ReDim Preserve tlVlfSort(0 To ilUpperBound) As VLFSORT
                    Else
                        If llDate < llEffDate Then
                            Exit Do
                        End If
                    End If
                End If
                ilRet = btrGetNext(hlVlf, tlVlfSort(ilUpperBound).tVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
    ilUpperBound = UBound(tlVlfSort)
    If ilUpperBound > 0 Then
        'ArraySortTyp fnAV(tlVlfSort(), 1), ilUpperBound - 1, 0, LenB(tlVlfSort(1)), 0, LenB(tlVlfSort(1).sKey), 0
        ArraySortTyp fnAV(tlVlfSort(), 0), ilUpperBound, 0, LenB(tlVlfSort(0)), 0, LenB(tlVlfSort(0).sKey), 0
    End If
End Sub

'           gCreateAiringInv - create prepass to generate a report of airing vehicles and the total inventory defined.  Break out
'           the inventory for the week for M-F, Saturday & Sunday buckets.  By option, also show the selling vehicles total inventory
'
Public Sub gCreateAiringInv()
    Dim ilRet As Integer
    Dim ilVehicle As Integer            'loop on vehicles to process
    Dim blFound As Boolean
    Dim slNameCode As String
    Dim slName As String
    Dim ilVefCode As Integer            'vehicle code
    Dim slCode As String
    Dim ilVefIndex As Integer           ' index into vehicle array
    Dim llEffectiveMF As Long           'effective MF date
    Dim llEffectiveSa As Long           'effective Sa date
    Dim llEffectiveSU As Long           'effective Su date
    Dim llEffectiveDate As Long
    Dim llTemp As Long
    Dim ilTemp As Integer
    Dim ilLoopOnDay As Integer          '3 passes for each vehicle:  MF, Sa & Su
    Dim ilDayOfWeek As Integer
    Dim ilSellList() As Integer         'list of selling vehicles to process at a time:  only 1 if sell to air, list of selling if airing selection
    Dim ilLoopOnSellList As Integer
    Dim ilLoopOnVlf As Integer
    Dim ilDate(0 To 1) As Integer
    Dim ilEvt As Integer
    Dim ilWeekDay As Integer
    Dim ilLoopOnLink As Integer
    Dim llTime As Long
    Dim ilNextLink As Integer
    Dim ilLen As Integer
    Dim ilAdjustMF As Integer
    Dim ilAdjustStart As Integer
    Dim ilAdjustEnd As Integer
    Dim blInclMF As Boolean
    Dim blInclSa As Boolean
    Dim blInclSu As Boolean
    Dim blIgnorePass As Boolean
    Dim ilSelectedVefCode As Integer
    Dim ilAnfCode As Integer
    Dim llInputDate As Long
    Dim ilInputDayOfWeek As Integer
    Dim clValue As Currency
    Dim ilUpperAir As Integer
    Dim llTempEffectiveDate As Long
    Dim ilTempDayOfWeek As Integer
    Dim llTempStartDate As Long
    Dim llTempEndDate As Long
    Dim llSsfDate As Long
    Dim llLatestDate As Long
    Dim slDate As String
    ReDim ilEvtType(0 To 14) As Integer
    Dim ilType As Integer

    
    imVlfRecLen = Len(tmVlf)
    hmVLF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateAiringInvErr
    gBtrvErrorMsg ilRet, "gCreateAiringInv (btrOpen: VLF.BTR)", RptSel
    On Error GoTo 0
    
    imCbfRecLen = Len(tmCbf)
    hmCbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateAiringInvErr
    gBtrvErrorMsg ilRet, "gCreateAiringInv (btrOpen: Cbf.BTR)", RptSel
    On Error GoTo 0
    
    imSsfRecLen = Len(tmSsf)
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateAiringInvErr
    gBtrvErrorMsg ilRet, "gCreateAiringInv (btrOpen: SSF.BTR)", RptSel
    On Error GoTo 0
    
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreateAiringInvErr
    gBtrvErrorMsg ilRet, "gCreateAiringInv (btrOpen: lcf.BTR)", RptSel
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    
    gObtainCodesForMultipleLists 5, tgNamedAvail(), imInclAnfCodes, imUseAnfCodes(), RptSel
    
    'set the type of events to get fro the day (only  avails)
    For ilTemp = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilTemp) = False
    Next ilTemp
    ilEvtType(2) = True
    
    blInclMF = True
    blInclSa = True
    blInclSu = True
    If RptSel!ckcSel2(0).Value = vbUnchecked Then
        blInclMF = False
    End If
    If RptSel!ckcSel2(1).Value = vbUnchecked Then
        blInclSa = False
    End If
    If RptSel!ckcSel2(2).Value = vbUnchecked Then
        blInclSu = False
    End If
    
    '2-6-14 determine the effective date for m-f (backup date entered to monday), sa (backup or increment to sa), su (increment to sun)
    'all valid transactions need to show
    '        llInputDate = gDateValue(RptSel!edcSelA.Text)
    llInputDate = gDateValue(RptSel!CSI_CalDateA.Text)      '12-11-19 change to use csi calendar control
    llTemp = llInputDate
    ilInputDayOfWeek = gWeekDayLong(llInputDate)
    ilDayOfWeek = ilInputDayOfWeek
    
    Do While ilDayOfWeek <> 0           'backup MF to monday
        llTemp = llTemp - 1
        ilDayOfWeek = gWeekDayLong(llTemp)
    Loop
    llEffectiveMF = llTemp
    llInputDate = llEffectiveMF                 'save the original start of week from user entered input date.  will use it to retrieve linkages
    
    'Sat
    llTemp = llInputDate
    ilDayOfWeek = ilInputDayOfWeek
    If ilDayOfWeek = 5 Then
        llEffectiveSa = llTemp
    Else
        If ilDayOfWeek = 6 Then     'its sunday, backup to sat
            llTemp = llTemp - 1
        Else
            Do While ilDayOfWeek <> 5           'its a m-f date, increment to sa
                llTemp = llTemp + 1
                ilDayOfWeek = gWeekDayLong(llTemp)
            Loop
        End If
        llEffectiveSa = llTemp
    End If
    
    'Sun
    llTemp = llInputDate
    ilDayOfWeek = ilInputDayOfWeek
    If ilDayOfWeek = 6 Then
        llEffectiveSU = llTemp
    Else
        Do While ilDayOfWeek <> 6
            llTemp = llTemp + 1
            ilDayOfWeek = gWeekDayLong(llTemp)
        Loop
        llEffectiveSU = llTemp
    End If
    
    ReDim tmAiringInv(0 To 0) As AIRING_INV
    
    'process each vehicle 3 times.  Process each vehicle for MF links, Sa links & Su links
    For ilVehicle = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
        If (RptSel!lbcSelection(1).Selected(ilVehicle)) Then
            slNameCode = tgAirNameCode(ilVehicle).sKey
          
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSelectedVefCode = Val(slCode)
            ilVefIndex = gBinarySearchVef(ilSelectedVefCode)
            
            'obtain all the SSF records for the week
            ReDim tmSSFWeek(0 To 0) As SSFWEEK
                            
            'first find the closest week of ssf airing records (to determine which week to gather)
            'find the closest record to the entered week, then gather M-f for that week
            gPackDateLong llEffectiveMF, ilDate(0), ilDate(1)
            tmSsfSrchKey.iType = 0 'slType
            tmSsfSrchKey.iVefCode = ilSelectedVefCode
            tmSsfSrchKey.iDate(0) = ilDate(0)
            tmSsfSrchKey.iDate(1) = ilDate(1)
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            imSsfRecLen = Len(tmSsf)
            ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llSsfDate
                'determine the start week of the ssf returned
                ilTempDayOfWeek = gWeekDayLong(llSsfDate)
                
                Do While ilTempDayOfWeek <> 0           'backup MF to monday
                    llSsfDate = llSsfDate - 1
                    ilTempDayOfWeek = gWeekDayLong(llSsfDate)
                Loop
    
                'gather the entire week
                llTempStartDate = llSsfDate
                llTempEndDate = llSsfDate + 6
                gPackDateLong llTempStartDate, ilDate(0), ilDate(1)
                tmSsfSrchKey.iType = 0 'slType
                tmSsfSrchKey.iVefCode = ilSelectedVefCode
                tmSsfSrchKey.iDate(0) = ilDate(0)
                tmSsfSrchKey.iDate(1) = ilDate(1)
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                imSsfRecLen = Len(tmSsf)
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llSsfDate
    
                    Do While (ilRet = BTRV_ERR_NONE) And (llSsfDate >= llTempStartDate And llSsfDate <= llTempEndDate) And (tmSsf.iVefCode = ilSelectedVefCode)
                        'determine if this day should be included
                        ilTempDayOfWeek = gWeekDayLong(llSsfDate)
                        
                        If (blInclMF And ilTempDayOfWeek <= 4) Or (blInclSa And ilTempDayOfWeek = 5) Or (blInclSu And ilTempDayOfWeek = 6) Then
                            ilUpperAir = UBound(tmSSFWeek)
                            tmSSFWeek(ilUpperAir).iVefCode = tmSsf.iVefCode
                            tmSSFWeek(ilUpperAir).lSSFCode = tmSsf.lCode
                            tmSSFWeek(ilUpperAir).iDate(0) = tmSsf.iDate(0)
                            tmSSFWeek(ilUpperAir).iDate(1) = tmSsf.iDate(1)
                            tmSSFWeek(ilUpperAir).iDayOfWeek = ilTempDayOfWeek
                            ReDim Preserve tmSSFWeek(0 To ilUpperAir + 1) As SSFWEEK
                        End If
                        
                        imSsfRecLen = Len(tmSsf)
                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llSsfDate
                        Else
                            Exit Do
                        End If
                    Loop
                    'obtain the airing vehicle info
    '                        mObtainVlf "A", hmVLF, ilSelectedVefCode, llTempStartDate, tmVlfSortMF()        'get the associated selling vehicles links with this airing  vehicle
    '                        mObtainVlf "A", hmVLF, ilSelectedVefCode, llTempStartDate + 5, tmVlfSortSa()      'get the associated selling vehicles links with this airing  vehicle
    '                        mObtainVlf "A", hmVLF, ilSelectedVefCode, llTempStartDate + 6, tmVlfSortSu()      'get the associated selling vehicles links with this airing  vehicle
                    mObtainVlf "A", hmVLF, ilSelectedVefCode, llEffectiveMF, tmVlfSortMF()        'get the associated selling vehicles links with this airing  vehicle
                    mObtainVlf "A", hmVLF, ilSelectedVefCode, llEffectiveSa, tmVlfSortSa()       'get the associated selling vehicles links with this airing  vehicle
                    mObtainVlf "A", hmVLF, ilSelectedVefCode, llEffectiveSU, tmVlfSortSu()       'get the associated selling vehicles links with this airing  vehicle
                    mGatherAiringInv ilSelectedVefCode
                End If
            End If
            
            
            For ilLoopOnDay = 1 To 3        'pass 1 = M-F, pass 2 = SA, pass 3 = su
                blIgnorePass = False
                If ilLoopOnDay = 1 Then     'm-f
                    llEffectiveDate = llEffectiveMF
                    'looping factor in case the first days of the week are just placeholders.  different libraries across the week
                    ilAdjustStart = 1
                    ilAdjustEnd = 5
                    If Not blInclMF Then            'dont include MF pass
                        blIgnorePass = True         'ignore the processing
                    End If
                ElseIf ilLoopOnDay = 2 Then     'sa
                    llEffectiveDate = llEffectiveSa
                    ilAdjustStart = 1
                    ilAdjustEnd = 1
                    If Not blInclSa Then            'dont includeSA pass
                        blIgnorePass = True         'ignore the processing
                    End If
                Else                            'su
                    llEffectiveDate = llEffectiveSU
                    ilAdjustStart = 1
                    ilAdjustEnd = 1
                    If Not blInclSu Then            'dont includeSU pass
                        blIgnorePass = True         'ignore the processing
                    End If
                End If
                'ReDim tmVlfSort(1 To 1) As VLFSORT
                ReDim ilSellList(0 To 0) As Integer
                
                If Not blIgnorePass Then
                    llTempEffectiveDate = llEffectiveDate
                    If RptSel!ckcInclCommentsA.Value = vbChecked Then              'include the selling vehicles inventory?
                        ilUpperAir = UBound(tmAiringInv)
         
                        If ilLoopOnDay = 1 Then
                            mMoveVLFSortToTemp tmVlfSortMF()
                        ElseIf ilLoopOnDay = 2 Then
                            mMoveVLFSortToTemp tmVlfSortSa()
                        Else
                            mMoveVLFSortToTemp tmVlfSortSu()
                        End If
                         
                        'linkages have already been created
                        'mObtainVlf "A", hmVLF, ilSelectedVefCode, llTempEffectiveDate, tmVlfSort()        'get the associated selling vehicles links with this airing  vehicle
                        
                        'air to sell, get the list of selling vehicles associated with the airing.  List to be used to get the avail lengths from SSF
                        For ilLoopOnVlf = LBound(tmVlfSort) To UBound(tmVlfSort) - 1
                            ilTemp = tmVlfSort(ilLoopOnVlf).tVlf.iSellCode
                            blFound = False
                            For ilLoopOnSellList = 0 To UBound(ilSellList) - 1
                                If ilSellList(ilLoopOnSellList) = ilTemp Then
                                    blFound = True
                                    Exit For
                                End If
                            Next ilLoopOnSellList
                            If Not blFound Then
                                ilSellList(UBound(ilSellList)) = ilTemp
                                ReDim Preserve ilSellList(0 To UBound(ilSellList) + 1) As Integer
                            End If
                        Next ilLoopOnVlf
                    
                        'list of selling vehicles are now created so that the SSF can be read to get the avail lengths
                        For ilLoopOnSellList = LBound(ilSellList) To UBound(ilSellList) - 1
                            ilVefCode = ilSellList(ilLoopOnSellList)
                            llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
                            ilType = 0
                            For ilAdjustMF = ilAdjustStart To ilAdjustEnd       '
                                gPackDateLong llTempEffectiveDate + ilAdjustMF - 1, ilDate(0), ilDate(1)
                                slDate = Format$(llTempEffectiveDate + ilAdjustMF - 1, "m/d/yy")
    
                                tmSsfSrchKey.iType = 0
                                tmSsfSrchKey.iVefCode = ilVefCode
                                tmSsfSrchKey.iDate(0) = ilDate(0)
                                tmSsfSrchKey.iDate(1) = ilDate(1)
                                tmSsfSrchKey.iStartTime(0) = 0
                                tmSsfSrchKey.iStartTime(1) = 0
                                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                                'ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    
                                'If no ssf built, need to retrieve from calendar
                                If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iVefCode <> ilVefCode) Or ((tmSsf.iDate(0) <> ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))) Then
                                    If (llTempEffectiveDate + ilAdjustMF - 1 > llLatestDate) Then
                                        
                                        ReDim tlLLC(0 To 0) As LLC  'Merged library names
                                        If tgMVef(ilVefIndex).sType <> "G" Then
                                            ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                                            If ilWeekDay = 1 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 2 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 3 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 4 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 5 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 6 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                                            ElseIf ilWeekDay = 7 Then
                                                 ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                                            End If
                                        End If
                         
                                        tmSsf.iType = 0
                                        tmSsf.iVefCode = ilVefCode
                                        tmSsf.iDate(0) = ilDate(0)
                                        tmSsf.iDate(1) = ilDate(1)
                                        gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                                        tmSsf.iCount = 0
                        
                                        For ilTemp = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                        
                                            tmAvail.iRecType = Val(tlLLC(ilTemp).sType)
                                            gPackTime tlLLC(ilTemp).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                                            tmAvail.iLtfCode = tlLLC(ilTemp).iLtfCode
                                            tmAvail.iAvInfo = tlLLC(ilTemp).iAvailInfo Or tlLLC(ilTemp).iUnits
                                            tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilTemp).sLength))
                                            tmAvail.ianfCode = Val(tlLLC(ilTemp).sName)
                                            tmAvail.iNoSpotsThis = 0
                                            tmAvail.iOrigUnit = 0
                                            tmAvail.iOrigLen = 0
                                            tmSsf.iCount = tmSsf.iCount + 1
                                            tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                                        Next ilTemp
                                        ilRet = BTRV_ERR_NONE
                                    End If
                                End If
                                
                                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))
                                   'loop thru the SSF records and pick up the selling avail time and length.  Find all links in tmVlfSort that have the matching sell vehicle to create
                                   'a record for the associated airiing vehicle
                                    ilEvt = 1
                                    ilNextLink = LBound(tmVlfSort)
                                    Do While ilEvt <= tmSsf.iCount
                                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then 'Contract Avails only
                                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                            'ilNextLink = LBound(tmVlfSort)
                                            For ilLoopOnLink = ilNextLink To UBound(tmVlfSort) - 1
                                                If tmVlfSort(ilLoopOnLink).lAvailTime = llTime And tmVlfSort(ilLoopOnLink).tVlf.iSellCode = ilVefCode Then
                                                    'get the avail length & create an entry to output
                                                    ilLen = tmAvail.iLen
                                                    ilAnfCode = tmAvail.ianfCode
                                                    If gFilterLists(ilAnfCode, imInclAnfCodes, imUseAnfCodes()) Then            'valid avail name to use?
                                                        If ilLoopOnDay = 1 Then     'M-F
                                                            tmAiringInv(ilUpperAir).lSellMFInv = tmAiringInv(ilUpperAir).lSellMFInv + ilLen
                                                        ElseIf ilLoopOnDay = 2 Then     'SAT
                                                            tmAiringInv(ilUpperAir).lSellSatInv = tmAiringInv(ilUpperAir).lSellSatInv + ilLen
                                                        Else                            'sun
                                                            tmAiringInv(ilUpperAir).lSellSunInv = tmAiringInv(ilUpperAir).lSellSunInv + ilLen
                                                        End If
                                                    End If                              'named avail filter
                                                Else
                                                    If tmVlfSort(ilLoopOnLink).lAvailTime > llTime Then      'if this time is past the time just finished, stop and save the current index.
                                                                                                            'use that index as a starting point in table
                                                        ilNextLink = ilLoopOnLink
                                                        Exit For
                                                    End If
                                                End If
                                                
                                            Next ilLoopOnLink
                                            ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                                        End If
                                        ilEvt = ilEvt + 1                                       'get next event
                                    Loop                                            'Do While ilEvt <= tmSsf.iCount
                                    Exit Do
                                    'read next ssf to see if theres a continuation of same day and vehicle
                                    imSsfRecLen = Len(tmSsf)
                                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop            'go see if the next ssf is matching day and vehicle
                                
                            Next ilAdjustMF             'loop on m-f if pass 1 (only looping on multiple days if 0 avail lengths found
                        Next ilLoopOnSellList           'process next selling vehicle (could be multiple selling vehicles if airing to selling option was selected)
                    End If                              'RptSel!ckcInclCommentsA.Value = vbChecked
                End If                          'blIgnorePass
            Next ilLoopOnDay                    'process next pass for MF, Sa, Su
            ReDim Preserve tmAiringInv(0 To UBound(tmAiringInv) + 1) As AIRING_INV
    
        End If
       
    Next ilVehicle                              'next sell or air vehicle in selection list
    
    'write out prepass airing vehicle inventory
    For ilVehicle = LBound(tmAiringInv) To UBound(tmAiringInv) - 1
        'tmCbf.lValue(1) = airing veh M-F total inv in secs
        'tmCbf.lValue(2) = airing veh Sat total inv
        'tmCbf.lValue(3) = airing veh Sun total inv
        'tmCbf.lValue(4) = airing veh total week inv
        'tmCbf.lValue(5) = selling veh M-F total inv in sec
        'tmCbf.lValue(6) = selling veh Sat total inv
        'tmCbf.lValue(7) = selling veh Sun total Inv
        'tmCbf.lvalue(8) = selling veh total week inv
        '
        'tmCbf.sdystms = length in hrs/min/sec in string for each bucket retained:  1-10 = airing m-f, 11-20 = airing sat, 21-30 = airing sun, 31-40 = airing week total
        '                                                                           41-50 = sell m-f, 51-60 = selling sat, 61-70 = selling sun, 71-80 selling total
        tmCbf.lGenTime = lgNowTime
        tmCbf.iGenDate(0) = igNowDate(0)
        tmCbf.iGenDate(1) = igNowDate(1)
        tmCbf.iVefCode = tmAiringInv(ilVehicle).iVefAirCode          'airing vehicle code
        
        'tmCbf.lValue(1) = tmAiringInv(ilVehicle).lAirMFInv
        'tmCbf.lValue(2) = tmAiringInv(ilVehicle).lAirSatInv
        'tmCbf.lValue(3) = tmAiringInv(ilVehicle).lAirSunInv
        'tmCbf.lValue(4) = tmAiringInv(ilVehicle).lAirMFInv + tmAiringInv(ilVehicle).lAirSatInv + tmAiringInv(ilVehicle).lAirSunInv      'airing weeks total
        
        'tmCbf.lValue(5) = tmAiringInv(ilVehicle).lSellMFInv
        'tmCbf.lValue(6) = tmAiringInv(ilVehicle).lSellSatInv
        'tmCbf.lValue(7) = tmAiringInv(ilVehicle).lSellSunInv
        'tmCbf.lValue(8) = tmAiringInv(ilVehicle).lSellMFInv + tmAiringInv(ilVehicle).lSellSatInv + tmAiringInv(ilVehicle).lSellSunInv          'selling weeks total
        
        tmCbf.lValue(0) = tmAiringInv(ilVehicle).lAirMFInv
        tmCbf.lValue(1) = tmAiringInv(ilVehicle).lAirSatInv
        tmCbf.lValue(2) = tmAiringInv(ilVehicle).lAirSunInv
        tmCbf.lValue(3) = tmAiringInv(ilVehicle).lAirMFInv + tmAiringInv(ilVehicle).lAirSatInv + tmAiringInv(ilVehicle).lAirSunInv      'airing weeks total
        
        tmCbf.lValue(4) = tmAiringInv(ilVehicle).lSellMFInv
        tmCbf.lValue(5) = tmAiringInv(ilVehicle).lSellSatInv
        tmCbf.lValue(6) = tmAiringInv(ilVehicle).lSellSunInv
        tmCbf.lValue(7) = tmAiringInv(ilVehicle).lSellMFInv + tmAiringInv(ilVehicle).lSellSatInv + tmAiringInv(ilVehicle).lSellSunInv          'selling weeks total
        
        
        'if discrepancy only, compare the total week airing against selling inventory.  If different, show on discrep only
        'ckcInclCommentsA = Include selling inventory
        'ckcADate= if including selling inventory, discrepancy only option
        'if Airing Inventory only - exclude vehicles without inventory
        'if Selling included - exclude if both airing and selling without inventory
        'If ((RptSel!ckcInclCommentsA.Value = vbUnchecked) And (tmCbf.lValue(4) <> 0)) Or ((RptSel!ckcADate.Value = vbUnchecked) And (tmCbf.lValue(4) <> 0 And tmCbf.lValue(8) <> 0)) Or ((RptSel!ckcADate.Value = vbChecked) And (tmCbf.lValue(4) <> tmCbf.lValue(8))) Then
        If ((RptSel!ckcInclCommentsA.Value = vbUnchecked) And (tmCbf.lValue(3) <> 0)) Or ((RptSel!ckcADate.Value = vbUnchecked) And (tmCbf.lValue(3) <> 0 And tmCbf.lValue(7) <> 0)) Or ((RptSel!ckcADate.Value = vbChecked) And (tmCbf.lValue(3) <> tmCbf.lValue(7))) Then
            tmCbf.sDysTms = ""
            For ilLoopOnDay = 1 To 8
                If tmCbf.lValue(ilLoopOnDay - 1) > 0 Then
                    clValue = tmCbf.lValue(ilLoopOnDay - 1)
                    slCode = gCurrencyToLength(clValue)
                    slCode = Trim$(slCode)
                    Do While Len(slCode) < 9
                        slCode = " " & slCode            'fill with blanks for length of 10
                    Loop
                Else
                    slCode = "         "
                End If
                slCode = slCode & "|"
                tmCbf.sDysTms = RTrim(tmCbf.sDysTms) & slCode
                
            Next ilLoopOnDay
            
            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            On Error GoTo gCreateAiringInvErr
            gBtrvErrorMsg ilRet, "gCreateAiringInv (btrInsert: Cbf.BTR)", RptSel
            On Error GoTo 0
        End If              'discrepany only
    Next ilVehicle
    
    Erase ilSellList
    Erase tmVlfSort, tmVlfSortMF, tmVlfSortSa, tmVlfSortSu
    Erase tmAiringInv
    Erase tmSSFWeek
    Erase tlLLC
    
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmCbf)
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmVLF
    btrDestroy hmCbf
    
    Exit Sub
gCreateAiringInvErr:
    Resume Next
End Sub

'           mGatherAiringInv - Accumulate avails length from airing vehicles ssf
'           <input> ilVefCode - airing vehicle internal code
'           tmAiringInv - array of airing vehicle and its accumulated avail lengths (in seconds)
'Public Sub mGatherAiringInv(ilVefCode As Integer, llEffectiveDate As Long, ilAdjustStart As Integer, ilAdjustEnd As Integer)
Public Sub mGatherAiringInv(ilVefCode As Integer)
    Dim ilAdjustMF As Integer
    Dim ilDate(0 To 1) As Integer
    Dim ilRet As Integer
    Dim ilEvt As Integer
    Dim llTime As Long
    Dim ilLen As Integer
    Dim ilAnfCode As Integer
    Dim ilUpperAir As Integer
    Dim ilDayOfWeek As Integer
    Dim ilWhichDay As Integer
    Dim llSsfDate As Long
    Dim ilSSFDayOfWeek As Integer
    Dim llTempEffectiveDate As Long
    Dim ilTempDate(0 To 1) As Integer
    Dim blActiveLink As Boolean
    
    ilUpperAir = UBound(tmAiringInv)
    tmAiringInv(ilUpperAir).iVefAirCode = ilVefCode
    'loop thru the array of SSF records gathered for the week
    For ilAdjustMF = LBound(tmSSFWeek) To UBound(tmSSFWeek) - 1
        imSsfRecLen = Len(tmSsf)
        tmSSFSrchKey3.lCode = tmSSFWeek(ilAdjustMF).lSSFCode
        ilRet = gSSFGetEqualKey3(hmSsf, tmSsf, imSsfRecLen, tmSSFSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
    
        If (ilRet = BTRV_ERR_NONE) Then
            'date found must be for the same day of effective date sent to process.  That is, if effective date is Saturday, the ssf must be a sat date.
            'Otherwise will overstate the avails because the program vehicle isnt defined for M-F when  SSF found with earlier date
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llSsfDate
    
            ilSSFDayOfWeek = gWeekDayLong(llSsfDate)         '0-4 = mon-fri, 5 = sat, 6 = sun
            'loop thru the SSF records and pick up the selling avail time and length.  Find all links in tmVlfSort that have the matching sell vehicle to create
            'a record for the associated airiing vehicle
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
                LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then 'Contract Avails only
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    'determine if this link has been terminated
'If llTime = 34920 Then
'llTime = llTime
'End If
                    blActiveLink = mTestTermDate(ilVefCode, ilSSFDayOfWeek, llTime)
                    If blActiveLink Then
                        'get the avail length & create an entry to output
                        ilLen = tmAvail.iLen
                        ilAnfCode = tmAvail.ianfCode                '10-30-14
                        'test for inclusion of avail name
                        If gFilterLists(ilAnfCode, imInclAnfCodes, imUseAnfCodes()) Then
                            'If ilWhichDay = 1 Then
                            If tmSSFWeek(ilAdjustMF).iDayOfWeek <= 4 Then           'mo-fr = 0-4, sa = 5, su = 6
                                tmAiringInv(ilUpperAir).lAirMFInv = tmAiringInv(ilUpperAir).lAirMFInv + ilLen
                            'ElseIf ilWhichDay = 2 Then
                            ElseIf tmSSFWeek(ilAdjustMF).iDayOfWeek = 5 Then
                                tmAiringInv(ilUpperAir).lAirSatInv = tmAiringInv(ilUpperAir).lAirSatInv + ilLen
                            Else
                                tmAiringInv(ilUpperAir).lAirSunInv = tmAiringInv(ilUpperAir).lAirSunInv + ilLen
                            End If
                        Else
                            ilAnfCode = ilAnfCode
                        End If
                    Else
                        blActiveLink = blActiveLink
                    End If              'blactivelink
                    ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                End If
                ilEvt = ilEvt + 1                                       'get next event
            Loop                                            'Do While ilEvt <= tmSsf.iCount
        End If
    Next ilAdjustMF             'loop on m-f if pass 1 (only looping on multiple days if 0 avail lengths found
    Exit Sub
End Sub

Private Sub mSetLibProcessedFlag(ByVal slDay As String, ByVal slStartTime As String, ByVal slEndTime As String)
    Dim ilTimeCounter As Integer
    Dim ilCounter As Integer
    Dim slStartTimeA As String
    Dim slCurDay As String
    Dim ilDayOfWeek As Integer
    Dim slDayOfWeek As String

    If slStartTime = "11PM" Then
        For ilCounter = 0 To UBound(tmLibDaysTimes)
            'get start times for library version
            gUnpackTime tmLibDaysTimes(ilCounter).iStartTime(0), tmLibDaysTimes(ilCounter).iStartTime(1), "A", 1, slStartTimeA
            
            'bug found in gUnpackTime returning "12M" instead of "12AM/12PM"
            If StrComp(slStartTimeA, "12M", vbTextCompare) = 0 Then
                slStartTimeA = "12AM"
            ElseIf StrComp(slStartTimeA, "12N", vbTextCompare) = 0 Then
                slStartTimeA = "12PM"
            End If
            
            gUnpackDate tmLibDaysTimes(ilCounter).iDays(0), tmLibDaysTimes(ilCounter).iDays(1), slCurDay
            If tmLibDaysTimes(ilCounter).iDays(0) <= 7 Then
                ilDayOfWeek = tmLibDaysTimes(ilCounter).iDays(0)
            Else
                ilDayOfWeek = (gWeekDayStr(slCurDay) + 1)
            End If
            
            slDayOfWeek = mGetDayOfWeek(ilDayOfWeek)        'returns Mo ... Su
            
            If (slDayOfWeek = slDay And CDate(slStartTimeA) = CDate(slStartTime)) Then
                tmLibDaysTimes(ilCounter).bProcessed = True
                Exit For
            End If
        Next ilCounter
    Else
        Do While CDate(slStartTime) < CDate(slEndTime)
            For ilCounter = 0 To UBound(tmLibDaysTimes)
                'get start times for library version
                gUnpackTime tmLibDaysTimes(ilCounter).iStartTime(0), tmLibDaysTimes(ilCounter).iStartTime(1), "A", 1, slStartTimeA
                
                'bug found in gUnpackTime returning "12M" instead of "12AM/12PM"
                If StrComp(slStartTimeA, "12M", vbTextCompare) = 0 Then
                    slStartTimeA = "12AM"
                ElseIf StrComp(slStartTimeA, "12N", vbTextCompare) = 0 Then
                    slStartTimeA = "12PM"
                End If
                
                gUnpackDate tmLibDaysTimes(ilCounter).iDays(0), tmLibDaysTimes(ilCounter).iDays(1), slCurDay
                If tmLibDaysTimes(ilCounter).iDays(0) <= 7 Then
                    ilDayOfWeek = tmLibDaysTimes(ilCounter).iDays(0)
                Else
                    ilDayOfWeek = (gWeekDayStr(slCurDay) + 1)
                End If
                
                slDayOfWeek = mGetDayOfWeek(ilDayOfWeek)        'returns Mo ... Su
                
                If (slDayOfWeek = slDay And CDate(slStartTimeA) = CDate(slStartTime)) Then
                    tmLibDaysTimes(ilCounter).bProcessed = True
                    Exit For
                End If
            Next ilCounter
            slStartTime = DateAdd("h", 1, CDate(slStartTime))
        Loop
    End If
End Sub

'               mTestTermDate - check to see if this avail is still active.  Terminated links are not built into array
'               <input> ilVefCode - airing vehicle code
'                       ilSSFDayofWEek - 0-4 = M-F, 5 = Sat, 6 = Sun
Public Function mTestTermDate(ilVefCode As Integer, ilSSFDayOfWeek As Integer, llTime As Long) As Boolean
    Dim ilLoopOnLinks As Integer
    Dim blFoundEntry As Boolean
    Dim llVlfAirTime As Long
    
    blFoundEntry = False
    If ilSSFDayOfWeek <= 4 Then             'm-f
        For ilLoopOnLinks = LBound(tmVlfSortMF) To UBound(tmVlfSortMF) - 1
            gUnpackTimeLong tmVlfSortMF(ilLoopOnLinks).tVlf.iAirTime(0), tmVlfSortMF(ilLoopOnLinks).tVlf.iAirTime(1), False, llVlfAirTime
            'vlf keeps mon-fri as 0 , sat = 6, sun = 7 , Dayofweek conversion converts mon-sun as 0-6
            If (llVlfAirTime = llTime) And ((ilSSFDayOfWeek <= 4 And tmVlfSortMF(ilLoopOnLinks).tVlf.iAirDay = 0)) Then
                mTestTermDate = True            ' active link
                blFoundEntry = True
                Exit For
            End If
        Next ilLoopOnLinks
    ElseIf ilSSFDayOfWeek = 5 Then     'Sat
        For ilLoopOnLinks = LBound(tmVlfSortSa) To UBound(tmVlfSortSa) - 1
            gUnpackTimeLong tmVlfSortSa(ilLoopOnLinks).tVlf.iAirTime(0), tmVlfSortSa(ilLoopOnLinks).tVlf.iAirTime(1), False, llVlfAirTime
            
            If (llVlfAirTime = llTime) And ((ilSSFDayOfWeek = 5 And tmVlfSortSa(ilLoopOnLinks).tVlf.iAirDay = 6)) Then
                mTestTermDate = True            ' active link
                blFoundEntry = True
                Exit For
            End If
        Next ilLoopOnLinks
    Else
        For ilLoopOnLinks = LBound(tmVlfSortSu) To UBound(tmVlfSortSu) - 1
            gUnpackTimeLong tmVlfSortSu(ilLoopOnLinks).tVlf.iAirTime(0), tmVlfSortSu(ilLoopOnLinks).tVlf.iAirTime(1), False, llVlfAirTime
            
            If (llVlfAirTime = llTime) And (ilSSFDayOfWeek = 6 And tmVlfSortSu(ilLoopOnLinks).tVlf.iAirDay = 7) Then
                mTestTermDate = True            ' active link
                blFoundEntry = True
                Exit For
            End If
        Next ilLoopOnLinks
    End If
    If Not blFoundEntry Then
        mTestTermDate = False
    End If
    Exit Function
End Function

'               mMoveVLFSortToTemp - copy the contents of the VLF records into a temporary common array
'               <input> tlVLFSort() - array of M-F linkages, Sa linkages, or Su linkages
Public Sub mMoveVLFSortToTemp(tlVlfSort() As VLFSORT)
    Dim illoop As Integer

        'ReDim tmVlfSort(1 To UBound(tlVlfSort))
        ReDim tmVlfSort(0 To UBound(tlVlfSort))
        For illoop = LBound(tlVlfSort) To UBound(tlVlfSort) - 1
            tmVlfSort(illoop) = tlVlfSort(illoop)
        Next illoop
End Sub

'               gGenPkgList - generate list of package vehicles and their associated hidden vehicle definitions
'
Public Sub gGenPkgList()
    Dim ilRet As Integer
    Dim ilLoopOnVef As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVefIndex As Integer           'index to package vehicle
    Dim illoop As Integer
    Dim hlPvf As Integer
    Dim tlPvf As PVF
    Dim tlPvfSrchKey As LONGKEY0
    Dim ilHiddenVefInx As Integer           'index for hidden line vehicle

        hlPvf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlPvf, "", sgDBPath & "Pvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenPkgListErr
        gBtrvErrorMsg ilRet, "gGenPkgList (btrOpen: PvF.BTR)", RptSel
        On Error GoTo 0
        
        hmGrf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenPkgListErr
        gBtrvErrorMsg ilRet, "gGenPkgList (btrOpen: GrF.BTR)", RptSel
        On Error GoTo 0

        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime

        For ilLoopOnVef = 0 To RptSel!lbcSelection(10).ListCount - 1
            If RptSel!lbcSelection(10).Selected(ilLoopOnVef) = True Then
'                slNameCode = tgVehicle(ilLoopOnVef).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                ilVefCode = RptSel!lbcSelection(10).ItemData(ilLoopOnVef)
'                ilVefCode = Val(slCode)
                ilVefIndex = gBinarySearchVef(ilVefCode)
                If ilVefIndex >= 0 Then
                    tmVef = tgMVef(ilVefIndex)
                    If tmVef.lPvfCode > 0 And ((RptSel!ckcSelC3(0).Value = vbChecked And tmVef.sState = "D") Or (tmVef.sState = "A")) Then           'std pkg, include dormant?
                        tlPvfSrchKey.lCode = tmVef.lPvfCode
                        ilRet = btrGetEqual(hlPvf, tlPvf, Len(tlPvf), tlPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                        Do While ilRet = BTRV_ERR_NONE
                            For illoop = LBound(tlPvf.iVefCode) To UBound(tlPvf.iVefCode) Step 1
                                If tlPvf.iNoSpot(illoop) > 0 Then           'ignore the vehicle without any spot definition
                                    'see if pkg within vehicle is dormnat
                                    ilHiddenVefInx = gBinarySearchVef(tlPvf.iVefCode(illoop))
                                    If ilHiddenVefInx >= 0 Then     '4-17-20 ignore vef codes not found
                                        If ((RptSel!ckcSelC3(0).Value = vbChecked And tgMVef(ilHiddenVefInx).sState = "D") Or (tgMVef(ilHiddenVefInx).sState = "A")) Then
                                            tmGrf.iVefCode = ilVefCode              'package vehicle
                                            tmGrf.iCode2 = tlPvf.iVefCode(illoop)      'vehicle within package
                                            tmGrf.iRdfCode = tlPvf.iRdfCode(illoop)     'daypart
                                            'tmGrf.iPerGenl(1) = tlPvf.iNoSpot(ilLoop)  '#spots
                                            'tmGrf.iPerGenl(2) = tlPvf.iPctRate(ilLoop)  '%rate splt (for vefStdPrice = 2) (xxx.xx)
                                            tmGrf.iPerGenl(0) = tlPvf.iNoSpot(illoop)  '#spots
                                            tmGrf.iPerGenl(1) = tlPvf.iPctRate(illoop)  '%rate splt (for vefStdPrice = 2) (xxx.xx)
                                            tmGrf.sGenDesc = Trim$(tlPvf.sName)
                                            ilRet = btrInsert(hmGrf, tmGrf, Len(tmGrf), INDEXKEY0)
                                        End If
                                    Else                '
                                        ilHiddenVefInx = ilHiddenVefInx
                                    End If
                                End If
                            Next illoop
                            If tlPvf.lLkPvfCode > 0 Then
                                tlPvfSrchKey.lCode = tlPvf.lLkPvfCode
                                ilRet = btrGetEqual(hlPvf, tlPvf, Len(tlPvf), tlPvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                On Error GoTo gGenPkgListErr
                                gBtrvErrorMsg ilRet, "gGenPkgList (btrInsert: GrF.BTR)", RptSel
                                On Error GoTo 0
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            End If
        Next ilLoopOnVef
    
        ilRet = btrClose(hlPvf)
        ilRet = btrClose(hmGrf)
        btrDestroy hlPvf
        btrDestroy hmGrf

        Exit Sub
gGenPkgListErr:
        On Error GoTo 0
        Resume Next
End Sub

'TTP 10791 - Copy Rotations by Advertiser report: add special export option
Function mExportCopyRotRecord(tlCpr As CPR, tlCrf As CRF, slVehicleNameString As String) As String
    Dim slExportLine As String
    Dim slDate As String
    Dim slString As String
    Dim slTime As String
    Dim slvalue As String
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim tlChf As CHF
    Dim tlCif As CIF
    Dim tlCpf As CPF
    Dim tlMcf As MCF
    
    '---------------------------------
    'Load Contract
    tmChfSrchKey.lCode = tlCrf.lChfCode
    ilRet = btrGetEqual(hmCHF, tlChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mExportCopyRotRecord = "Error getting CHF: " & ilRet
        Exit Function
    End If
    '---------------------------------
    'Load Copy Inventory
    tmCifSrchKey.lCode = tlCpr.lFt2CefCode
    ilRet = btrGetEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mExportCopyRotRecord = "Error getting CIF: " & ilRet
        Exit Function
    End If
    '---------------------------------
    'Load Copy Prodct ISCI
    tmCpfSrchKey.lCode = tlCif.lcpfCode
    ilRet = btrGetEqual(hmCpf, tlCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mExportCopyRotRecord = "Error getting CPF: " & ilRet
        Exit Function
    End If
    '---------------------------------
    'Load Media Code
    tmMcfSrchKey.iCode = tlCif.iMcfCode
    ilRet = btrGetEqual(hmMcf, tlMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mExportCopyRotRecord = "Error getting MCF: " & ilRet
        Exit Function
    End If
    
    '------------------
    'Advertiser
    'tlcpr.iAdfCode
    'tgAdvertiser
    slString = ""
    If tlCrf.iAdfCode > 0 Then
        For illoop = 0 To UBound(tgAdvertiser)
            ilRet = gParseItem(tgAdvertiser(illoop).sKey, 2, "\", slvalue)
            If Val(slvalue) = tlCrf.iAdfCode Then
                ilRet = gParseItem(tgAdvertiser(illoop).sKey, 1, "\", slvalue)
                slString = Trim(slvalue)
                Exit For
            End If
        Next illoop
    End If
    slExportLine = """" & slString & """"
    
    '------------------
    'Contract
    slExportLine = slExportLine & "," & tlChf.lCntrNo
    
    '------------------
    'Rotation number
    slExportLine = slExportLine & "," & tlCrf.iRotNo
    
    '------------------
    'Vehicle
    'tmcpr.iVefCode
    'tgVehicle
    slString = ""
    If tlCrf.iVefCode > 0 Then
        For illoop = 0 To UBound(tgVehicle)
            ilRet = gParseItem(tgVehicle(illoop).sKey, 2, "\", slvalue)
            If Val(slvalue) = tlCrf.iVefCode Then
                ilRet = gParseItem(tgVehicle(illoop).sKey, 3, "|", slvalue)
                ilRet = gParseItem(slvalue, 1, "\", slvalue)
                slString = Trim(slvalue)
                Exit For
            End If
        Next illoop
    End If
    'Fix v81 TTP 10791 - copy rotations by advertiser report - test results
    If slString = "" And slVehicleNameString <> "" Then
        slExportLine = slExportLine & ",""" & slVehicleNameString & """"
    Else
        slExportLine = slExportLine & ",""" & slString & """"
    End If
    
    '------------------
    'Restrictions
    '{@IsItBB} + crfLen + " " + {@Dates} + " " + {@Times} + " " + {@Days} + " " + {@Avails} + " " + {@Zones}
    slString = ""
    '{@IsItBB}
    If tlCrf.sRotType = "O" Then slString = "OBB-"
    If tlCrf.sRotType = "C" Then slString = "CBB-"
    'crfLen
    slString = slString & tlCrf.iLen
    slString = slString & " "
    '{@Dates}
    gUnpackDate tlCrf.iStartDate(0), tlCrf.iStartDate(1), slDate
    slString = slString & slDate
    slString = slString & "-"
    gUnpackDate tlCrf.iEndDate(0), tlCrf.iEndDate(1), slDate
    slString = slString & slDate
    slString = slString & " "
    '{@Times}
    If (tlCrf.iStartTime(0) + tlCrf.iStartTime(1) <> 0) Or (tlCrf.iEndTime(0) + tlCrf.iEndTime(1)) <> 0 Then
        gUnpackTime tlCrf.iStartTime(0), tlCrf.iStartTime(1), "A", "1", slTime
        slString = slString & slTime
        slString = slString & "-"
        gUnpackTime tlCrf.iEndTime(0), tlCrf.iEndTime(1), "A", "1", slTime
        slString = slString & slTime
        slString = slString & " "
    End If
    '{@Days}
    slDate = ""
    If tlCrf.sDay(0) = "Y" And tlCrf.sDay(1) = "Y" And tlCrf.sDay(2) = "Y" And tlCrf.sDay(3) = "Y" And tlCrf.sDay(4) = "Y" And tlCrf.sDay(5) = "Y" And tlCrf.sDay(6) = "Y" Then
        slDate = ""
    Else
        If tlCrf.sDay(0) = "Y" Then
            slDate = slDate & "M"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(1) = "Y" Then
            slDate = slDate & "T"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(2) = "Y" Then
            slDate = slDate & "W"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(3) = "Y" Then
            slDate = slDate & "T"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(4) = "Y" Then
            slDate = slDate & "F"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(5) = "Y" Then
            slDate = slDate & "S"
        Else
            slDate = slDate & "x"
        End If
        If tlCrf.sDay(6) = "Y" Then
            slDate = slDate & "S"
        Else
            slDate = slDate & "x"
        End If
    End If
    slString = slString & slDate
    slString = slString & " "
    '{@Avails}
    slvalue = ""
    If (tlCrf.ianfCode = 0) Or (tlCrf.sInOut <> "I" And tlCrf.sInOut <> "O") Then
        slvalue = ""
    Else
        For illoop = 0 To UBound(tmAnfTable)
            If tmAnfTable(illoop).iCode = tlCrf.ianfCode Then
                slvalue = Trim(tmAnfTable(illoop).sName) & " avails only"
                Exit For
            End If
        Next illoop
    End If
    slString = slString & slvalue
    slString = slString & " "
    '{@Zones}
    slvalue = ""
    If Trim(tlCrf.sZone) = "R" Then
        slvalue = "Regional copy"
    Else
        If Trim(tlCrf.sZone) <> "" Then
            slvalue = tlCrf.sZone & " zone only"
        End If
    End If
    slString = slString & slvalue
    slExportLine = slExportLine & ",""" & Trim(slString) & """"
    
    '------------------
    'Entry Date
    gUnpackDate tlCrf.iEntryDate(0), tlCrf.iEntryDate(1), slDate
    slExportLine = slExportLine & ",""" & slDate & """"
    
    '------------------
    'Feed Date
    gUnpackDate tlCrf.iFeedDate(0), tlCrf.iFeedDate(1), slDate
    slExportLine = slExportLine & ",""" & slDate & """"
    
    '------------------
    'Feed Status
    slString = ""
    If tlCrf.sFeedStatus = "S" Then slString = "Sent"
    If tlCrf.sFeedStatus = "X" Then slString = "Sent"
    If tlCrf.sFeedStatus = "D" Then slString = "Defer"
    If tlCrf.sFeedStatus = "R" Then slString = "Ready"
    If tlCrf.sFeedStatus = "P" Then slString = "Suppress"
    slExportLine = slExportLine & ",""" & slString & """"
    
    '------------------
    'Earliest Date Spot
    'crfEarliestDateAssg
    gUnpackDate tlCrf.iEarliestDateAssg(0), tlCrf.iEarliestDateAssg(1), slDate
    slExportLine = slExportLine & ",""" & slDate & """"
    
    '------------------
    'Latest Assign
    'crfLatestDateAssg
    gUnpackDate tlCrf.iLatestDateAssg(0), tlCrf.iLatestDateAssg(1), slDate
    slExportLine = slExportLine & ",""" & slDate & """"
    
    '------------------
    'Last Assign Done
    slString = ""
    gUnpackDate tlCrf.iDateAssgDone(0), tlCrf.iDateAssgDone(1), slDate
    slString = slString & ",""" & slDate
    If Trim(slDate) <> "" Then
        slString = slString & " "
        gUnpackTime tlCrf.iTimeAssgDone(0), tlCrf.iTimeAssgDone(1), "A", "1", slTime
        slString = slString & slTime
    End If
    slString = slString & """"
    slExportLine = slExportLine & slString
    
    '------------------
    'Contract Product
    'chfProduct
    slExportLine = slExportLine & ",""" & Trim(tlChf.sProduct) & """"
    
    '------------------
    'Cart Number
    slExportLine = slExportLine & ",""" & gStripChr0(tlMcf.sName) & gStripChr0(tlCif.sName) & """"
    
    '------------------
    'ISCI
    'CPF_Copy_Prodct_ISCI.cpfISCI
    slExportLine = slExportLine & ",""" & gStripChr0(tlCpf.sISCI) & """"
    
    '------------------
    'Creative Title
    'cpfCreative
    slExportLine = slExportLine & ",""" & gStripChr0(tlCpf.sCreative) & """"
    
    '------------------
    'Inventory Product
    'cpfName
    slExportLine = slExportLine & ",""" & gStripChr0(tlCpf.sName) & """"
    
    Print #hmExport, slExportLine
End Function
