Attribute VB_Name = "RPTGEN"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptgen.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmCffSrchKey                  tmSmfSrchKey                                            *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  mGetEarliestCopyDate                                                                  *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptGen.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Rm**'Declare Sub ISortT2 Lib "QPRO200.DLL" (Array As Any, Index%, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%)
'Rm**'Declare Sub ArraySortTyp Lib "QPRO200.DLL" (Array() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%, ByVal CaseSenitive%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Type ODFEXT
'    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sZone As String * 3
'    iEtfCode As Integer         'Event type code
'    iEnfCode As Integer         'Event name code
'    sProgCode As String * 5 'Program code #
'    ianfCode As Integer 'Avail name code
'    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sProduct As String * 35 'Product (either from contract or copy)
'    iMnfSubFeed As Integer
'    iBreakNo As Integer 'Reset at start of each program
'    iPositionNo As Integer
'    lCefCode As Long
'    sShortTitle As String * 15
'End Type

'7-6-15 moved to rptrec so modules can be removed from traffic
'Type TYPESORT
'    sKey As String * 100
'    lRecPos As Long
'End Type

Type SPOTTYPESORT
    sKey As String * 80 'Office Advertiser Contract
    iVefCode As Integer 'line airing vehicle
    sCostType As String * 12    'string of spot type (0,00, bonus, adu, recapturable, etc)
    sDyWk As String * 1   'wkly vs daily
    iSpotsWk As Integer 'Spots per week, if zero, then daily
    iDay(0 To 6)  As Integer    'Spot per day if daily or flag if weekly
                                'For weekly 1=Air day; 0=not air day
                                'Index 0=Mo; 1=tu,...6=Su
    iXSpotsWk As Integer 'Spots per week, if zero, then daily
    sXDay(0 To 6)  As String * 1    'Spot per day if daily or flag if weekly
                                'For weekly 1=Air day; 0 or blank=not air day
                                'Index 0=Mo; 1=tu,...6=Su
    sLiveCopy As String * 1     '5-31-12 copy type L = Live cml, M = live promo, S = recorded promo, P = pre-recorded coml, Q = pre-recorded promo, Blank or R = recorded
    tSdf As SDF
End Type

'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
Type PCFTYPESORT
    sKey As String * 80         '
    lCntrNo As Long             'Contract Number
    lExtCntrNo As Long          'Ext Contract Number
    sContractType As String     'ContractType
    sCntrStartDate As String    'Contract Start Date
    sCntrEndDate As String      'Contract End Date
    iAgfCode As Integer         'Agency Code
    iAdfCode As Integer         'Advertiser Code
    sProduct As String          'Product
    
    iVefCode As Integer         'line vehicle
    sCalType As String          'Caledar Type (S=Standard, C=Calendar)
    tPcf As PCF
    lReceivables(0 To 25) As Long 'Receivables
End Type

Type COPYCNTRSORT
    sKey As String * 100 '3-10-00 add 20 due to veh name expansion ,Agency, City Contrct # ID Vehicle Len
    lChfCode As Long
    iVefCode As Integer
    sVehName As String * 40     '3-10-00 chged from 20 to 40, selling or (rotation) vehicle
    sAirVehName As String * 40  '3-10-00 chged from 20 to 40,airing vehicle
    iLen As Integer
    iNoSpots As Integer 'Number of spots that have no copy
    iNoUnAssg As Integer    'Number of spots not assigned
    iNoToReassg As Integer    'Number of spots should be reassigned
    iAsgnVefCode As Integer
    iNoSpotsMiss As Integer   'flag indicating that at least 1 spot missed for spots with no copy
    iNoUnAssgMiss As Integer   'flag indicating that at least 1 spot not assigned is missed
    iNoToReassgMiss As Integer  'flag indicating that at least 1 spot to reassign is missed
'8-18-00
    iRegionNoSpots As Integer 'Number of spots that have no copy
    iRegionNoUnAssg As Integer    'Number of spots not assigned
    iRegionNoToReassg As Integer    'Number of spots should be reassigned
    iRegionNoSpotsMiss As Integer   'flag indicating that at least 1 spot missed for spots with no copy
    iRegionNoUnAssgMiss As Integer   'flag indicating that at least 1 spot not assigned is missed
    iRegionNoToReassgMiss As Integer  'flag indicating that at least 1 spot to reassign is missed
    lFsfCode As Long                '8-2-04 feed spot code
    sLiveFlag As String * 1         '11-16-05 L = Live, M = both Live & recorded, Blank = recorded
    iRdfCode As Integer             '1-05-06 show DP on Contracts Missing Copy
    iLineNo As Integer              '2-28-07
    iStartDate(0 To 1) As Integer   '2-28-07 contract hdr start date or line start day date
    iEndDate(0 To 1) As Integer     '2-28-07 contract hdr end date or line end date
End Type
'7-6-15 moved to copy.bas so some modules can be removed from traffic
'Type COPYROTNO
'    iRotNo As Integer
'    sZone As String * 3
'End Type
Type COPYSORT
    sKey As String * 100 '3-10-00 added 20 due to veh name expansion,Agency, City ID
    iCopyStatus As Integer '0=No Copy; 1=Assigned; 2=Copy but not assigned; 3= Supersede; 4=Zone missing
    '8-11-00
    iRegionalStatus As Integer  '0 = no warnings or errors, 1 =not assigned (no regional ever assigned), 2 = regional superseded (assigned before),
                                '3 = non-regional copy not found , 4 = rotation defined for contract other than spot
                                'contract that is valid for date of spot and vehicle
    iRegionalSort As Integer    '0 = nonregion, 1 = region; used to blank out fields  when outputting regional copy lines

    tSdf As SDF
    sVehName As String * 40   '3-10-00 chged from 20 to 40,
End Type
Type CODESTNCONV
    sName As String * 20
    sCodeStn As String * 5
End Type
Type DALLASFDSORT
    sKey As String * 30
    sRecord As String * 104
End Type
Dim tmPLSdf() As SPOTTYPESORT
Dim tmCopyCntr() As COPYCNTRSORT
Dim tmCopy() As COPYSORT
Dim imSellVefSelected() As Integer      '7-16-14
Dim tmSelAdvt() As Integer
Dim imNoZones As Integer
'Dim tmRotNo(1 To 6) As COPYROTNO
Dim tmRotNo(0 To 6) As COPYROTNO    'Index zero ignored
Dim tmCodeStn() As CODESTNCONV
Dim tmDallasFdSort() As DALLASFDSORT
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0            'ANF record image
Dim imAnfRecLen As Integer        'ANF record length
Dim hmCef As Integer            'Event comments file handle
Dim tmCef As CEF                'CEF record image
Dim tmCefSrchKey As LONGKEY0            'CEF record image
Dim imCefRecLen As Integer        'CEF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            '11-16-05 CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0            'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract line flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF
Dim hmVsf As Integer            'Vehicle combo file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer            'Agency name file handle
Dim tmAgf As AGF                'AGF record image
Dim tmAgfSrchKey As INTKEY0            'AGF record image
Dim imAgfRecLen As Integer        'AGF record length
Dim hmSmf As Integer            'MG and outside Times file handle
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey2 As LONGKEY0   'smf key 0
Dim imSmfRecLen As Integer      'SMF record length
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF
'Short Title
Dim hmSif As Integer        'Short Title file handle
Dim tmSif As SIF            'SIF record image
Dim imSifRecLen As Integer     'SIF record length

'Copy rotation vehicle table
Dim hmCvf As Integer        'Copy rotation vehicle handle
Dim tmCvf As CVF            'CVF record image
Dim imCvfRecLen As Integer     'CVF record length

'Copy rotation
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim tmCrfSrchKey4 As CRFKEY4
Dim imCrfRecLen As Integer     'CRF record length

Type CRFBYCNTR
    sKey As String * 20
    tCrf As CRF
End Type

Dim tmCRFByCntr() As CRFBYCNTR


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
'  Copy CPR pre-pass file
Dim hmCpr As Integer        'copy pre-pass file
Dim tmCpr As CPR            'CPR record image
Dim imCprRecLen As Integer     'CPR record length
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
'  Library calendar File
Dim hmLcf As Integer        'Library calendar file handle
Dim tmLcf As LCF            'LCF record image
Dim imLcfRecLen As Integer     'LCF record length
Dim tmLcfSrchKey0 As LCFKEY0

Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVLF As Integer            'Vehiclee file handle
'Dim tmVlf As VLF                'VEF record image
Dim tmVlf() As VLF
Dim tmSlf As SLF                'SLF record image
Dim hmMnf As Integer            'MultiName file handle
Dim tmMnf As MNF                'MNF record image
Dim imMnfRecLen As Integer        'MNF record length


Dim hmFsf As Integer            'Feed spot file handle
Dim tmFSFSrchKey As LONGKEY0     'FSF record image
Dim imFsfRecLen As Integer       'FSF record length
Dim tmFsf As FSF

Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey As LONGKEY0     'PrF record image
Dim imPrfRecLen As Integer       'PrF record length
Dim tmPrf As PRF

'8-11-2000 Regional copy
Dim imNonRegionDefined As Integer
Dim imRegionMissing As Integer
Dim imRegionSuperseded As Integer
Dim hmRaf As Integer            'Regional areas
Dim tmRaf As RAF
Dim tmRafSrchKey1 As RAFKEY1
Dim imRafRecLen As Integer        'RAF record length
Dim hmRsf As Integer              'Regional copy assignment
Dim tmRsf As RSF
Dim tmRsfSrchKey1 As LONGKEY0
Dim imRsfRecLen As Integer        'RAF record length

'Copy Air Game
Dim tmCaf As CAF            'CAF record image
Dim tmCafSrchKey As LONGKEY0  'CAF key record image
Dim tmCafSrchKey1 As CAFKEY1  'CAF key record image
Dim hmCaf As Integer        'CAF Handle
Dim imCafRecLen As Integer      'CAF record length
'Game schedule
Dim hmGsf As Integer
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length
Dim tmGsf As GSF

'******
Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim lmTodayDate As Long             'todays date (NOW)
'Dim tmOdf0() As ODFEXT
'Dim tmOdf1() As ODFEXT
'Dim tmOdf2() As ODFEXT
'Dim tmOdf3() As ODFEXT
'Dim tmOdf4() As ODFEXT
'Dim tmOdf5() As ODFEXT
'Dim tmOdf6() As ODFEXT

Dim tmTxr As TXR
Dim hmTxr As Integer
Dim imTxrRecLen As Integer        'TXR record length

Type TXRLOG                     'this image aligns with TXR layout. Starting with Stime it overlaps the TXR Text field string of 200 bytes
    iFiller1(0 To 1) As Integer 'Filler -retain        (Byte 1-4)
    iFiller2(0 To 1) As Integer 'Filler -retain        (byte 5-8)
    lFiller3 As Long              'Filler - retain     (byte 9-12)
    iFiller4 As Integer           'filler - retain     (byte 13-14)
    'starting here parallels txr.stext for 200 bytes
    sTime As String * 10        'XX:XX:XXAM/PM         (byte 1-10)
    'sCopy As String * 6         'CXXXXX                (byte 11-16)
    sCopy As String * 10         'CCCCCXXXXX            (byte 11-16)   -----> (byte 11-20)
    sLen As String * 4          'XXXX                  (byte 17-20)    -----> (byte 21-24)
    sAdvtProd As String * 71    'advt(35),short title(35)  (byte 21-91) ----> (byte 25-95)
    sVehicle As String * 40                            '(byte 92-131) ------> (byte 96-135)
    sTimeKey As String * 6      'time if hhmmss (military (byte 132-137) ---> (byte 136-141)
    sLogDate As String * 20     'Day of week followed by date  (byte 138-157) ----> (byte 142-161)
    'sFiller5 As String * 43                             '(byte 158-200)
    sFiller5 As String * 39     'fILLER                   (byte 158-200) ---------> (BYTE 162 - 200
    sFiller6 As Long            'csfcode
    sUnused As String * 20
End Type

'4-28-11    show day is complete status on Log Posting Status
Type DAYISCOMPLETE
    iVefCode As Integer
    iGameNo As Integer
    iDate(0 To 1) As Integer
    sAffPost As String * 1      'posted flag from LCF
End Type

Dim tmDayIsComplete() As DAYISCOMPLETE
Dim tmTxrLog As TXRLOG
Dim lmCntrCode As Long          '11-16-05

Dim lgMainCvfCount As Long
Dim lgRegionCvfCount As Long
Dim lgSupercedeCount As Long
'
'           Find the LCF for the Summary version of Log Posting Status report
'           Games will go thru the spots to get all the games for each day
'
Public Sub mFindNonGamePostStatus(slStartDate As String, slEndDate As String, ilVefCode As Integer)
Dim llStartDate As Long
Dim llEndDate As Long
Dim ilDate(0 To 1) As Integer
Dim llLoopOnDate As Long
Dim ilRet As Integer

        llStartDate = gDateValue(slStartDate)
        llEndDate = gDateValue(slEndDate)
        
        For llLoopOnDate = llStartDate To llEndDate
            gPackDateLong llLoopOnDate, ilDate(0), ilDate(1)
            'get the day is complete status from LCF
            tmLcfSrchKey0.iLogDate(0) = ilDate(0)
            tmLcfSrchKey0.iLogDate(1) = ilDate(1)
            tmLcfSrchKey0.iSeqNo = 1
            tmLcfSrchKey0.iType = 0         'not a game
            tmLcfSrchKey0.iVefCode = ilVefCode
            tmLcfSrchKey0.sStatus = "C"
            ilRet = btrGetEqual(hmLcf, tmLcf, Len(tmLcf), tmLcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCpr.lGenTime = lgNowTime
            tmCpr.iGenDate(0) = igNowDate(0)
            tmCpr.iGenDate(1) = igNowDate(1)
            tmCpr.iVefCode = ilVefCode
            tmCpr.iSpotDate(0) = ilDate(0)
            tmCpr.iSpotDate(1) = ilDate(1)
            tmCpr.iRemoteID = 0             'not a game
            tmCpr.sLive = "N"               'assume not posted if lcf not found
            If ilRet = BTRV_ERR_NONE Then
                tmCpr.sLive = tmLcf.sAffPost
                ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
            End If
     
        Next llLoopOnDate
        Exit Sub
End Sub
'************************************************************************
'*                                                                      *
'*      Procedure Name:gCopyAdvtRpt                                     *
'*                                                                      *
'*             Created:4/21/94       By:D. LeVine                       *
'*            Modified:              By:                                *
'*                                                                      *
'*            Comments: Generate Copy by Advertiser                     *
'*                                                                      *
'*     8/28/97 dh Nothing printing due to wrong list box                *
'*             being tested as well as list box not even                *
'*             populated with vehicles                                  *
'*                                                                      *
'*     9/25/00 DS Coverted to Crystal from Bridge                       *
'*      3-23-03 Determine by advt whether to show on inv or not.  All
'*              fills/extras changed to +fill (show on inv) or -fill (do
'*              not show on inv)
'*           7-27-04 Add option to include/exclude contract/feed spots
'*     11-30-04 change accessing smf from key0 to key2 for speed
'*      dh 3-29-05 create BB spots if not yet created to test if copy exists
'************************************************************************
Sub gCopyAdvtRpt()
    Dim slFileInError As String
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBRet As Integer
    'Dim ilDummy As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llIndex As Long         '7-17-09 chged from integer to long
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim ilSpotType As Integer
    Dim ilAdvt As Integer
    Dim slCopyProduct As String
    Dim slCopyZone As String
    Dim slCopyISCI As String
    Dim slCopyCart As String
    Dim ilIncludeUnassigned As Integer
    Dim tlVef As VEF
    Dim slProdOrShortT As String
    Dim slName As String
    Dim slShowOnInv As String * 1
    Dim ilContrSpots As Integer
    Dim ilFeedSpots As Integer
    Dim slChfFsfProduct As String
    Dim slSaveAdvName  As String
    Dim llStartDate As Long         '3-29-05 loop on dates requested by vehicle to create bb spots
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slDate As String

    Screen.MousePointer = vbHourglass
'    slStartDate = RptSel!edcSelCFrom.Text   'Start date
'    slEndDate = RptSel!edcSelCTo.Text   'End date
'   8-22-19 use csi calendar control vs edit box
    slStartDate = RptSel!CSI_CalFrom.Text   'Start date
    slEndDate = RptSel!CSI_CalTo.Text   'End date
    slDateRange = "From " & slStartDate & " To " & slEndDate
    If RptSel!rbcSelCSelect(0).Value Then  'All spots
        ilSpotType = 0
        slDateRange = slDateRange & " for All Spots"
    ElseIf RptSel!rbcSelCSelect(1).Value Then  'Only spots with copy
        ilSpotType = 1
        slDateRange = slDateRange & " for Spots with Copy"
    Else    'Only spots without copy
        ilSpotType = 2
        slDateRange = slDateRange & " for Spots without Copy"
    End If
    If Not gSetFormula("ShowDateBanner", "'" & slDateRange & "'") Then
        Exit Sub
    End If
    If ilSpotType = 0 Then
        ilIncludeUnassigned = True
    Else
        If RptSel!rbcSelCInclude(0).Value Then
            ilIncludeUnassigned = True
        Else
            ilIncludeUnassigned = False
        End If
    End If

    ilContrSpots = True                                     'assume to include both contract & feed spots
    ilFeedSpots = True
    If RptSel!ckcAll.Value = vbChecked Then
        If Not RptSel!ckcSelC10(0).Value = vbChecked Then       'include Contracts spots?
            ilContrSpots = False
        End If
        If Not RptSel!ckcSelC10(1).Value = vbChecked Then       'include feed spots?
            ilFeedSpots = False
        End If
    Else                                                    'selected advt, see if feed and/or contracts selected
        If RptSel!lbcSelection(5).Selected(0) Then        'selected feed?  (first entry)
            ilFeedSpots = True
        Else
            ilFeedSpots = False
        End If
        'no need to test for selected contract because if not selected wont match the contract code against spot
    End If



    slFileInError = mOpenCopyStatusFiles()         'open all applicable files
    If slFileInError <> "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Error opening " & Trim$(slFileInError) & " - Rptgen: gCopyAdvtRpt"
        Exit Sub
    End If

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)

    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCpfRecLen = Len(tmCpf)

    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMcfRecLen = Len(tmMcf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmSmf
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)

'    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSdf)
'        ilRet = btrClose(hmSmf)
'        ilRet = btrClose(hmMcf)
'        ilRet = btrClose(hmCpf)
'        ilRet = btrClose(hmCif)
'        btrDestroy hmSmf
'        btrDestroy hmMcf
'        btrDestroy hmCpf
'        btrDestroy hmCif
'        btrDestroy hmSdf
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If

    llStartDate = gDateValue(slStartDate)       '3-29-05 convert string date to long for looping
    llEndDate = gDateValue(slEndDate)

    tmAgf.iCode = 0
    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0
    If tgSpf.sUseProdSptScr = "P" Then    'use short title vs contr hdr product
        slProdOrShortT = "Advertiser, Short Title"
    Else
        slProdOrShortT = "Advertiser, Product"
    End If
    If Not gSetFormula("ShowAdvProdOrShort", "'" & slProdOrShortT & "'") Then
        mCloseCopyFiles
        btrDestroy hmSmf
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        'btrDestroy hmSdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    DoEvents
    ReDim tmSelAdvt(0 To 0) As Integer
    For ilAdvt = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
        If RptSel!lbcSelection(0).Selected(ilAdvt) Then
            slNameCode = tgAdvertiser(ilAdvt).sKey 'Traffic!lbcAdvertiser.List(ilAdvt)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmSelAdvt(UBound(tmSelAdvt)) = Val(slCode)
            ReDim Preserve tmSelAdvt(0 To UBound(tmSelAdvt) + 1) As Integer
        End If
    Next ilAdvt
    ReDim tmCopy(0 To 0) As COPYSORT
    For ilVehicle = 0 To RptSel!lbcSelection(6).ListCount - 1 Step 1
        slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        ilVpfIndex = -1
        'For ilLoop = 0 To UBound(tgVpf) Step 1
        '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
            illoop = gBinarySearchVpf(ilVefCode)
            If illoop <> -1 Then
                ilVpfIndex = illoop
        '        Exit For
            End If
        'Next ilLoop

        '3-29-05 go out and create all BB spots if necessary for dates requested
        For llDate = llStartDate To llEndDate
            slDate = Format(llDate, "m/d/yy")
            ilRet = gCreateBBSpots(hmSdf, ilVefCode, slDate)
        Next llDate
        mObtainCopyDate 1, ilVefCode, ilVpfIndex, slStartDate, slEndDate, ilSpotType, ilIncludeUnassigned, ilContrSpots, ilFeedSpots
    Next ilVehicle
    'outer loop - one loop per page
    llIndex = LBound(tmCopy)
    If llIndex >= UBound(tmCopy) Then
        ilDBRet = 1
    If ilSpotType = 2 Then  'Nothing to display. We can send one record with 32000 in
        'the veh code to cause Crystal to print the **** NONE **** ,but it causes bad
        'time, length etc values to display
        'tmCpr.iGenTime(0) = igNowTime(0)
        'tmCpr.iGenTime(1) = igNowTime(1)
        'tmCpr.iGenDate(0) = igNowDate(0)
        'tmCpr.iGenDate(1) = igNowDate(1)
        'tmCpr.iVefCode = 32000  '
        'ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
    End If
    Else
        'Sort key
        ArraySortTyp fnAV(tmCopy(), 0), UBound(tmCopy), 0, LenB(tmCopy(0)), 0, LenB(tmCopy(0).sKey), 0
        ilDBRet = BTRV_ERR_NONE
    End If
    While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0 'VB6** And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
        'While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0 And ilRet = 0 '5-12-05 remove, VB6** And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
            tmSdf = tmCopy(llIndex).tSdf
            'tmCpr.iGenTime(0) = igNowTime(0)
            'tmCpr.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCpr.lGenTime = lgNowTime
            tmCpr.iGenDate(0) = igNowDate(0)
            tmCpr.iGenDate(1) = igNowDate(1)
            If tmCopy(llIndex).iRegionalSort = 0 Then           '8-16-00 spot line (vs regional copy line_
                tmCpr.iSpotDate(0) = tmSdf.iDate(0)
                tmCpr.iSpotDate(1) = tmSdf.iDate(1)
                If tmSdf.iVefCode <> tmVef.iCode Then
                    tmVefSrchKey.iCode = tmSdf.iVefCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        tmVef.sName = "Missing"
                    End If
                End If
                tmCpr.iVefCode = tmSdf.iVefCode
                tmCpr.iSpotTime(0) = tmSdf.iTime(0)
                tmCpr.iSpotTime(1) = tmSdf.iTime(1)
                tmCpr.iLen = tmSdf.iLen

                If tmSdf.sSpotType = "O" Then           '3-29-05 set flag to show bb on report
                    tmCpr.lFt2CefCode = 1
                ElseIf tmSdf.sSpotType = "C" Then       'closed bb
                    tmCpr.lFt2CefCode = 2
                Else
                    tmCpr.lFt2CefCode = 0
                End If

                slChfFsfProduct = ""
                illoop = gBinarySearchAdf(tmSdf.iAdfCode)       'find the advertiser record
                slSaveAdvName = ""
                If illoop <> -1 Then
                    'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                    '    slSaveAdvName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                    'Else
                        slSaveAdvName = Trim$(tgCommAdf(illoop).sName)
                    'End If
                End If


               'obtain feed spot or contract header for product
                If tmSdf.lChfCode = 0 Then              'feed spot
                    tmFSFSrchKey.lCode = tmSdf.lFsfCode
                    ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        tmPrfSrchKey.lCode = tmFsf.lPrfCode
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            slChfFsfProduct = Trim$(tmPrf.sName)
                        End If
                        mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct
                        If Trim$(slCopyProduct) = "" Then
                            slStr = Trim$(slSaveAdvName) & "," & Trim$(slChfFsfProduct)
                        Else
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slCopyProduct)
                        End If
                        tmCpr.lCntrNo = 0
                        tmCpr.iLineNo = 0
                        tmCpr.lFt1CefCode = tmFsf.lCode
                    End If
                Else
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slChfFsfProduct = Trim$(tmChf.sProduct)
                    mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct

                    If Trim$(slCopyProduct) = "" Then         'no copy found, get short title (if applicable from chf)
                        If tgSpf.sUseProdSptScr = "P" Then    'use short title vs contr hdr product
                            slStr = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slStr)
                        Else
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slChfFsfProduct)
                        End If
                    Else                      'copy found, proper short title or product retrieved
                        slStr = Trim$(slSaveAdvName) & ", " & Trim$(slCopyProduct)
                    End If
                    tmCpr.lCntrNo = tmChf.lCntrNo
                    tmCpr.iLineNo = tmSdf.iLineNo

                End If
                tmCpr.sProduct = slStr
                If tmCopy(llIndex).iCopyStatus = 0 Then 'None defined
                    tmCpr.sZone = ""
                    tmCpr.sCartNo = ""
                    tmCpr.sISCI = ""
                ElseIf tmCopy(llIndex).iCopyStatus = 2 Then 'Defined, not assigned
                    tmCpr.sZone = ""
                    tmCpr.sCartNo = "*"
                    tmCpr.sISCI = ""
                Else    'Copy exist
                    tmCpr.sZone = slCopyZone
                    If tmCopy(llIndex).iCopyStatus = 1 Then 'Ok
                        tmCpr.sCartNo = slCopyCart
                    ElseIf tmCopy(llIndex).iCopyStatus = 3 Then 'Superseded
                        tmCpr.sCartNo = "^" & slCopyCart
                    ElseIf tmCopy(llIndex).iCopyStatus = 4 Then 'Zone missing
                        tmCpr.sCartNo = "~" & slCopyCart
                    End If
                    tmCpr.sISCI = slCopyISCI
                End If
                If tmCopy(llIndex).iRegionalStatus = 1 Then     '8-16-00
                     tmCpr.sStatus = "W"         'other cnts for same advt have regional copy defined
                    'ilDummy = LLDefineFieldExt(hdJob, "Region", slStr, LL_TEXT, "")
                ElseIf tmCopy(llIndex).iRegionalStatus = 2 Then
                    tmCpr.sStatus = "*"
                    'ilDummy = LLDefineFieldExt(hdJob, "Region", slStr, LL_TEXT, "")
                ElseIf tmCopy(llIndex).iRegionalStatus = 3 Then
                    tmCpr.sStatus = "^"
                    'ilDummy = LLDefineFieldExt(hdJob, "Region", slStr, LL_TEXT, "")
                ElseIf tmCopy(llIndex).iRegionalStatus = 4 Then     'ok
                    tmCpr.sStatus = ""
                    'ilDummy = LLDefineFieldExt(hdJob, "Region", "", LL_TEXT, "")
                Else                'none defined
                    'ilDummy = LLDefineFieldExt(hdJob, "Region", "", LL_TEXT, "")
                    tmCpr.sStatus = ""
                End If
                tmSmf.iOrigSchVef = tmSdf.iVefCode
                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    '11-30-04 change reading of smf from key0 to key2 for speed
                    'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                    'tmSmfSrchKey.lFsfCode = tmSdf.lFsfCode
                    'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                    'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                    'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation

                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation

                    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo) And (tmSmf.lFsfCode = tmSdf.lFsfCode)
                        If (tmSmf.lSdfCode = tmSdf.lCode) Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    Loop
                End If
                If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                    tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                If tmSdf.sSchStatus = "S" Then
                    If tmSdf.sSpotType <> "X" Then
                        slStr = "Scheduled"
                    Else                            'extra or fill?
                        'If tmSdf.sPriceType = "N" Then
                        '3-23-03 Test the advt instead of spot to determine a fill or extra
                        '1-19-04 change way in which fill/extra are shown.  If spot price type
                        'isnt a "-" and "+", then use the advt to see how to show; otherwise use
                        'the spot price type
                        slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)
                        If slShowOnInv = "N" Then
                        'If tmAdf.sBonusOnInv = "N" Then
                            slStr = "-Schd Fill"
                        Else
                            'slStr = "+Schd Extra"
                            slStr = "+Schd Fill"
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "M" Then
                    slStr = "Missed"
                ElseIf tmSdf.sSchStatus = "R" Then
                    slStr = "Ready"
                ElseIf tmSdf.sSchStatus = "U" Then
                    slStr = "UnSched"
                ElseIf tmSdf.sSchStatus = "G" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                        Else
                            'If tmSdf.sPriceType = "N" Then          'filled makegood (s/n happen!)
                            '3-23-03 Test the advt instead of spot to determine a fill or extra

                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)
                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                            Else
                                'slStr = "+Extra Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                                slStr = "+Fill Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                            End If
                        End If
                    Else
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Makegood"
                        Else                            'extra or fill?
                            'If tmSdf.sPriceType = "N" Then  'fill
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Extra Fill"
                            Else
                                'slStr = "+Extra Makegood"
                                 slStr = "+Fill Makegood"
                            End If
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "O" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Outside" & Chr$(10) & Trim$(tlVef.sName)
                        Else                    'some form of extra
                            'If tmSdf.sPriceType = "N" Then  'outside fill (s/n happen!)
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Outside" & Chr$(10) & Trim$(tlVef.sName)
                            Else
                                'slStr = "+Extra Outside" & Chr$(10) & Trim$(tlVef.sName)
                                slStr = "+Fill Outside" & Chr$(10) & Trim$(tlVef.sName)
                            End If
                        End If
                        'slStr = "Outside" & Chr$(10) & Trim$(tlVef.sName)
                    Else
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Outside"
                        Else
                            'If tmSdf.sPriceType = "N" Then  'outside fill (s/n happen!)
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Outside"
                            Else
                                'slStr = "+Extra Outside"
                                slStr = "+Fill Outside"
                            End If
                        End If
                        'slStr = "Outside"
                    End If
                ElseIf tmSdf.sSchStatus = "C" Then
                    slStr = "Cancelled"
                ElseIf tmSdf.sSchStatus = "H" Then
                    slStr = "Hidden"
                ElseIf tmSdf.sSchStatus = "A" Then
                    slStr = "On Alt"
                ElseIf tmSdf.sSchStatus = "B" Then
                    slStr = "On Alt & MG"
                End If
                tmCpr.sCreative = slStr
                tmCpr.lHd1CefCode = 1
                ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)

                DoEvents
                If tmSdf.sSchStatus = "G" Then          'show where makegood came from
                    tmCpr.sCartNo = " "
                    tmCpr.sISCI = " "
                    tmCpr.sCreative = " "
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slStr
                    tmCpr.sISCI = gAddDayToDate(slStr)
                    gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slStr
                    tmCpr.sCreative = slStr

                    If tmSmf.iOrigSchVef <> tmVef.iCode Then
                        tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmVef.sName = "Missing"
                        End If
                    End If
                    tmCpr.sCartNo = tmVef.sName
                    tmCpr.lHd1CefCode = 2
                    ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                    tmCpr.sCartNo = " "
                    tmCpr.sISCI = " "
                    tmCpr.sCreative = " "
                End If
            Else                        'regional copy line
                mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct
                tmCpr.sCartNo = slCopyCart
                tmCpr.sISCI = slCopyISCI
                tmCpr.lHd1CefCode = 3
                ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                DoEvents
            End If

            'next data set if no error or warning
            If ilRet = 0 Then
                llIndex = llIndex + 1
                llRecNo = llRecNo + 1
                If llIndex >= UBound(tmCopy) Then
                    ilDBRet = 1
                End If
            End If
        'Wend  ' inner loop     5-12-05 remove
    Wend    ' while not EOF

    Erase tmSelAdvt
    Erase tmCopy
    Screen.MousePointer = vbDefault
    mCloseCopyFiles
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmMcf)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmCif)
    btrDestroy hmSmf
    btrDestroy hmMcf
    btrDestroy hmCpf
    btrDestroy hmCif

    Exit Sub
End Sub
'
'
'*******************************************************************
'*
'*      Procedure Name:gCopyCntrRpt
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Generate Contract Missing Copy
'*                      report
'*
'*      DH 7/28/98 Sort by Advertiser, not agency
'*      DH 8/23/99 Show missing copy for airing vehicles
'*      dh 12/7/99 Convert to Crystal from "bridge"
'*      dh 8/17/00 add regional copy feature
'*      dh 8-3-04 Option to include/exclude contract/station spots
'*      dh 3-29-05 create BB spots if not yet created to test if copy exists
'       dh 5-18-05 remove option to exclude missing copy - which is the
'           entire intent of this report
'       dh 1-05-06 Add option to show Daypart name on second line
'********************************************************************
Sub gCopyCntrRpt()
    Dim ilRet As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilIndex As Integer
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim ilIncludeUnAssg As Integer
    Dim ilIncludeReassg As Integer
    Dim slMissingFlag As String * 1
    Dim slUnassignFlag As String * 1
    Dim slReadyFlag As String * 1
    Dim ilInclCntrSpots As Integer
    Dim ilInclFeedSpots As Integer
    Dim slFileInError As String
    Dim llStartDate As Long         '3-29-05 loop on dates requested by vehicle to create bb spots
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim llCntrNo As Long            '11-16-05 option for single cntr #
    Dim ilVefIndex As Integer       '7-16-14

    Screen.MousePointer = vbHourglass
'    slStartDate = RptSel!edcSelCFrom.Text   'Start date
'    slEndDate = RptSel!edcSelCTo.Text   'End date
'   8-22-19 use csi calendar control vs edit box
    slStartDate = RptSel!CSI_CalFrom.Text   'Start date
    slEndDate = RptSel!CSI_CalTo.Text   'End date

    ilIncludeUnAssg = gSetCheck(RptSel!ckcSelC3(0).Value)   '9-12-02 = vbChecked
    ilIncludeReassg = gSetCheck(RptSel!ckcSelC3(1).Value)   '9-12-02 = vbChecked
    ilInclCntrSpots = gSetCheck(RptSel!ckcSelC10(0).Value)   'Include contract spots
    ilInclFeedSpots = gSetCheck(RptSel!ckcSelC10(1).Value)   'include feed spots

    slFileInError = mOpenCopyStatusFiles()         'open all applicable common copy files
    If slFileInError <> "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Error opening " & Trim$(slFileInError) & " - Rptgen: gCopyAdvtRpt"
        Exit Sub
    End If


    ' files required for this report in addition to the files opened in general routine
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmVLF)
        btrDestroy hmVLF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmSmf)
        btrDestroy hmVLF
        btrDestroy hmSmf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If RptSel!rbcSelCInclude(1).Value = True Then           'show line info
        hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseCopyFiles
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmVLF)
            ilRet = btrClose(hmSmf)
            btrDestroy hmCff
            btrDestroy hmVLF
            btrDestroy hmSmf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If



    llStartDate = gDateValue(slStartDate)       '3-29-05 convert string date to long for looping
    llEndDate = gDateValue(slEndDate)

    llCntrNo = 0                'ths is for debugging on a single contract
    slStr = RptSel!edcCheck
    If slStr <> "" Then
        llCntrNo = Val(slStr)
    End If

    lmCntrCode = 0
    'test for valid single contract if entered
    If llCntrNo <> 0 Then
        tmChfSrchKey1.lCntrNo = llCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_END_OF_FILE Or tmChf.lCntrNo <> llCntrNo Then
            'MsgBox "Contract # does not exist"
            mCloseCopyFiles
            btrDestroy hmCvf
            btrDestroy hmVLF
            btrDestroy hmSmf
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            lmCntrCode = tmChf.lCode
        End If
    End If

    tmSsf.iVefCode = 0
    tmAgf.iCode = 0
    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0
    
    lgMainCvfCount = 0
    lgRegionCvfCount = 0
    lgSupercedeCount = 0

    ReDim tmCopyCntr(0 To 0) As COPYCNTRSORT
    ReDim imSellVefSelected(0 To 0) As Integer
    For ilVehicle = 0 To RptSel!lbcSelection(6).ListCount - 1 Step 1
        If RptSel!lbcSelection(6).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilVefIndex = gBinarySearchVef(ilVefCode)
            If ilVefIndex <> -1 Then
                If tgMVef(ilVefIndex).sType = "S" Then              'keep track of all the selling vehicles selected.  if an airing vehicle is processed, make sure its not
                                                                    'processed twice or it has been selected
                    imSellVefSelected(UBound(imSellVefSelected)) = ilVefCode
                    ReDim Preserve imSellVefSelected(LBound(imSellVefSelected) To UBound(imSellVefSelected) + 1) As Integer
                End If
            End If
        End If
    Next ilVehicle
    
    For ilVehicle = 0 To RptSel!lbcSelection(6).ListCount - 1 Step 1
        If RptSel!lbcSelection(6).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            
            ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                illoop = gBinarySearchVpf(ilVefCode)
                If illoop <> -1 Then
                    ilVpfIndex = illoop
            '        Exit For
                End If
            'Next ilLoop

            '3-29-05 go out and create all BB spots if necessary for dates requested
            For llDate = llStartDate To llEndDate
                slDate = Format(llDate, "m/d/yy")
                ilRet = gCreateBBSpots(hmSdf, ilVefCode, slDate)
            Next llDate

            mObtainCopyCntr ilVefCode, slName, ilVpfIndex, slStartDate, slEndDate, ilInclCntrSpots, ilInclFeedSpots
        End If
    Next ilVehicle
    If (Not ilIncludeUnAssg) Or (Not ilIncludeReassg) Then
        For illoop = UBound(tmCopyCntr) - 1 To 0 Step -1
            If (Not ilIncludeUnAssg) And (Not ilIncludeReassg) Then
                If tmCopyCntr(illoop).iNoSpots = 0 Then
                    For ilIndex = illoop To UBound(tmCopyCntr) - 2 Step 1
                        tmCopyCntr(ilIndex) = tmCopyCntr(ilIndex + 1)
                    Next ilIndex
                    ReDim Preserve tmCopyCntr(0 To UBound(tmCopyCntr) - 1) As COPYCNTRSORT
                End If
            ElseIf Not ilIncludeUnAssg Then
                If (tmCopyCntr(illoop).iNoSpots = 0) And (tmCopyCntr(illoop).iNoToReassg = 0) Then
                    For ilIndex = illoop To UBound(tmCopyCntr) - 2 Step 1
                        tmCopyCntr(ilIndex) = tmCopyCntr(ilIndex + 1)
                    Next ilIndex
                    ReDim Preserve tmCopyCntr(0 To UBound(tmCopyCntr) - 1) As COPYCNTRSORT
                End If
            Else
                If (tmCopyCntr(illoop).iNoSpots = 0) And (tmCopyCntr(illoop).iNoUnAssg = 0) Then
                    For ilIndex = illoop To UBound(tmCopyCntr) - 2 Step 1
                        tmCopyCntr(ilIndex) = tmCopyCntr(ilIndex + 1)
                    Next ilIndex
                    ReDim Preserve tmCopyCntr(0 To UBound(tmCopyCntr) - 1) As COPYCNTRSORT
                End If
            End If
        Next illoop
    End If
    'Sort key
    'ArraySortTyp fnAV(tmCopyCntr(),0), UBound(tmCopyCntr), 0, LenB(tmCopyCntr(0)), 0, Len(tmCopyCntr(0).sKey), 0
    tmCpr.lHd1CefCode = 0           '2-19-15 init for running seq # to make sure dp name record is sorted properly
    For ilIndex = LBound(tmCopyCntr) To UBound(tmCopyCntr) - 1

        If tmCopyCntr(ilIndex).lChfCode = 0 Then
            tmFSFSrchKey.lCode = tmCopyCntr(ilIndex).lFsfCode
            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation

            If tmFsf.lPrfCode = 0 Then
                tmCpr.sProduct = ""
            Else
                tmPrfSrchKey.lCode = tmFsf.lPrfCode
                ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                tmCpr.sProduct = Trim$(tmPrf.sName)
            End If
            tmCpr.iAdfCode = tmFsf.iAdfCode
            tmCpr.iVefCode = 0
            tmCpr.iLineNo = 0
            tmCpr.lCntrNo = 0
            tmCpr.lFt1CefCode = tmFsf.lCode
            tmCpr.iRemoteID = 0                 '1-5-06     no daypart association
            'setup common fields in the contract header
            tmChf.iStartDate(0) = tmFsf.iStartDate(0)
            tmChf.iStartDate(1) = tmFsf.iStartDate(1)
            tmChf.iEndDate(0) = tmFsf.iEndDate(0)
            tmChf.iEndDate(1) = tmFsf.iEndDate(1)
        Else
           tmChfSrchKey.lCode = tmCopyCntr(ilIndex).lChfCode
           ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation

           'If tmChf.iAgfCode > 0 Then
           '    If tmAgf.iCode <> tmChf.iAgfCode Then
           '        tmAgfSrchKey.iCode = tmChf.iAgfCode
           '        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
           '    End If
           '    slStr = Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID)
           ' Else
           '    If tmAdf.iCode <> tmChf.iAdfCode Then
           '        tmAdfSrchKey.iCode = tmChf.iAdfCode
           '        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
           '    End If
           '    slStr = Trim$(tmAdf.sName)
           'End If
           'If tmAdf.iCode <> tmChf.iAdfCode Then
           '    tmAdfSrchKey.iCode = tmChf.iAdfCode
           '    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
           'End If
            tmCpr.iAdfCode = tmChf.iAdfCode
            tmCpr.iVefCode = tmChf.iSlfCode(0)
            tmCpr.iLineNo = tmChf.iAgfCode
            tmCpr.lCntrNo = tmChf.lCntrNo
            tmCpr.lFt1CefCode = 0
           'If tmSlf.iCode <> tmChf.iSlfCode(0) Then
           '    tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
           '    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
           'End If
           'slStr = Left$(Trim$(tmSlf.sFirstName), 1) & " " & Trim(tmSlf.sLastName)

            If tgSpf.sUseProdSptScr = "P" Then    'use short title vs contr hdr product
                tmCpr.sProduct = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
            Else
                tmCpr.sProduct = Trim$(tmChf.sProduct)
            End If
        End If
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmCpr.lGenTime = lgNowTime
        tmCpr.iGenDate(0) = igNowDate(0)
        tmCpr.iGenDate(1) = igNowDate(1)

        tmCpr.sISCI = Trim$(tmCopyCntr(ilIndex).sVehName)   'rotation vehicle
        tmCpr.sCreative = Trim$(tmCopyCntr(ilIndex).sAirVehName)    'airing vehicle
        tmCpr.iLen = (tmCopyCntr(ilIndex).iLen)
        'If tmChf.iAgfCode > 0 Then
        '    tmCpr.iLineNo = ilAgfCode
        'Else
        '    tmCpr.iLineNo = 0           'agency doesnt exist
        'End If

        If RptSel!rbcSelCInclude(0).Value Then              'show cnt start date vs line start date.
            tmCpr.iSpotDate(0) = tmChf.iStartDate(0)                 'contr start date
            tmCpr.iSpotDate(1) = tmChf.iStartDate(1)
            gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slStr
        Else
            tmCpr.iSpotDate(0) = tmCopyCntr(ilIndex).iStartDate(0)                 'contr start date
            tmCpr.iSpotDate(1) = tmCopyCntr(ilIndex).iStartDate(1)
            gUnpackDate tmCopyCntr(ilIndex).iEndDate(0), tmCopyCntr(ilIndex).iEndDate(1), slStr
            tmCpr.lFt2CefCode = tmCopyCntr(ilIndex).iLineNo         '2-28-07
        End If

        tmCpr.sCartNo = Trim$(slStr)                'contr end date

        slMissingFlag = " "
        slUnassignFlag = " "
        slReadyFlag = " "
        tmCpr.sZone = ""        'zone is 3 bytes, each byte represents at least 1 spot missed
                                'for copy missing, copy unassigned, and copy ready to assign
        tmCpr.iMissing = tmCopyCntr(ilIndex).iNoSpots '# spots missing copy
        If RptSel!ckcSelC3(2).Value = vbChecked Then                    '12-12-14
            If tmCopyCntr(ilIndex).iNoSpotsMiss = 1 Then
                slMissingFlag = "*"
            End If
        Else
            slMissingFlag = "-"
            tmCpr.iMissing = 0
        End If
        
        If ilIncludeUnAssg Then
            tmCpr.iUnassign = tmCopyCntr(ilIndex).iNoUnAssg
            If tmCopyCntr(ilIndex).iNoUnAssgMiss = 1 Then
                slUnassignFlag = "*"
            End If
        Else
            slUnassignFlag = "-"
            tmCpr.iUnassign = 0     '5-20-05
        End If
        If ilIncludeReassg Then
            tmCpr.iReady = tmCopyCntr(ilIndex).iNoToReassg
            If tmCopyCntr(ilIndex).iNoToReassgMiss = 1 Then
                slReadyFlag = "*"
            End If
        Else
            slReadyFlag = "-"
            tmCpr.iReady = 0            '5-20-05
        End If
        tmCpr.sZone = slMissingFlag & slUnassignFlag & slReadyFlag
        '2-19-15 when requesting to show Daypart name under the spot, the sort isnt correct due to spots with same copy.  Use a running seq # in tmCpr.lHd1CefCode
        'use tmCpr.sStatus to test in Crystal whether to show record or not
        tmCpr.sStatus = "1"             'sort  to place non regional information before regional info
        tmCpr.lHd1CefCode = tmCpr.lHd1CefCode + 1       'keep running number for sorting
        tmCpr.sLive = tmCopyCntr(ilIndex).sLiveFlag     '11-16-05 L = Live, M = both live/recorded, R = recorded
        'ilDummy = LLDefineFieldExt(hdJob, "NoToReassg", slStr, LL_TEXT, "")
        tmCpr.iRemoteID = tmCopyCntr(ilIndex).iRdfCode                     '1-05-06
        ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)

        '8-18-00
        tmCpr.iUnassign = 0
        tmCpr.iReady = 0
        'no regional copy can exist for feed spots
        '5-23-05  regional copy unassigned is possibility a region missing-always show it
        'If ((ilIncludeUnAssg And tmCopyCntr(ilIndex).iRegionNoUnAssg > 0) Or (ilIncludeReassg And tmCopyCntr(ilIndex).iRegionNoToReassg > 0)) And (tmCopyCntr(ilIndex).lChfCode > 0) Then
        If (((tmCopyCntr(ilIndex).iRegionNoUnAssg > 0) Or (ilIncludeReassg And tmCopyCntr(ilIndex).iRegionNoToReassg > 0)) And (tmCopyCntr(ilIndex).lChfCode > 0)) Or tmCopyCntr(ilIndex).iRdfCode > 0 Then
            tmCpr.iRemoteID = tmCopyCntr(ilIndex).iRdfCode      '1-5-06
            tmCpr.sZone = ""
            tmCpr.iMissing = 0 '# spots missing
            'If ilIncludeUnAssg Then        'always show the regional unassigned, its considered missing
                tmCpr.iUnassign = tmCopyCntr(ilIndex).iRegionNoUnAssg
            'End If
            If ilIncludeReassg Then
                tmCpr.iReady = tmCopyCntr(ilIndex).iRegionNoToReassg
            End If
            '2-19-15 when requesting to show Daypart name under the spot, the sort isnt correct due to spots with same copy.  Use a running seq # in tmCpr.lHd1CefCode
            'use tmCpr.sStatus to test in Crystal whether to show record or not
            'the Status =2 records get sorted together within the same cnt & copy
            tmCpr.sStatus = "2"             'sort  to place  regional information after non regional info
            tmCpr.lHd1CefCode = tmCpr.lHd1CefCode + 1       'keep running number for sorting
            ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)     'write out regional record
        End If
        '**********


    Next ilIndex

    Erase tmCopyCntr, tmVlf
    Erase tmCRFByCntr
    Screen.MousePointer = vbDefault
    mCloseCopyFiles
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmVLF)
    btrDestroy hmVLF
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff

    Exit Sub
End Sub
'**************************************************************************
'*                                                                        *
'*      Procedure Name:gCopyDateRpt                                       *
'*                                                                        *
'*            Created:4/21/94        By:D. LeVine                         *
'*            Modified:              By:D. Smith                          *
'*                                                                        *
'*            DS 9/25/00 Converted to Crystal from Bridge                 *
'*                                                                        *
'*            Comments: Generate Copy by Date report                      *
'*           7-27-04 Add option to include/exclude contract/feed spots
'*           11-30-04 Change access of smf from key0 to key2 for speed                                                                       *
'*      dh 3-29-05 create BB spots if not yet created to test if copy exists
'**************************************************************************
Sub gCopyDateRpt()
    Dim slFileInError As String
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llIndex As Long         '7-17-09 chged from integer to long
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim ilSpotType As Integer
    Dim slCopyProduct As String
    Dim slCopyZone As String
    Dim slCopyISCI As String
    Dim slCopyCart As String
    Dim ilIncludeUnassigned As Integer
    Dim tlVef As VEF
    Dim slProdOrShortT As String
    Dim slName As String
    Dim slShowOnInv As String * 1
    Dim ilContrSpots As Integer
    Dim slChfFsfProduct As String
    Dim ilFeedSpots As Integer
    Dim slSaveAdvName As String
    Dim llStartDate As Long         '3-29-05 loop on dates requested by vehicle to create bb spots
    Dim llEndDate As Long
    Dim llDate As Long
    Dim slDate As String

    Screen.MousePointer = vbHourglass
'    slStartDate = RptSel!edcSelCFrom.Text   'Start date
'    slEndDate = RptSel!edcSelCTo.Text       'End date
'   8-22-19 use csi calendar control vs edit box
    slStartDate = RptSel!CSI_CalFrom.Text   'Start date
    slEndDate = RptSel!CSI_CalTo.Text       'End date
    slDateRange = "From " & slStartDate & " To " & slEndDate
    If RptSel!rbcSelCSelect(0).Value Then  'All spots
        ilSpotType = 0
        slDateRange = slDateRange & " for All Spots"
    ElseIf RptSel!rbcSelCSelect(1).Value Then  'Only spots with copy
        ilSpotType = 1
        slDateRange = slDateRange & " for Spots with Copy"
    Else    'Only spots without copy
        ilSpotType = 2
        slDateRange = slDateRange & " for Spots without Copy"
    End If

    If Not gSetFormula("ShowDateBanner", "'" & slDateRange & "'") Then
        Exit Sub
    End If
    If ilSpotType = 0 Then
        ilIncludeUnassigned = True
    Else
        If RptSel!rbcSelCInclude(0).Value Then
            ilIncludeUnassigned = True
        Else
            ilIncludeUnassigned = False
        End If
    End If


    ilContrSpots = True                                     'assume to include both contract & feed spots
    ilFeedSpots = True
    If Not RptSel!ckcSelC10(0).Value = vbChecked Then       'include Contracts spots?
        ilContrSpots = False
    End If
    If Not RptSel!ckcSelC10(1).Value = vbChecked Then       'include feed spots?
        ilFeedSpots = False
    End If

    slFileInError = mOpenCopyStatusFiles()          'open all common copy files
    If slFileInError <> "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Error opening " & Trim$(slFileInError) & " - Rptgen: gCopyDateRpt"
        Exit Sub
    End If

    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)

    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCpfRecLen = Len(tmCpf)

    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMcfRecLen = Len(tmMcf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseCopyFiles
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmSmf
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)

'    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSdf)
'        ilRet = btrClose(hmSmf)
'        ilRet = btrClose(hmMcf)
'        ilRet = btrClose(hmCpf)
'        ilRet = btrClose(hmCif)
'        btrDestroy hmSmf
'        btrDestroy hmMcf
'        btrDestroy hmCpf
'        btrDestroy hmCif
'        btrDestroy hmSdf
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If

    llStartDate = gDateValue(slStartDate)       '3-29-05 convert string date to long for looping
    llEndDate = gDateValue(slEndDate)

    tmAgf.iCode = 0
    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0
    If tgSpf.sUseProdSptScr = "P" Then    'use short title vs contr hdr product
        slProdOrShortT = "Advertiser, Short Title"
    Else
        slProdOrShortT = "Advertiser, Product"
    End If
    If Not gSetFormula("ShowAdvProdOrShort", "'" & slProdOrShortT & "'") Then
        mCloseCopyFiles
        'ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        btrDestroy hmSmf
        btrDestroy hmMcf
        btrDestroy hmCpf
        btrDestroy hmCif
        'btrDestroy hmSdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    DoEvents
    ReDim tmCopy(0 To 0) As COPYSORT
    For ilVehicle = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
        If RptSel!lbcSelection(1).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                illoop = gBinarySearchVpf(ilVefCode)
                If illoop <> -1 Then
                    ilVpfIndex = illoop
            '        Exit For
                End If
            'Next ilLoop

            '3-29-05 go out and create all BB spots if necessary for dates requested
            For llDate = llStartDate To llEndDate
                slDate = Format(llDate, "m/d/yy")
                ilRet = gCreateBBSpots(hmSdf, ilVefCode, slDate)
            Next llDate
            mObtainCopyDate 0, ilVefCode, ilVpfIndex, slStartDate, slEndDate, ilSpotType, ilIncludeUnassigned, ilContrSpots, ilFeedSpots
        End If
    Next ilVehicle

    'outer loop - one loop per page
    llIndex = LBound(tmCopy)
    If llIndex >= UBound(tmCopy) Then
        ilDBRet = 1
        If ilSpotType = 2 Then  'Nothing to display. We can send one record with 32000 in
            'the veh code to cause Crystal to print the **** NONE **** ,but it causes bad
            'time, length etc values to display
            'tmCpr.iGenTime(0) = igNowTime(0)
            'tmCpr.iGenTime(1) = igNowTime(1)
            'tmCpr.iGenDate(0) = igNowDate(0)
            'tmCpr.iGenDate(1) = igNowDate(1)
            'tmCpr.iVefCode = 32000  '
            'ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
        End If
    Else
        'Sort key
        ArraySortTyp fnAV(tmCopy(), 0), UBound(tmCopy), 0, LenB(tmCopy(0)), 0, LenB(tmCopy(0).sKey), 0
        ilDBRet = BTRV_ERR_NONE
    End If
    While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0
        'While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0 And ilRet = 0      5-12-05 remove
            tmSdf = tmCopy(llIndex).tSdf
            'tmCpr.iGenTime(0) = igNowTime(0)
            'tmCpr.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCpr.lGenTime = lgNowTime
            tmCpr.iGenDate(0) = igNowDate(0)
            tmCpr.iGenDate(1) = igNowDate(1)
            If tmCopy(llIndex).iRegionalSort = 0 Then           '8-16-00 spot line (vs regional copy line_
                tmCpr.iSpotDate(0) = tmSdf.iDate(0)
                tmCpr.iSpotDate(1) = tmSdf.iDate(1)
                If tmSdf.iVefCode <> tmVef.iCode Then
                    tmVefSrchKey.iCode = tmSdf.iVefCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        tmVef.sName = "Missing"
                    End If
                End If
                tmCpr.iVefCode = tmSdf.iVefCode
                tmCpr.iSpotTime(0) = tmSdf.iTime(0)
                tmCpr.iSpotTime(1) = tmSdf.iTime(1)
                tmCpr.iLen = tmSdf.iLen
                If tmSdf.sSpotType = "O" Then           '3-29-05 set flag to show bb on report
                    tmCpr.lFt2CefCode = 1
                ElseIf tmSdf.sSpotType = "C" Then       'closed bb
                    tmCpr.lFt2CefCode = 2
                Else
                    tmCpr.lFt2CefCode = 0
                End If

                slChfFsfProduct = ""
                illoop = gBinarySearchAdf(tmSdf.iAdfCode)       'find the advertiser record
                slSaveAdvName = ""
                If illoop <> -1 Then
                    'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                    '    slSaveAdvName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
                    'Else
                        slSaveAdvName = Trim$(tgCommAdf(illoop).sName)
                    'End If
                End If

                'obtain feed spot or contract header for product
                If tmSdf.lChfCode = 0 Then              'feed spot
                    tmFSFSrchKey.lCode = tmSdf.lFsfCode
                    ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        tmPrfSrchKey.lCode = tmFsf.lPrfCode
                        ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            slChfFsfProduct = Trim$(tmPrf.sName)
                        End If
                        mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct
                        If Trim$(slCopyProduct) = "" Then
                            slStr = Trim$(slSaveAdvName) & "," & Trim$(slChfFsfProduct)
                        Else
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slCopyProduct)
                        End If
                        tmCpr.lCntrNo = 0
                        tmCpr.iLineNo = 0
                        tmCpr.lFt1CefCode = tmFsf.lCode
                    End If
                Else
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slChfFsfProduct = Trim$(tmChf.sProduct)
                    mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct

                    If Trim$(slCopyProduct) = "" Then         'no copy found, get short title (if applicable from chf)
                        If tgSpf.sUseProdSptScr = "P" Then    'use short title vs contr hdr product
                            slStr = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slStr)
                        Else
                            slStr = Trim$(slSaveAdvName) & ", " & Trim$(slChfFsfProduct)
                        End If
                    Else                      'copy found, proper short title or product retrieved
                        slStr = Trim$(slSaveAdvName) & ", " & Trim$(slCopyProduct)
                    End If
                    tmCpr.lCntrNo = tmChf.lCntrNo
                    tmCpr.iLineNo = tmSdf.iLineNo

                End If

                tmCpr.sProduct = slStr
                If tmCopy(llIndex).iCopyStatus = 0 Then 'Not defined
                    tmCpr.sZone = " "
                    tmCpr.sCartNo = " "
                    tmCpr.sISCI = " "
                ElseIf tmCopy(llIndex).iCopyStatus = 2 Then 'defined, not assigned
                    tmCpr.sZone = " "
                    tmCpr.sCartNo = "*"
                    tmCpr.sISCI = " "
                Else
                    tmCpr.sZone = slCopyZone
                    If tmCopy(llIndex).iCopyStatus = 1 Then 'Ok
                        tmCpr.sCartNo = slCopyCart
                    ElseIf tmCopy(llIndex).iCopyStatus = 3 Then 'Superseded
                        tmCpr.sCartNo = "^" & slCopyCart
                    ElseIf tmCopy(llIndex).iCopyStatus = 4 Then 'Zone missing
                        tmCpr.sCartNo = "~" & slCopyCart
                    End If
                    tmCpr.sISCI = slCopyISCI
                End If
                If tmCopy(llIndex).iRegionalStatus = 1 Then     '8-16-00
                    tmCpr.sStatus = "W"         'other cnts for same advt have regional copy defined
                ElseIf tmCopy(llIndex).iRegionalStatus = 2 Then
                    tmCpr.sStatus = "*"
                ElseIf tmCopy(llIndex).iRegionalStatus = 3 Then
                    tmCpr.sStatus = "^"
                ElseIf tmCopy(llIndex).iRegionalStatus = 4 Then     'ok
                    tmCpr.sStatus = ""
                Else                'none defined
                    tmCpr.sStatus = ""
                End If

                tmSmf.iOrigSchVef = tmSdf.iVefCode
                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    '11-30-04 change access of smf from key0 to key2 for speed
                    'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                    'tmSmfSrchKey.lFsfCode = tmSdf.lFsfCode
                    'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                    'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                    'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo) And (tmSmf.lFsfCode = tmSdf.lFsfCode)
                        If (tmSmf.lSdfCode = tmSdf.lCode) Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    Loop
                End If
                If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                    tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                If tmSdf.sSchStatus = "S" Then
                    If tmSdf.sSpotType <> "X" Then
                        slStr = "Scheduled"
                    Else
                        'If tmSdf.sPriceType = "N" Then  'fill?
                        '3-23-03 Test the advt instead of spot to determine a fill or extra


                        '1-19-04 change way in which fill/extra are shown
                        slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                        'If tmAdf.sBonusOnInv = "N" Then
                        If slShowOnInv = "N" Then
                            slStr = "-Schd Fill"
                        Else
                            'slStr = "Schd Extra"
                            slStr = "+Schd Fill"
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "M" Then
                    slStr = "Missed"
                ElseIf tmSdf.sSchStatus = "R" Then
                    slStr = "Ready"
                ElseIf tmSdf.sSchStatus = "U" Then
                    slStr = "UnSched"
                ElseIf tmSdf.sSchStatus = "G" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                        Else
                            'If tmSdf.sPriceType = "N" Then 'fill?
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                            Else
                                'slStr = "Extra Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                                slStr = "+Fill Makegood" '& Chr$(10) & Trim$(tlVef.sName)
                            End If
                        End If
                    Else
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Makegood"
                        Else
                            'If tmSdf.sPriceType = "N" Then      'fill?
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Makegood"
                            Else
                                'slStr = "Extra Makegood"
                                slStr = "+Fill Makegood"
                            End If
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "O" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Outside" & Chr$(10) & Trim$(tlVef.sName)
                        Else
                            'If tmSdf.sPriceType = "N" Then
                            '3-23-03 Test the advt instead of spot to determine a fill or extra


                            '1-19-04 change way in which fill/extra are shown
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Outside" & Chr$(10) & Trim$(tlVef.sName)
                            Else
                                'slStr = "Extra Outside" & Chr$(10) & Trim$(tlVef.sName)
                                slStr = "+Fill Outside" & Chr$(10) & Trim$(tlVef.sName)
                            End If
                        End If
                        'slStr = "Outside" & Chr$(10) & Trim$(tlVef.sName)
                    Else
                        If tmSdf.sSpotType <> "X" Then
                            slStr = "Outside"
                        Else
                            '6-8-09
                            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                            'If tmAdf.sBonusOnInv = "N" Then
                            If slShowOnInv = "N" Then
                                slStr = "-Fill Outside"
                            Else
                                'slStr = "+Extra Outside"
                                slStr = "+Fill Outside"
                            End If
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "C" Then
                    slStr = "Cancelled"
                ElseIf tmSdf.sSchStatus = "H" Then
                    slStr = "Hidden"
                ElseIf tmSdf.sSchStatus = "A" Then
                    slStr = "On Alt"
                ElseIf tmSdf.sSchStatus = "B" Then
                    slStr = "On Alt & MG"
                End If
                tmCpr.sCreative = slStr
                tmCpr.lHd1CefCode = 1
                ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                DoEvents
                If tmSdf.sSchStatus = "G" Then          'show where makegood came from
                    tmCpr.sCartNo = " "
                    tmCpr.sISCI = " "
                    tmCpr.sCreative = " "
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slStr
                    tmCpr.sISCI = gAddDayToDate(slStr)
                    gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slStr
                    tmCpr.sCreative = slStr

                    If tmSmf.iOrigSchVef <> tmVef.iCode Then
                        tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmVef.sName = "Missing"
                        End If
                    End If
                    tmCpr.sCartNo = tmVef.sName
                    tmCpr.lHd1CefCode = 2
                    ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                    tmCpr.sCartNo = " "
                    tmCpr.sISCI = " "
                    tmCpr.sCreative = " "
                End If

            Else                        'regional copy line
                mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slChfFsfProduct
                tmCpr.sCartNo = slCopyCart
                tmCpr.sISCI = slCopyISCI
                tmCpr.lHd1CefCode = 3
                ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                DoEvents
            End If

            'next data set if no error or warning
            If ilRet = 0 Then
                llIndex = llIndex + 1
                llRecNo = llRecNo + 1
                If llIndex >= UBound(tmCopy) Then
                    ilDBRet = 1
                End If
            End If
        'Wend  ' inner loop         5-12-05 remove
    Wend      ' while not EOF

    Erase tmCopy
    Screen.MousePointer = vbDefault
    mCloseCopyFiles
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmMcf)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmCif)
    btrDestroy hmSmf
    btrDestroy hmMcf
    btrDestroy hmCpf
    btrDestroy hmCif
    Exit Sub
End Sub
'
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gDumpFileRpt                  *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate Dallas feed report    *
'*      5-14-02 DH Convert from Bridge to Crystal
'*                                                     *
'*******************************************************
Sub gDumpFileRpt(ilReadMode As Integer, ilListIndex As Integer, slTitle As String)
'Sub gDumpFileRpt (ilReadMode As Integer, ilPreview As Integer, slName As String, ilListIndex As Integer, slTitle As String)
'
'   gDumpFileRpt ilReadMode, ilPreview, slName, ilListIndex, slTitle
'   Where:
'       ilReadMode(I)- 0-Sequential Mode; 1=Invoice Export
'       ilPreview(I)- 0=Print; <>0=Preview
'       slName(I)- File Name of lst structure (DumpFile.Lst)
'       ilListIndex(I)- lbcSelection index
'       slTitle(I)- Title for the dump
'
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBEof As Integer
    Dim ilRet As Integer
    Dim hlRead As Integer
    Dim illoop As Integer
    Dim slDumpFileName As String
    Dim slLine As String
    Dim llNoRecsToProc As Long
    Dim llFileLen As Long
    Dim slChar As String * 1
    Dim ilFirstRead As Integer
    Dim slPrevChar As String * 1

    hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTxr)
        ilRet = btrClose(hmVef)
        btrDestroy hmTxr
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imTxrRecLen = Len(tmTxr)


    llNoRecsToProc = 0
    llFileLen = 0
    For illoop = 0 To RptSel!lbcSelection(ilListIndex).ListCount - 1 Step 1
        If RptSel!lbcSelection(ilListIndex).Selected(illoop) Then
            slDumpFileName = RptSel!lbcSelection(ilListIndex).List(illoop)
            llNoRecsToProc = llNoRecsToProc + FileLen(sgExportPath & slDumpFileName)
        End If
    Next illoop
    'slFileName = sgRptPath & slName & Chr$(0)
    'gDumpFile


    'ilAnyOutput = False
    'slLangText = "Printing..."
    'If ilPreview <> 0 Then
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
    '    ilDummy = LlPreviewSetTempPath(hdJob, sgRptSavePath)
    '    ilDummy = LlPreviewSetResolution(hdJob, 200)
    'Else
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
    'End If
    'DoEvents
    'If (ilRet = 0) Then
        llRecNo = 1
        ilErrorFlag = 0
        ilFirstRead = True
        slPrevChar = " "
        'slPrinter = LlVBPrintGetPrinter(hdJob)
        'slPort = LlVBPrintGetPort(hdJob)
        For illoop = 0 To RptSel!lbcSelection(ilListIndex).ListCount - 1 Step 1
            If RptSel!lbcSelection(ilListIndex).Selected(illoop) Then
                slDumpFileName = RptSel!lbcSelection(ilListIndex).List(illoop)
                Do While Len(slDumpFileName) < 12
                    slDumpFileName = slDumpFileName & " "  'fill with blanks for 12 char:
                Loop
                tmTxr.lSeqNo = 0

                ilDBEof = False
                'On Error GoTo mDumpFileErr:
                'hlRead = FreeFile
                If ilReadMode = 0 Then
                    'Open sgExportPath & slDumpFileName For Input Access Read As hlRead
                    ilRet = gFileOpen(sgExportPath & slDumpFileName, "Input Access Read", hlRead)
                Else
                    'Open sgExportPath & slDumpFileName For Binary Access Read As hlRead
                    ilRet = gFileOpen(sgExportPath & slDumpFileName, "Binary Access Read", hlRead)
                End If
                If ilRet <> 0 Then
                    ilDBEof = True
                End If
                If Not ilDBEof Then
                    err.Clear
                    ilDBEof = False
                    'ilDummy = LlDefineVariableExt(hdJob, "Logo", sgLogoPath & "RptLogo.Bmp", LL_DRAWING, "")
                    'ilDummy = LLDefineVariableExtHandle(hdJob, "CSILogo", Traffic!imcCSILogo, LL_DRAWING_HBITMAP, "")
                    'ilDummy = LlDefineVariableExt(hdJob, "Title", slTitle, LL_TEXT, "")
                    'ilDummy = LlDefineVariableExt(hdJob, "FileName", slDumpFileName, LL_TEXT, "")
                    If ilReadMode = 0 Then
                        If EOF(hlRead) Then
                            ilDBEof = True
                        Else
                            Line Input #hlRead, slLine
                        End If
                    ElseIf ilReadMode = 1 Then
                        slLine = ""
                        Do
                            If ilFirstRead Then
                                Get #hlRead, 1, slChar
                                ilFirstRead = False
                            Else
                                Get #hlRead, , slChar
                            End If
                            If err.Number <> 0 Then
                                ilDBEof = True
                            End If
                            If (ilDBEof) Or ((Asc(slPrevChar) = 21) And (Asc(slChar) = 0)) Then
                                ilDBEof = True
                                Exit Do
                            End If
                            slPrevChar = slChar
                            If Asc(slChar) = 21 Then
                                slLine = slLine & slChar
                                Exit Do
                            Else
                                slLine = slLine & slChar
                            End If
                        Loop
                    End If
                    If err.Number <> 0 Then
                        ilDBEof = True
                    End If
                    If Len(slLine) > 0 Then
                        If (Asc(slLine) = 26) Then    'Ctrl Z
                            ilDBEof = True
                        End If
                    End If
                    llRecNo = llRecNo + 1
                End If
                'outer loop - one loop per page
                While (Not ilDBEof) 'And ilErrorFlag = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
                    'ilAnyOutput = True
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
                    'ilRet = LLPrint(hdJob)
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)

                    While (Not ilDBEof) 'And ilErrorFlag = 0 And ilRet = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)

                        'Call DefineFields
                        'LlDefineFieldStart hdJob
                        'ilDummy = LLDefineFieldExt(hdJob, "Dump", slLine, LL_TEXT, "")
                        tmTxr.sText = slDumpFileName & Mid$(slLine, 1, Len(slLine))

                        tmTxr.iGenDate(0) = igNowDate(0)
                        tmTxr.iGenDate(1) = igNowDate(1)
                        tmTxr.lGenTime = lgNowTime
                        '6/30/06: Added Comment field for Commercial Log generation in Station Feed to handle long comments
                        tmTxr.lCsfCode = 0
                        '6/30/06: End of Change
                        ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                        tmTxr.lSeqNo = tmTxr.lSeqNo + 1     'keep in order

                        'notify the user (how far have we come?)
                        If ilReadMode = 0 Then
                            llFileLen = llFileLen + Len(slLine) + 2
                        ElseIf ilReadMode = 1 Then
                            llFileLen = llFileLen + Len(slLine)
                        End If
                       ' ilRet = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * llFileLen / llNoRecsToProc))
                       ' DoEvents
                        'tell L&L to print the table line
                        'If ilRet = 0 Then
                        '    ilRet = LlPrintFields(hdJob)
                        'End If


                        'next data set if no error or warning
                        If ilRet = 0 Then

                            'Call DBNext
                            If ilReadMode = 0 Then
                                If EOF(hlRead) Then
                                    ilDBEof = True
                                Else
                                    Line Input #hlRead, slLine
                                End If
                            ElseIf ilReadMode = 1 Then
                                slLine = ""
                                Do
                                    Get #hlRead, , slChar
                                    If err.Number <> 0 Then
                                        ilDBEof = True
                                    End If
                                    If (ilDBEof) Or ((Asc(slPrevChar) = 21) And (Asc(slChar) = 0)) Then
                                        ilDBEof = True
                                        Exit Do
                                    End If
                                    slPrevChar = slChar
                                    If Asc(slChar) = 21 Then
                                        slLine = slLine & slChar
                                        Exit Do
                                    Else
                                        slLine = slLine & slChar
                                    End If
                                Loop
                            End If
                            If err.Number <> 0 Then
                                ilDBEof = True
                            End If
                            If Len(slLine) > 0 Then
                                If (Asc(slLine) = 26) Then    'Ctrl Z
                                    ilDBEof = True
                                End If
                            End If
                            llRecNo = llRecNo + 1
                        Else
                            If ilReadMode = 0 Then
                                llFileLen = llFileLen - Len(slLine) - 2
                            ElseIf ilReadMode = 1 Then
                                llFileLen = llFileLen - Len(slLine)
                            End If
                        End If
                    Wend  ' inner loop

                    'if error or warning: different reactions:
                    'If ilRet < 0 Then
                    '    If ilRet <> LL_WRN_REPEAT_DATA Then
                    '        ilErrorFlag = ilRet
                    '    End If
                    'End If

                Wend    ' while not EOF
                Close hlRead
            End If
        Next illoop
        'end print
        'If Not ilAnyOutput Then
        '    ilRet = LLPrintEnableObject(hdJob, ":Spots", False)
        '    ilRet = LLPrint(hdJob)
        'Else
        '    ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
        'End If
        'ilRet = LlPrintEnd(hdJob, 0)

        'in case of preview: show the preview
        'If ilPreview <> 0 Then
        '    If ilErrorFlag = 0 Then
        '        ilDummy = LlPreviewDisplay(hdJob, slFileName, sgRptSavePath, RptSel.hWnd)
        '    Else
        '        mErrMsg ilErrorFlag
        '    End If
        '    ilDummy = LlPreviewDeleteFiles(hdJob, slFileName, sgRptSavePath)
        'Else
        '    If ilErrorFlag <> 0 Then
        '        mErrMsg ilErrorFlag
        '    End If
        'End If

    'Else  ' LlPrintWithBoxStart
    '    ilErrorFlag = ilRet
    '    ilRet = LlPrintEnd(hdJob, 0)
    '    mErrMsg ilErrorFlag
    'End If  ' LlPrintWithBoxStart

    'End If  ' LlSelecTLileDlgTitle
    Exit Sub
'mDumpFileErr:
'    ilDBEof = True
'    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gListFileRpt                    *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate List report (dump file*
'*                      include form feed)             *
'*                                                     *
'*       5-14-02 DH Convert from Bridge to Crystal
'*******************************************************
Sub gListFileRpt(ilListIndex As Integer, slTitle As String)
'Sub gListFileRpt(ilReadMode As Integer, ilPreview As Integer, slName As String, ilListIndex As Integer, slTitle As String)
'
'   gDumpFileRpt ilReadMode, ilPreview, slName, ilListIndex, slTitle
'   Where:
'       ilReadMode(I)- 0-Sequential Mode
'       ilPreview(I)- 0=Print; <>0=Preview
'       slName(I)- File Name of lst structure (ListFile.Lst)
'       ilListIndex(I)- lbcSelection index
'       slTitle(I)- Title for the dump
'
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBEof As Integer
    Dim ilRet As Integer
    Dim hlRead As Integer
    Dim illoop As Integer
    Dim slListFileName As String
    Dim slLine As String
    Dim llNoRecsToProc As Long
    Dim llFileLen As Long
    Dim ilFirstRead As Integer
    Dim slPrevChar As String * 1
    hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTxr)
        ilRet = btrClose(hmVef)
        btrDestroy hmTxr
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imTxrRecLen = Len(tmTxr)

    llNoRecsToProc = 0
    llFileLen = 0
    For illoop = 0 To RptSel!lbcSelection(ilListIndex).ListCount - 1 Step 1
        If RptSel!lbcSelection(ilListIndex).Selected(illoop) Then
            slListFileName = RptSel!lbcSelection(ilListIndex).List(illoop)
            llNoRecsToProc = llNoRecsToProc + FileLen(sgExportPath & slListFileName)
        End If
    Next illoop


    'slFileName = sgRptPath & slName & Chr$(0)
    'gListFile

    'ilAnyOutput = False
    'slLangText = "Printing..."
    'If ilPreview <> 0 Then
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
    '    ilDummy = LlPreviewSetTempPath(hdJob, sgRptSavePath)
    '    ilDummy = LlPreviewSetResolution(hdJob, 200)
    'Else
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
    'End If
    'DoEvents
    'If (ilRet = 0) Then
        llRecNo = 1
        ilErrorFlag = 0
        ilFirstRead = True
        slPrevChar = " "
        'slPrinter = LlVBPrintGetPrinter(hdJob)
        'slPort = LlVBPrintGetPort(hdJob)
        tmTxr.iType = 0         'no form feed 1st time thru
        For illoop = 0 To RptSel!lbcSelection(ilListIndex).ListCount - 1 Step 1
            If RptSel!lbcSelection(ilListIndex).Selected(illoop) Then
                slListFileName = RptSel!lbcSelection(ilListIndex).List(illoop)
                ilDBEof = False
                'On Error GoTo mListFileErr:
                'hlRead = FreeFile
                'Open sgExportPath & slListFileName For Input Access Read As hlRead
                ilRet = gFileOpen(sgExportPath & slListFileName, "Input Access Read", hlRead)
                If ilRet <> 0 Then
                    ilDBEof = True
                End If
                If Not ilDBEof Then
                    ilDBEof = False
                    err.Clear
                    If EOF(hlRead) Then
                        ilDBEof = True
                    Else
                        Line Input #hlRead, slLine
                    End If
                    If err.Number <> 0 Then
                        ilDBEof = True
                    End If
                    If Len(slLine) > 0 Then
                        If (Asc(slLine) = 26) Then    'Ctrl Z
                            ilDBEof = True
                        End If
                    End If
                    llRecNo = llRecNo + 1
                End If
                'outer loop - one loop per page
                err.Clear
                While (Not ilDBEof) 'And ilErrorFlag = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
                    'ilAnyOutput = True
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
                    'ilRet = LLPrint(hdJob)
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)

                    While (Not ilDBEof) 'And ilErrorFlag = 0 And ilRet = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)

                        'Call DefineFields
                        'LlDefineFieldStart hdJob
                        'ilDummy = LLDefineFieldExt(hdJob, "Dump", slLine, LL_TEXT, "")
                        tmTxr.sText = Mid$(slLine, 1, Len(slLine))
                        tmTxr.iGenDate(0) = igNowDate(0)
                        tmTxr.iGenDate(1) = igNowDate(1)
                        tmTxr.lGenTime = lgNowTime

                        ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                        tmTxr.iType = 0     'no form feed flag
                        'notify the user (how far have we come?)
                        'llFileLen = llFileLen + Len(slLine) + 2
                        'ilRet = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * llFileLen / llNoRecsToProc))
                        'DoEvents
                        'tell L&L to print the table line
                        'If ilRet = 0 Then
                        '    ilRet = LlPrintFields(hdJob)
                        'End If

                        'next data set if no error or warning
                        'If ilRet = 0 Then

                            'Call DBNext
                            If EOF(hlRead) Then
                                ilDBEof = True
                            Else
                                Line Input #hlRead, slLine
                            End If
                            If err.Number <> 0 Then
                                ilDBEof = True
                            End If
                            If Len(slLine) > 0 Then
                                If (Asc(slLine) = 26) Then    'Ctrl Z
                                    ilDBEof = True
                                ElseIf (Asc(slLine) = 12) Then  'Form Feed
                                    'ilRet = LL_WRN_REPEAT_DATA
                                    tmTxr.iType = 1     'force form feed in Crystal
                                    If EOF(hlRead) Then
                                        ilDBEof = True
                                    Else
                                        Line Input #hlRead, slLine
                                    End If
                                    If err.Number <> 0 Then
                                        ilDBEof = True
                                    End If
                                    If Len(slLine) > 0 Then
                                        If (Asc(slLine) = 26) Then    'Ctrl Z
                                            ilDBEof = True
                                        End If
                                    End If
                                End If
                            End If
                            llRecNo = llRecNo + 1
                        'Else
                            llFileLen = llFileLen - Len(slLine) - 2
                        'End If
                    Wend  ' inner loop

                    'if error or warning: different reactions:
                    'If ilRet < 0 Then
                    '    If ilRet <> LL_WRN_REPEAT_DATA Then
                    '        ilErrorFlag = ilRet
                    '    End If
                    'End If

                Wend    ' while not EOF
                Close hlRead
            End If
            tmTxr.iType = 1    'force form feed for next report
        Next illoop
        'end print
        'If Not ilAnyOutput Then
        '    ilRet = LLPrintEnableObject(hdJob, ":Spots", False)
        '    ilRet = LLPrint(hdJob)
        'Else
        '    ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
        'End If
        'ilRet = LlPrintEnd(hdJob, 0)

        'in case of preview: show the preview
        'If ilPreview <> 0 Then
        '    If ilErrorFlag = 0 Then
        '        ilDummy = LlPreviewDisplay(hdJob, slFileName, sgRptSavePath, RptSel.hWnd)
        '    Else
        '        mErrMsg ilErrorFlag
        '    End If
        '    ilDummy = LlPreviewDeleteFiles(hdJob, slFileName, sgRptSavePath)
        'Else
        '    If ilErrorFlag <> 0 Then
        '        mErrMsg ilErrorFlag
        '    End If
        'End If

    'Else  ' LlPrintWithBoxStart
    '    ilErrorFlag = ilRet
    '    ilRet = LlPrintEnd(hdJob, 0)
    '    mErrMsg ilErrorFlag
    'End If  ' LlPrintWithBoxStart

    'End If  ' LlSelecTLileDlgTitle
    Exit Sub
'mListFileErr:
'    ilDBEof = True
'    Resume Next
End Sub
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gPLogRpt                                    *
'*                                                                 *
'*            Created:4/21/94       By:D. LeVine                   *
'*            Modified:             By:D. Smith                    *
'*            Converted to Crystal  7/11/00                        *
'*                                                                 *
'*            Comments: Generate Post Log report                   *
'
'           9-1-00 Show orig date missed for Outside &             *
'               MG spots.  Show flag if Hidden in Status           *
'               field
'*
'       1-24-03 tmSdf.sAffChg was using cpr.sunused field,         *
'               which has been used for another field.             *
'               Store .sAffChg in cpr.iUnassign                    *
'       7-21-04 option to include/exclude contract/feed spots      *
'       11-30-04 change access of smf from key0 to key2 for speed
'       3-28-05 show open/close bb notation on spot
'       5-26-06 If spot is crossing midnight, adjust the date to the
'               true date of airing (next day)
'*******************************************************************
Sub gPlogRpt(ilISCIOnly As Integer)
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilIndex As Integer
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim slOrdered As String
    Dim slDateRange As String
    Dim slSpotAmount As String
    Dim ilNoDays As Integer
    Dim slCopyProduct As String
    Dim slProduct As String         'common field for product name for: contract spot or feed spot
    Dim slCopyZone As String
    Dim slCopyISCI As String
    Dim slCopyCart As String
    Dim slContrNo As String
    Dim ilUpper As Integer
    Dim ilIncludePSA As Integer
    Dim ilSpotType As Integer
    Dim ilBillType As Integer
    Dim ilMissedType As Integer
    Dim ilCostType As Integer           '-1 (ignore testing of spot type in mobtainsdf routine)
    Dim ilByOrderOrAir As Integer   '0=Order; 1=Aired
    Dim tlVef As VEF
    'ReDim ilDayNoSpots(1 To 4) As Integer    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    'ReDim slDayDollars(1 To 4) As String    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    'ReDim ilVehNoSpots(1 To 4) As Integer    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    'ReDim slVehDollars(1 To 4) As String    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    'ReDim ilAllNoSpots(1 To 4) As Integer    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    'ReDim slAllDollars(1 To 4) As String    '1=Unbilled; 2=Billed; 3=Missed; 4=Cancelled
    Dim slStatus As String * 3                '9-1-00 (H) for hidden
    Dim slShowOnInv As String * 1
    Dim ilCntrSpots As Integer
    Dim ilFeedSpots As Integer
    Dim slContrNumber As String
    Dim llDate As Long
    Dim ilUpperDay As Integer
    Dim ilfoundDay As Integer
    Dim ilVefIndex As Integer
    Dim ilGameSelect As Integer
    Dim ilLineSelect As Integer
    Dim llChfSelect As Long
    
    'ilCostType = -1                 'ignore testing of spot types in mObtainSdf
    ilCostType = 0
    'set all spot types included
    ilCostType = ilCostType Or SPOT_CHARGE          'bit 0
    ilCostType = ilCostType Or SPOT_00
    ilCostType = ilCostType Or SPOT_ADU
    ilCostType = ilCostType Or SPOT_BONUS

    If RptSel!ckcSelC3(7).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_EXTRA
    End If
    If RptSel!ckcSelC3(8).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_FILL
    End If

    ilCostType = ilCostType Or SPOT_NC
    ilCostType = ilCostType Or SPOT_MG
    ilCostType = ilCostType Or SPOT_RECAP
    ilCostType = ilCostType Or SPOT_SPINOFF
    ilCostType = ilCostType Or SPOT_BB

    Screen.MousePointer = vbHourglass
'    slStartDate = RptSel!edcSelCFrom.Text   'Start date
'    slEndDate = RptSel!edcSelCTo.Text   'End date
    slStartDate = RptSel!CSI_CalFrom.Text   'Start date
    slEndDate = RptSel!CSI_CalTo.Text   'End date

    slContrNo = ""
    slContrNo = RptSel!edcSelCTo1.Text   'Contr #
    slDateRange = "From " & slStartDate & " To " & slEndDate
    If (slStartDate = "") And (slEndDate = "") Then
        slDateRange = "All Dates"
    End If
    ilBillType = 0
    If Not ilISCIOnly Then
        If RptSel!ckcSelC3(0).Value = vbChecked Then    'Billed
            ilBillType = ilBillType Or 1
        End If
        If RptSel!ckcSelC3(1).Value = vbChecked Then    'Unbilled
            ilBillType = ilBillType Or 2
        End If
        If RptSel!ckcSelC3(3).Value = vbChecked Then    'Include PSA/Promo
            ilIncludePSA = True
        Else
            ilIncludePSA = False
        End If
        ilMissedType = 0
        If RptSel!ckcSelC3(4).Value = vbChecked Then
            ilMissedType = 1
        End If
        If RptSel!ckcSelC3(5).Value = vbChecked Then
            ilMissedType = ilMissedType + 2
        End If
        If RptSel!ckcSelC3(6).Value = vbChecked Then
            ilMissedType = ilMissedType + 4
        End If
        If ilMissedType > 0 Then    'Include missed
            If ilBillType = 0 Then
                ilSpotType = 2  'Missed only
            Else
                ilSpotType = 3  'Both
            End If
        Else
            ilSpotType = 1  'Scheduled only
        End If
    Else
        ilBillType = 3          'include bill & unbilled
        ilIncludePSA = True
        ilMissedType = 0
        ilSpotType = 1  'Scheduled only
    End If

    ilByOrderOrAir = 1
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCpfRecLen = Len(tmCpf)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
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
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMcfRecLen = Len(tmMcf)
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)


    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmCpr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCpr
        btrDestroy hmCff
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCprRecLen = Len(tmCpr)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmCpr)
        btrDestroy hmFsf
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmCpr
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmCpr)
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmLcf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmCpr
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)

    tmAdf.iCode = 0
    tmVef.iCode = 0

    '*                        CPR.BTR Database Field Cross Reference
    '*
    '*  iGenDate    =   Generation Date for Crystal to key on
    '*  lGenTime    =   Generation Time for Crystal to key on
    '*  ivefCode    =   Vehicle Code
    '*  iSpotDate   =   Spot Date on contract
    '*  lSpotTime   =   Spot Time on contract
    '*  lCntrNo     =   Contract number
    '*  iLineNo     =   Line number on contract
    '*  iLen        =   Spot length
    '*  sProduct    =   Spot Advertiser and Product
    '*  sZone       =   Time zone for spot
    '*  sCartNo     =   Cart number for spot
    '*  sISCI       =   ISCI number for spot
    '*  sCreative   =   Vehicle name for Status where veh name is different from current veh
    '*                  This occurs on MG and Outside status
    '*  iRemoteID   =   7-17-09 changed to use for game no to link to sub report
    '*  iMissing    =   Scheduled or Not
    '*  lHd1CefCode =   Spot amount in dollars
    '*  lFt1CefCode =   Feed spot code
    '*  iStatus     =   Status
    '*  iUnassign   =   Exception posting audit trail (tmsdf.sAffchg)
    '*  lFt2CefCode =   7-17-09 Billed or not billed (used to be Cpr.iRemoteID)

    If Not ilISCIOnly Then
        If Not gSetFormula("ReportHeader", "'" & "Log Posting Status" & "'") Then
            Exit Sub
        End If
    Else
        If Not gSetFormula("ReportHeader", "'" & "Missing ISCI Codes" & "'") Then
            Exit Sub
        End If
    End If
    ilCntrSpots = gSetCheck(RptSel!ckcSelC10(0).Value)       'include local spots (vs network feed)
    ilFeedSpots = gSetCheck(RptSel!ckcSelC10(1).Value)       'include network (feed) vs local

    'For ilLoop = 1 To 4 Step 1
    '    ilAllNoSpots(ilLoop) = 0
    '    slAllDollars(ilLoop) = ".00"
    'Next ilLoop
    For ilVehicle = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
        If RptSel!lbcSelection(0).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmCpr.iVefCode = ilVefCode
            ReDim tmDayIsComplete(0 To 0) As DAYISCOMPLETE
            ReDim tmPLSdf(0 To 0) As SPOTTYPESORT
            ilVefIndex = gBinarySearchVef(ilVefCode)
            If ilVefIndex >= 0 Then
                '5-9-11 Remove all the invalid bb spots that doesnt belong
                ilGameSelect = 0
                ilLineSelect = 0
                llChfSelect = 0
                ilRet = gRemoveBBSpots(hmSdf, ilVefCode, ilGameSelect, slStartDate, slEndDate, llChfSelect, ilLineSelect)

                If RptSel!rbcSelC8(1).Value = True And tgMVef(ilVefIndex).sType <> "G" Then
                    mFindNonGamePostStatus slStartDate, slEndDate, ilVefCode
                Else
                    mObtainSdf ilVefCode, slStartDate, slEndDate, ilSpotType, ilBillType, ilIncludePSA, ilMissedType, ilISCIOnly, ilCostType, ilByOrderOrAir, ilCntrSpots, ilFeedSpots
                End If
            End If
                
            ilUpper = UBound(tmPLSdf)
            If ilUpper > 0 Then
                ArraySortTyp fnAV(tmPLSdf(), 0), ilUpper, 0, LenB(tmPLSdf(0)), 0, LenB(tmPLSdf(0).sKey), 0
            End If
            'For ilLoop = 1 To 4 Step 1
            '    ilVehNoSpots(ilLoop) = 0
            '    slVehDollars(ilLoop) = ".00"
            'Next ilLoop
            ilNoDays = 0
            'For ilLoop = 1 To 4 Step 1
            '    ilDayNoSpots(ilLoop) = 0
            '    slDayDollars(ilLoop) = ".00"
            'Next ilLoop

            For ilIndex = LBound(tmPLSdf) To UBound(tmPLSdf) - 1
                tmSdf = tmPLSdf(ilIndex).tSdf
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                'date
                If tmSdf.sXCrossMidnight = "Y" Then         '5-26-06 if cross mid night spot, adjust the date aired
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate     'to the true date of airing
                    gPackDateLong llDate + 1, tmCpr.iSpotDate(0), tmCpr.iSpotDate(1)
                Else
                    tmCpr.iSpotDate(0) = tmSdf.iDate(0)                 'contr start date
                    tmCpr.iSpotDate(1) = tmSdf.iDate(1)
                End If
                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slStr
                tmCpr.iSpotTime(0) = tmSdf.iTime(0)
                tmCpr.iSpotTime(1) = tmSdf.iTime(1)
                'slStr = Trim$(Str$(tmSdf.iLen))
                tmCpr.iLen = tmSdf.iLen     ' (slStr)         'spot length

                '3-28-05 determine if bb spot
                tmCpr.iAdfCode = 0                          'use field to store bb flag,
                If tmSdf.sSpotType = "O" Then               'open bb
                    tmCpr.iAdfCode = 1
                ElseIf tmSdf.sSpotType = "C" Then           'close bb
                    tmCpr.iAdfCode = 2
                End If

                If tmAdf.iCode <> tmSdf.iAdfCode Then       'dont reread advt if already in memory
                    tmAdfSrchKey.iCode = tmSdf.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If

               tmCpr.iRemoteID = tmSdf.iGameNo              'swap remoteID with tmcpr.lft2cefcode for the game #.  It must be integer to link to subreport
                '7-22-04 obtain either contract or feed spot information
                tmCpr.lFt1CefCode = 0                   'used to store feed spot code
                If tmSdf.lChfCode = 0 Then              'feed spot
                    tmChfSrchKey.lCode = tmSdf.lFsfCode
                    ilRet = btrGetEqual(hmFsf, tmFsf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slContrNumber = ""
                    tmCpr.lFt1CefCode = tmSdf.lFsfCode
                    'obtain feed product
                    tmPrfSrchKey.lCode = tmFsf.lPrfCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slProduct = tmPrf.sName
                    Else
                        slProduct = ""
                    End If
                Else                                    'contract spot
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slContrNumber = Trim$(str$(tmChf.lCntrNo))
                    tmClfSrchKey.lChfCode = tmChf.lCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F")) 'And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    slProduct = tmChf.sProduct
                End If
                mObtainCopy slCopyProduct, slCopyZone, slCopyCart, slCopyISCI, slProduct
                tmCpr.sZone = slCopyZone
                tmCpr.sISCI = slCopyISCI
                tmCpr.sCartNo = slCopyCart

                If Trim$(slCopyProduct) = "" Then
                    slStr = Trim$(tmAdf.sName) & ", " & Trim$(slProduct)
                Else
                    slStr = Trim$(tmAdf.sName) & ", " & Trim$(slCopyProduct)
                End If
                'adv/prod
                tmCpr.sProduct = slStr
                 If ((slContrNo = "") Or (slContrNo = slContrNumber)) Then   'show if selective cntr matches the spot found, or all spots if no selective entered
                    'contract #
                    tmCpr.lCntrNo = Val(slContrNumber)
                    tmCpr.iLineNo = tmSdf.iLineNo           'sch line #
                    slStr = ""
                    slSpotAmount = ".00"
                    tmSmf.iOrigSchVef = tmSdf.iVefCode
                    ilRet = BTRV_ERR_NONE               'default in case SMF not retrieved
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        '11-30-04 change access of smf from key0 to key2 for speed
                        'tmSmfSrchKey.lFsfCode = tmSdf.lFsfCode
                        'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                        'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                        'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                        'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                        'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                        tmSmfSrchKey2.lCode = tmSdf.lCode
                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                        'Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                        '    If (tmSmf.lSdfCode = tmSdf.lCode) And (tmSmf.lFsfCode = tmSdf.lFsfCode) Then
                        '        Exit Do
                        '    End If
                        '    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        'Loop

                    End If

                    If ilRet <> BTRV_ERR_NONE Then          'bad record pointer
                        'assume same orig vef code
                        tmSmf.iOrigSchVef = tmSdf.iVefCode
                        slStatus = "SDFSMF Err:" & str(tmSdf.lCode)     'SMF missing with matching SDFcode
                        slSpotAmount = ".00"
                    Else
                        If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                            tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        tmCpr.iReady = 0  'initialize

                        If tmSdf.lChfCode = 0 Then          'feed spot
                            slSpotAmount = ".00"
                            tmCpr.iReady = 10               'flag for crystal to show Feed
                        Else                                'contract spot, determine price
                            slStatus = ""                                       '9-1-00
                            If tmClf.sType = "H" Then       '9-1-00
                                slStatus = "(H)"
                            End If
                            If tmSdf.sSpotType = "X" Then
                                'If tmSdf.sPriceType = "N" Then
                                '3-23-03 Test the advt instead of spot to determine a fill or extra


                                '1-19-04 change way in which fill/extra are shown
                                slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)

                                'If tmAdf.sBonusOnInv = "N" Then
                                If slShowOnInv = "N" Then
                                    tmCpr.iReady = 1
                                    slStr = "-Fill"
                                Else
                                    tmCpr.iReady = 2
                                    'slStr = "+Extra"
                                    slStr = "+Fill"
                                End If
                            Else
                                ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slOrdered)
                                If tgPriceCff.sPriceType = "N" Then
                                    tmCpr.iReady = 3
                                ElseIf tgPriceCff.sPriceType = "M" Then
                                    tmCpr.iReady = 4
                                ElseIf tgPriceCff.sPriceType = "B" Then
                                    tmCpr.iReady = 5
                                ElseIf tgPriceCff.sPriceType = "S" Then
                                    tmCpr.iReady = 6
                                ElseIf tgPriceCff.sPriceType = "P" Then
                                    tmCpr.iReady = 7
                                ElseIf tgPriceCff.sPriceType = "R" Then
                                    tmCpr.iReady = 8
                                ElseIf tgPriceCff.sPriceType = "A" Then
                                    tmCpr.iReady = 9
                                Else
                                    slStr = slOrdered
                                    If InStr(slOrdered, ".") = 0 Then    'didnt find period
                                        slSpotAmount = ".00"
                                    Else
                                        slSpotAmount = slOrdered
                                    End If
                                End If
                            End If
                        End If
                    End If                 'btrv_err_none

                    'price slOrdered
                    'tmCpr.lHd1CefCode = Val(slSpotAmount)
                    tmCpr.lHd1CefCode = gStrDecToLong(slSpotAmount, 2)
                    tmCpr.sCreative = ""
                    gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slCode  '9-1-00

                    '9-1-00 add "H" notation to statuses if applicable
                    If tmSdf.sSchStatus = "S" Then
                        slStr = "Scheduled" & " " & slStatus
                    ElseIf tmSdf.sSchStatus = "M" Then
                        slStr = "Missed" & " " & slStatus
                    ElseIf tmSdf.sSchStatus = "R" Then
                        slStr = "Ready" & " " & slStatus
                    ElseIf tmSdf.sSchStatus = "U" Then
                        slStr = "UnSched" & " " & slStatus
                    ElseIf tmSdf.sSchStatus = "G" Then
                        If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                            slStr = "MG " & Trim$(slStatus) & Trim$(slCode) & Chr$(10) & Trim$(tlVef.sName)
                        Else
                            slStr = "MG " & Trim$(slStatus) & Trim$(slCode)
                        End If
                    ElseIf tmSdf.sSchStatus = "O" Then
                        If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                            slStr = "Out " & Trim$(slStatus) & Trim$(slCode) & Chr$(10) & Trim$(tlVef.sName)
                            'tmCpr.sCreative = Trim$(tlVef.sName)  9-1-00
                        Else
                            slStr = "Out " & Trim$(slStatus) & Trim$(slCode)
                        End If
                    ElseIf tmSdf.sSchStatus = "C" Then
                        slStr = "Cancelled" & " " & slStatus
                    ElseIf tmSdf.sSchStatus = "H" Then
                        slStr = "Hidden" & " " & Trim$(slCode)
                    ElseIf tmSdf.sSchStatus = "A" Then
                        slStr = "On Alt" & " " & Trim$(slCode)
                    ElseIf tmSdf.sSchStatus = "B" Then
                        slStr = "On Alt & MG" & " " & Trim$(slCode)
                    End If

                    'status
                    tmCpr.sCreative = slStr         '9-1-00
                    tmCpr.sStatus = tmSdf.sSchStatus
                    ' When we sort in Crystal scheduled, outide and makegood orders are first then any others follow
                    If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G")) Then
                        tmCpr.iMissing = True
                    Else
                        tmCpr.iMissing = False
                    End If
                    If tmSdf.sBill = "Y" Then
                        '7-17-09 changed to use long; need the integer for game number info to link to subreport
                        tmCpr.lFt2CefCode = 1 'Billed
                    Else
                        tmCpr.lFt2CefCode = 0  'Not Billed
                    End If
                    'billed

                    tmCpr.iUnassign = 0     '1-24-03 change to use new field to store tmsdf.saffchg
                    If tmSdf.sAffChg = "C" Then         'C, T B are currently unused, but dont remove
                        tmCpr.iUnassign = 3
                    ElseIf tmSdf.sAffChg = "T" Then
                        tmCpr.iUnassign = 4
                    ElseIf tmSdf.sAffChg = "B" Then
                        tmCpr.iUnassign = 2
                    ElseIf tmSdf.sAffChg = "A" Then
                        tmCpr.iUnassign = 1
                    ElseIf tmSdf.sAffChg = "Y" Then
                        tmCpr.iUnassign = 5
                    End If

                    'build array of all the day is complete flags for the days/games/vehicle
                    ilfoundDay = False
                    
                    For ilUpperDay = LBound(tmDayIsComplete) To UBound(tmDayIsComplete) - 1
                        If tmSdf.iDate(0) = tmDayIsComplete(ilUpperDay).iDate(0) And tmSdf.iDate(1) = tmDayIsComplete(ilUpperDay).iDate(1) And tmSdf.iGameNo = tmDayIsComplete(ilUpperDay).iGameNo Then
                            ilfoundDay = True
                            Exit For
                        End If
                    Next ilUpperDay
                    
                    tmCpr.sLive = "N"           'default to not posted
                    If ilfoundDay Then
                        tmCpr.sLive = tmDayIsComplete(ilUpperDay).sAffPost
                    Else
                        tmDayIsComplete(ilUpperDay).iDate(0) = tmSdf.iDate(0)
                        tmDayIsComplete(ilUpperDay).iDate(1) = tmSdf.iDate(1)
                        tmDayIsComplete(ilUpperDay).iGameNo = tmSdf.iGameNo
                        tmDayIsComplete(ilUpperDay).iVefCode = tmSdf.iVefCode
                        tmDayIsComplete(ilUpperDay).sAffPost = "N"
                        'get the day is complete status from LCF
                        tmLcfSrchKey0.iLogDate(0) = tmSdf.iDate(0)
                        tmLcfSrchKey0.iLogDate(1) = tmSdf.iDate(1)
                        tmLcfSrchKey0.iSeqNo = 1
                        tmLcfSrchKey0.iType = tmSdf.iGameNo
                        tmLcfSrchKey0.iVefCode = tmSdf.iVefCode
                        tmLcfSrchKey0.sStatus = "C"
                        ilRet = btrGetEqual(hmLcf, tmLcf, Len(tmLcf), tmLcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                        If ilRet = BTRV_ERR_NONE Then
                            tmDayIsComplete(ilUpperDay).sAffPost = tmLcf.sAffPost
                            tmCpr.sLive = tmLcf.sAffPost
                        End If
                        ReDim Preserve tmDayIsComplete(LBound(tmDayIsComplete) To UBound(tmDayIsComplete) + 1)
                    End If

                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                    tmCpr.lGenTime = lgNowTime
                    tmCpr.iGenDate(0) = igNowDate(0)
                    tmCpr.iGenDate(1) = igNowDate(1)

                    ilRet = btrInsert(hmCpr, tmCpr, imCprRecLen, INDEXKEY0)
                End If
            Next ilIndex
        End If
    Next ilVehicle
    'Erase ilDayNoSpots
    'Erase slDayDollars
    'Erase ilVehNoSpots
    'Erase slVehDollars
    'Erase ilAllNoSpots
    'Erase slAllDollars
    Erase tmPLSdf, tmDayIsComplete
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmMcf)
    ilRet = btrClose(hmTzf)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmCpr
    btrDestroy hmFsf
    btrDestroy hmSmf
    btrDestroy hmAgf
    btrDestroy hmLcf
    btrDestroy hmMcf
    btrDestroy hmTzf
    btrDestroy hmCpf
    btrDestroy hmCif
    btrDestroy hmSdf
    btrDestroy hmVsf
    btrDestroy hmVef
    btrDestroy hmAdf
    btrDestroy hmClf
    btrDestroy hmCHF
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gStudioLogRpt                   *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate Studio Log report     *
'*      5-14-02 DH Convert from Bridge to Crystal
'*                                                     *
'*******************************************************
Sub gStudioLogRpt()
    Dim ilDBEof As Integer
    Dim ilRet As Integer
    Dim hlRead As Integer
    Dim illoop As Integer
    Dim slDallasFile As String
    Dim slLine As String
    Dim llNoRecsToProc As Long
    Dim ilCodeStn As Integer
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilRecIndex As Integer
    Dim ilRepeatIndex As Integer
    Dim ilPage As Integer
    Dim slHour As String
    Dim slMin As String
    Dim slSec As String
    Dim ilTimeAdj As Integer
    Dim ilSameHour As Integer
    Dim slSaveHour As String
    ReDim slField(0 To 3) As String
    Dim llNextTime As Long
    Dim ilAdjustPos As Integer          '5-13-13
   ' Dim ilAnyOutput As Integer

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTxr)
        ilRet = btrClose(hmVef)
        btrDestroy hmTxr
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imTxrRecLen = Len(tmTxr)

    ReDim tmCodeStn(0 To 0) As CODESTNCONV
    ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If Trim$(tmVef.sCodeStn) <> "" Then
            tmCodeStn(UBound(tmCodeStn)).sName = Trim$(tmVef.sName)
            tmCodeStn(UBound(tmCodeStn)).sCodeStn = tmVef.sCodeStn
            ReDim Preserve tmCodeStn(0 To UBound(tmCodeStn) + 1) As CODESTNCONV
        End If
        ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmVef)
        btrDestroy hmMnf
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    ilRet = btrGetFirst(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If (tmMnf.sType = "N") And (Trim$(tmMnf.sCodeStn) <> "") Then
            tmCodeStn(UBound(tmCodeStn)).sName = tmMnf.sName
            tmCodeStn(UBound(tmCodeStn)).sCodeStn = tmMnf.sCodeStn
            ReDim Preserve tmCodeStn(0 To UBound(tmCodeStn) + 1) As CODESTNCONV
        End If
        ilRet = btrGetNext(hmMnf, tmMnf, imMnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'llNoRecsToProc = 0
    For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
        If RptSel!lbcSelection(1).Selected(illoop) Then
            slDallasFile = RptSel!lbcSelection(1).List(illoop)
            llNoRecsToProc = llNoRecsToProc + FileLen(sgExportPath & slDallasFile) / 106
        End If
    Next illoop
    'slFileName = sgRptPath & slName & Chr$(0)
    'gStudioLg

    'ilAnyOutput = False
    'slLangText = "Printing..."
    'If ilPreview <> 0 Then
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
    '    ilDummy = LlPreviewSetTempPath(hdJob, sgRptSavePath)
    '    ilDummy = LlPreviewSetResolution(hdJob, 200)
     'Else
    '    ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, RptSel.hWnd, slLangText)
     'End If
     'DoEvents
     'If (ilRet = 0) Then
     '   ilRecNo = 1
     '   ilErrorFlag = 0

     '   slPrinter = LlVBPrintGetPrinter(hdJob)
     '   slPort = LlVBPrintGetPort(hdJob)
        ilDBEof = False
        For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
            If RptSel!lbcSelection(1).Selected(illoop) Then
                slDallasFile = RptSel!lbcSelection(1).List(illoop)
                ilDBEof = False
                ReDim tmDallasFdSort(0 To 0) As DALLASFDSORT
                ilUpper = 0
                'On Error GoTo mStudioLogErr:
                'hlRead = FreeFile
                'Open sgExportPath & slDallasFile For Input Access Read As hlRead
                ilRet = gFileOpen(sgExportPath & slDallasFile, "Input Access Read", hlRead)
                If ilRet <> 0 Then
                    ilDBEof = True
                End If
                If Not ilDBEof Then
                    ilDBEof = False
                    err.Clear
                    Do
                        If EOF(hlRead) Then
                            ilDBEof = True
                        Else
                            Line Input #hlRead, slLine
                        End If
                        If err.Number <> 0 Then
                            ilDBEof = True
                        End If
                        If (Asc(slLine) = 26) Then    'Ctrl Z
                            ilDBEof = True
                        End If
                        If ilDBEof Then
                            Exit Do
                        End If
                        'Build sort Array
                        If Left$(slLine, 1) = "A" Then
                            slStr = Mid$(slLine, 10, 5)
                            ilFound = False
                            For ilCodeStn = 0 To UBound(tmCodeStn) - 1 Step 1
                                If StrComp(Trim$(tmCodeStn(ilCodeStn).sCodeStn), Trim$(slStr), 1) = 0 Then
                                    ilFound = True
                                    slStr = Mid$(slLine, 16, 6)
                                    tmDallasFdSort(ilUpper).sKey = tmCodeStn(ilCodeStn).sName & " " & slStr
                                    tmDallasFdSort(ilUpper).sRecord = slLine
                                    ilUpper = ilUpper + 1
                                    ReDim Preserve tmDallasFdSort(0 To ilUpper) As DALLASFDSORT
                                End If
                            Next ilCodeStn
                            If Not ilFound Then
                                slStr = Mid$(slLine, 10, 6)
                                tmDallasFdSort(ilUpper).sKey = "Missing " & " " & slStr
                                tmDallasFdSort(ilUpper).sRecord = slLine
                                ilUpper = ilUpper + 1
                                ReDim Preserve tmDallasFdSort(0 To ilUpper) As DALLASFDSORT
                            End If
                        End If
                    Loop While Not ilDBEof
                End If
                'outer loop - one loop per page
                If ilUpper > 0 Then
                    ArraySortTyp fnAV(tmDallasFdSort(), 1), ilUpper - 1, 0, LenB(tmDallasFdSort(0)), 0, LenB(tmDallasFdSort(0).sKey), 0 '100, 0

                    ilDBEof = False
                    'ilDummy = LlDefineVariableExt(hdJob, "Logo", sgLogoPath & "RptLogo.Bmp", LL_DRAWING, "")
                    'ilDummy = LLDefineVariableExtHandle(hdJob, "CSILogo", Traffic!imcCSILogo, LL_DRAWING_HBITMAP, "")
                Else
                    ilDBEof = True
                End If
                ilRecIndex = 0
                ilPage = 1
                While (Not ilDBEof) 'And ilErrorFlag = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
                    'ilAnyOutput = True
                    'tmTxrLog.sVehicle = Trim$(Left$(tmDallasFdSort(ilRecIndex).sKey, 20))
                    'ilDummy = LlDefineVariableExt(hdJob, "Vehicle", slStr, LL_TEXT, "")
                    'slStr = Mid$(tmDallasFdSort(ilRecIndex).sRecord, 5, 2) & "/" & Mid$(tmDallasFdSort(ilRecIndex).sRecord, 7, 2) & "/" & Mid$(tmDallasFdSort(ilRecIndex).sRecord, 3, 2)
                    'tmTxrLog.sLogDate = Format$(gDateValue(slStr), "dddd") & ", " & slStr 'gAddDayToDate(slStr)

                    'ilDummy = LlDefineVariableExt(hdJob, "Date", slStr, LL_TEXT, "")
                    'slStr = Trim$(Str$(ilPage))
                    'ilDummy = LlDefineVariableExt(hdJob, "Page", slStr, LL_TEXT, "")
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
                    'ilRet = LLPrint(hdJob)
                    'ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
                   'ilPage = ilPage + 1
                While (Not ilDBEof) 'And ilErrorFlag = 0 And ilRet = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
                        ilRepeatIndex = ilRecIndex
                        slField(0) = ""
                        slField(1) = ""
                        slField(2) = ""
                        slField(3) = ""
                        tmTxrLog.sTime = ""
                        tmTxrLog.sCopy = ""
                        tmTxrLog.sLen = ""
                        tmTxrLog.sAdvtProd = ""
                        tmTxrLog.sVehicle = ""
                        tmTxrLog.sLogDate = ""
                        tmTxrLog.sTimeKey = ""

                        ilTimeAdj = False   'True
                        ilSameHour = False
                        Do
                            'Call DefineFields
                            'LlDefineFieldStart hdJob
                            slHour = Mid$(tmDallasFdSort(ilRecIndex).sRecord, 16, 2)
                            tmTxrLog.iFiller4 = Val(slHour)
                            slMin = Mid$(tmDallasFdSort(ilRecIndex).sRecord, 18, 2)
                            slSec = Mid$(tmDallasFdSort(ilRecIndex).sRecord, 20, 2)
                            slSaveHour = slHour
                            If Val(slHour) = 0 Then
                                slStr = "12"
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "AM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "AM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "AM"
                                End If
                            ElseIf Val(slHour) < 12 Then
                                slStr = Trim$(str$(Val(slHour)))
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "AM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "AM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "AM"
                                End If
                            ElseIf Val(slHour) = 12 Then
                                slStr = "12"
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "PM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "PM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "PM"
                                End If
                            Else
                                slStr = Trim$(str$(Val(slHour) - 12))
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "PM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "PM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "PM"
                                End If
                            End If
                            If ilTimeAdj Then
                                'If slField(0) = "" Then
                                    slField(0) = " "
                                'Else
                                '    slField(0) = slField(0) & Chr(10) & " "
                                'End If
                            Else
                                'If slField(0) = "" Then
                                    slField(0) = slStr
                                'Else
                                '    slField(0) = slField(0) & Chr(10) & slStr
                                'End If
                            End If

                            llNextTime = CLng(gTimeToCurrency(slStr, False))
'                            slStr = Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 70, 35)) & ", " & Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 34, 35))
'                            'If slField(1) = "" Then
'                                slField(1) = slStr
                            'Else
                            '    slField(1) = slField(1) & Chr(10) & slStr
                            'End If
                            slStr = Trim$(str$(Val(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 23, 4))))
                            llNextTime = llNextTime + Val(slStr)
                            'If slField(2) = "" Then
                                slField(2) = slStr
                            'Else
                            '    slField(2) = slField(2) & Chr(10) & slStr
                            'End If
                            ilAdjustPos = InStr(28, tmDallasFdSort(ilRecIndex).sRecord, " ")
                            If ilAdjustPos > 0 Then
                                slStr = Trim$(Mid(tmDallasFdSort(ilRecIndex).sRecord, 28, ilAdjustPos - 28))
                            Else
                                ilAdjustPos = 0         '5-13-13 size of copy field may vary
                            End If
                            'ADjustment to find Advertiser & product
                            ilAdjustPos = ilAdjustPos - 28 + 1 'len of copy, plus space to get to next field
                            'slStr = Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 28, 5))
                            
                            'If slField(3) = "" Then
                                slField(3) = slStr
                            'Else
                            '    slField(3) = slField(3) & Chr(10) & slStr
                            'End If
                            'If ilRecIndex + 1 >= UBound(tmDallasFdSort) - 1 Then
                            '    Exit Do
                            'End If
                            'If Trim$(Left$(tmDallasFdSort(ilRecIndex).sKey, 20)) <> Trim$(Left$(tmDallasFdSort(ilRecIndex + 1).sKey, 20)) Then
                            '    Exit Do
                           'End If
                            'Test time- if not adjacent- exit
                            slHour = Mid$(tmDallasFdSort(ilRecIndex + 1).sRecord, 16, 2)
                            slMin = Mid$(tmDallasFdSort(ilRecIndex + 1).sRecord, 18, 2)
                            slSec = Mid$(tmDallasFdSort(ilRecIndex + 1).sRecord, 20, 2)
                            If Val(slHour) = 0 Then
                                slStr = "12"
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "AM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "AM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "AM"
                                End If
                            ElseIf Val(slHour) < 12 Then
                                slStr = Trim$(str$(Val(slHour)))
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "AM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "AM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "AM"
                                End If
                            ElseIf Val(slHour) = 12 Then
                                slStr = "12"
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "PM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "PM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "PM"
                                End If
                            Else
                                slStr = Trim$(str$(Val(slHour) - 12))
                                If (Val(slMin) = 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & "PM"
                                ElseIf (Val(slMin) <> 0) And (Val(slSec) = 0) Then
                                    slStr = slStr & ":" & slMin & "PM"
                                Else
                                    slStr = slStr & ":" & slMin & ":" & slSec & "PM"
                                End If
                            End If
                            If Val(slSaveHour) <> Val(slHour) Then
                                ilSameHour = False
                                'Exit Do
                            Else
                                ilSameHour = True
                            End If
                            If CLng(gTimeToCurrency(slStr, False)) > llNextTime Then
                                ilTimeAdj = False
                            Else
                                ilTimeAdj = True
                            End If

                            'write out one spot to Text file
                            tmTxrLog.sVehicle = Trim$(Left$(tmDallasFdSort(ilRecIndex).sKey, 20))
                            slStr = Mid$(tmDallasFdSort(ilRecIndex).sRecord, 5, 2) & "/" & Mid$(tmDallasFdSort(ilRecIndex).sRecord, 7, 2) & "/" & Mid$(tmDallasFdSort(ilRecIndex).sRecord, 3, 2)
                            tmTxrLog.sLogDate = Format$(gDateValue(slStr), "dddd") & ", " & slStr 'gAddDayToDate(slStr)
                            'slStr = Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 70, 35)) & ", " & Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 34, 35))
                            '5-13-13 adjust for different lengths of copy fiel
                            slStr = Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 28 + ilAdjustPos + 35, 35)) & ", " & Trim$(Mid$(tmDallasFdSort(ilRecIndex).sRecord, 28 + ilAdjustPos, 35))
                            slField(1) = slStr
                            
                            tmTxrLog.sTime = slField(0)
                            tmTxrLog.sCopy = slField(3)         'copy (max 10)
                            tmTxrLog.sLen = slField(2)
                            tmTxrLog.sAdvtProd = slField(1)

                            LSet tmTxr = tmTxrLog

                            tmTxr.iGenDate(0) = igNowDate(0)
                            tmTxr.iGenDate(1) = igNowDate(1)
                            tmTxr.lGenTime = lgNowTime

                            ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
                            ilRecIndex = ilRecIndex + 1

                        Loop While ilSameHour
                        'ilDummy = LLDefineFieldExt(hdJob, "Time", slField(0), LL_TEXT, "")
                        'ilDummy = LLDefineFieldExt(hdJob, "CartNo", slField(3), LL_TEXT, "")
                        'ilDummy = LLDefineFieldExt(hdJob, "Length", slField(2), LL_TEXT, "")
                        'ilDummy = LLDefineFieldExt(hdJob, "AdvtProd", slField(1), LL_TEXT, "")
                        'ilDummy = LLDefineFieldExt(hdJob, "Notes", "", LL_TEXT, "")
                        'notify the user (how far have we come?)
                        'ilRet = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * ilRecNo / llNoRecsToProc))
                        'DoEvents
                        'tell L&L to print the table line
                        'If ilRet = 0 Then
                        '    ilRet = LlPrintFields(hdJob)
                        'End If

                        'next data set if no error or warning
                        'If ilRet = 0 Then
                            'ilRecIndex = ilRecIndex + 1
                            If ilRecIndex >= UBound(tmDallasFdSort) - 1 Then
                                ilDBEof = True
                            Else
                                If Trim$(Left$(tmDallasFdSort(ilRecIndex - 1).sKey, 20)) <> Trim$(Left$(tmDallasFdSort(ilRecIndex).sKey, 20)) Then
                                    ilRepeatIndex = ilRecIndex
                        '            ilRet = LL_WRN_REPEAT_DATA  'Force next page
                        '            ilPage = 1
                                Else
                        '            'Page eject if across 12n
                                    If (Val(slSaveHour) <= 11) And (Val(slHour) >= 12) Then
                                        ilRepeatIndex = ilRecIndex
                        '                ilRet = LL_WRN_REPEAT_DATA
                                    End If
                                End If
                            End If
                        '
                            'ilRecNo = ilRecNo + ilRecIndex - ilRepeatIndex
                        'End If
                    Wend  ' inner loop

                    'if error or warning: different reactions:
                    'If ilRet < 0 Then
                    '    If ilRet <> LL_WRN_REPEAT_DATA Then
                    '        ilErrorFlag = ilRet
                    '    Else
                    '        ilRecIndex = ilRepeatIndex
                    '    End If
                    'End If

                Wend    ' while not EOF
                Close hlRead
            End If
        Next illoop
        'end print
        'If Not ilAnyOutput Then
        '    ilRet = LLPrintEnableObject(hdJob, ":Spots", False)
       '     ilRet = LLPrint(hdJob)
       ' Else
       '     ilRet = LLPrintEnableObject(hdJob, ":Spots", True)
       ' End If
       ' ilRet = LlPrintEnd(hdJob, 0)

        'in case of preview: show the preview
        'If ilPreview <> 0 Then
        '    If ilErrorFlag = 0 Then
        '        ilDummy = LlPreviewDisplay(hdJob, slFileName, sgRptSavePath, RptSel.hWnd)
        '    Else
        '        mErrMsg ilErrorFlag
        '    End If
        '    ilDummy = LlPreviewDeleteFiles(hdJob, slFileName, sgRptSavePath)
        'Else
        '    If ilErrorFlag <> 0 Then
        '        mErrMsg ilErrorFlag
        '    End If
        'End If

    'Else  ' LlPrintWithBoxStart
    '    ilErrorFlag = ilRet
    '    ilRet = LlPrintEnd(hdJob, 0)
    '    mErrMsg ilErrorFlag
    'End If  ' LlPrintWithBoxStart

    Erase slField
    Erase tmCodeStn
    Erase tmDallasFdSort

    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmTxr)
    btrDestroy hmMnf
    btrDestroy hmVef
    btrDestroy hmTxr

    'End If  ' LlSelecTLileDlgTitle
    Exit Sub
'mStudioLogErr:
'    ilDBEof = True
'    Resume Next
End Sub
'*********************************************************************
'*
'*      Procedure Name:gAssignCopyTest
'*
'*             Created:7/19/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Assign copy test (taken from
'*                      gAssignCopyToSpots)
'*
'*      9/25/98 - Search spot & package vehicle for copy
'*      6/21/99 - dont test spots if spot has been chged
'*                (indicating in the past) and there
'*                is a pointertype
'       8-3-04 Option to include/exclude contract/feed spots
'       12-15-04 Implement Feed spots:  if spot has a copy pointer, copy exists
'                and no further testing required.  There are no rotations
'                attached to a feed spot.
'*      3-29-05 test for bb spots and ignore reading SSF for them; find bb copy rotation
'***********************************************************************
Function mAssignCopyTest(ilSSFType As Integer, ilVpfIndex As Integer, ilAsgnVefCode As Integer, ilRotNo As Integer, ilNonRegionDefined As Integer, ilRegionMissing As Integer, ilRegionSuperseded As Integer, ilRegionRotNo As Integer, ilCntrSpots As Integer, ilFeedSpots As Integer, ilIncludeUnAssg As Integer, ilIncludeReassg As Integer, slLive As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   ilRet = gAssignCopyTest( slSsfType, ilVpfIndex)
'
'   Where:
'       slSsfType(I) "O"=On Air; "A" = Alternate
'       ilVpfIndex(I): Vehicle index into tgVpf
'       ilAsgnVefCode(O): Vehicle code that requires copy
'       ilNonRegionDefined(O): True=Non-Regional copy found; False=Non-Regional not found
'       ilRegionMissing(O): True=Contract defined with regional copy other then spot contract
'       ilRegionSuperseded(0)= 0=None Defined; 1=Copy defined but not assigned; 2=Copy Assigned but superseded; 4=Copy assignment Ok
'       ilRet(O)- 0=None Defined; 1=Copy defined but not assigned;
'                 2=Copy Assigned but superseded; 3=Zone copy Missing;
'                 4=Copy assignment Ok
'
'       tmSdf(I)
'       ilCntrSpots - true to include contract spots
'       ilFeedSpots - true to include network feed spots
'       ilIncludeUnAssg - true to include unassigned
'       ilIncludeReassg - true to include to be reassigned (superceded)
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llSpotTime As Long
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim ilAsgnDate0 As Integer
    Dim ilAsgnDate1 As Integer
    Dim illoop As Integer
    Dim ilIndex As Integer
    Dim slType As String
    Dim ilDay As Integer
    Dim llAvailTime As Long
    Dim ilAvailOk As Integer
    Dim ilEvtIndex As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    Dim ilMatch As Integer
    Dim ilCrfVefCode As Integer
    Dim ilBypassCrf As Integer
    Dim ilVpf As Integer
    Dim ilVef As Integer
    Dim slSpotDate As String
    Dim ilLoopOnCrf As Integer

    'Dim imNoZones As Integer
    'ReDim slZone(1 To 6) As String * 3
    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate  'moved to first instr to test the spot date against todays date
    slSpotDate = slDate
    llDate = gDateValue(slDate)
    ilAsgnDate0 = tmSdf.iDate(0)
    ilAsgnDate1 = tmSdf.iDate(1)
    ilDay = gWeekDayStr(slDate)
    ilRegionRotNo = -1
    If ((tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3")) And (tmSdf.sAffChg = "Y" Or tmSdf.sAffChg = "B" Or tmSdf.sAffChg = "C") Then
        If llDate <= lmTodayDate Then           'date in past
            mAssignCopyTest = 4
            'Region test added: 8-4-00 (Start Point 0)
            ilRegionMissing = False
            ilNonRegionDefined = True
            'Region test added: 8-4-00 (End Point 0)
            Exit Function
        End If
    End If

    '12-15-04 if feed spot, if there is a copy pointer there is no need to check for unassigned, ready to assign.
    'Copy is missing if no copy  pointer.
    If tmSdf.lFsfCode > 0 Then          'feed spot
        If tmSdf.sPtType = 1 Then       'copy defined
            mAssignCopyTest = 4
            ilRegionMissing = False     'region copy N/A
            ilNonRegionDefined = True
        Else                            'no copy defined
            mAssignCopyTest = 0
            ilRegionMissing = False     'region copy N/A
            ilNonRegionDefined = True
        End If
        Exit Function
    End If

    imNoZones = 0
    For illoop = 1 To 6 Step 1
        tmRotNo(illoop).iRotNo = 0
        tmRotNo(illoop).sZone = ""
    Next illoop
    ilRotNo = -1
    'Region test added: 8-4-00 (Start Point 1)
    ilRegionMissing = True
    ilNonRegionDefined = False
    ilRegionSuperseded = -1
    'Region test added: 8-4-00 (End Point 1)
    'gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate     'move to above 6/22/99
    'llDate = gDateValue(slDate)
    'ilAsgnDate0 = tmSdf.iDate(0)
    'ilAsgnDate1 = tmSdf.iDate(1)
    'ilDay = gWeekDayStr(slDate)
    'If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefCode <> tmSdf.iVefCode) Or (tmSsf.iDate(0) <> ilAsgnDate0) Or (tmSsf.iDate(1) <> ilAsgnDate1) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then


    If tmSdf.sSpotType <> "O" And tmSdf.sSpotType <> "C" Then
        If (tmSsf.iType <> ilSSFType) Or (tmSsf.iVefCode <> tmSdf.iVefCode) Or (tmSsf.iDate(0) <> ilAsgnDate0) Or (tmSsf.iDate(1) <> ilAsgnDate1) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
            'tmSsfSrchKey.sType = slSsfType
            tmSsfSrchKey.iType = ilSSFType
            tmSsfSrchKey.iVefCode = tmSdf.iVefCode
            tmSsfSrchKey.iDate(0) = ilAsgnDate0
            tmSsfSrchKey.iDate(1) = ilAsgnDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            imSsfRecLen = Len(tmSsf)
            ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
        Else
            ilRet = BTRV_ERR_NONE
        End If
    End If
    'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
    If ((ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1)) Or (tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C") Then
        If ilSSFType > 0 Then
            If ilSSFType <> tmGsf.iGameNo Then
                tmGsfSrchKey3.iVefCode = tmSsf.iVefCode
                tmGsfSrchKey3.iGameNo = tmSsf.iType
                ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tmSsf.iVefCode) And (tmGsf.iGameNo = tmSsf.iType)
                    If (tmGsf.iAirDate(0) = tmSsf.iDate(0)) And (tmGsf.iAirDate(1) = tmSsf.iDate(1)) Then
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                Loop
                If (ilRet <> BTRV_ERR_NONE) Or (tmGsf.iVefCode <> tmSsf.iVefCode) Or (tmGsf.iGameNo <> tmSsf.iType) Then
                    tmGsf.iHomeMnfCode = -1
                    tmGsf.iVisitMnfCode = -1
                End If
            End If
        End If
        ilEvtIndex = 1
        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
        llSpotTime = CLng(gTimeToCurrency(slTime, False)) ' - 1
        'Find rotation to assign
        'Code later- test spot type to determine which rotation type
        'ilCrfVefCode = gGetCrfVefCode(hmClf, tmSdf)
        'ilAsgnVefCode = ilCrfVefCode
        ilCrfVefCode = ilAsgnVefCode
        ilAvailOk = True
        '3-29-05 see if copy defined for billboards
        If tmSdf.sSpotType = "C" Then
            slType = "C"
        ElseIf tmSdf.sSpotType = "O" Then
            slType = "O"
        Else
            slType = "A"
        End If
        
'        mBuildCRFByCntr ilCrfVefCode, slSpotDate, slType, slLive
        
        'Loop thru array that was build containing rotations to test
'        '3-2-15 change to new key
'        tmCrfSrchKey4.sRotType = slType
'        tmCrfSrchKey4.iEtfCode = 0
'        tmCrfSrchKey4.iEnfCode = 0
'        tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
'        tmCrfSrchKey4.lChfCode = tmSdf.lChfCode
'        tmCrfSrchKey4.lFsfCode = tmSdf.lFsfCode         'feed code
'        'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
'        tmCrfSrchKey4.iRotNo = 32000
'        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get last current record to obtain date
'        Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.lFsfCode = tmSdf.lFsfCode)
        lgMainCvfCount = lgMainCvfCount + 1
        For ilLoopOnCrf = LBound(tmCRFByCntr) To UBound(tmCRFByCntr) - 1
            tmCrf = tmCRFByCntr(ilLoopOnCrf).tCrf
            'Test date, time, day and zone
            ilBypassCrf = False
            'Test if looking for Live or Recorded rotations
    
             If (tmCrf.lRafCode > 0) And (Trim$(tmCrf.sZone) = "R") Then
                 If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY Then
                     'ilVpf = gBinarySearchVpf(tmCrf.iVefCode)
                     'If ilVpf <> -1 Then
                     '    If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                     '        ilBypassCrf = True
                     '    End If
                     'Else
                     '    ilBypassCrf = True
                     'End If
                     ilVef = gBinarySearchVef(ilCrfVefCode)      '5-1-15
                     If ilVef <> -1 Then
                         '5/11/11: Allow selling vehicle to be set to No
                         'If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Then
                         If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "S") Then
                             ilVpf = gBinarySearchVpf(ilCrfVefCode)      '5-1-15
                             If ilVpf <> -1 Then
                                 If tgVpf(ilVpf).sAllowSplitCopy <> "Y" Then
                                     ilBypassCrf = True
                                 End If
                             Else
                                 ilBypassCrf = True
                             End If
                         'ElseIf (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "P") Then
                         ElseIf (tgMVef(ilVef).sType = "P") Then
                             ilBypassCrf = False
                         Else
                             ilBypassCrf = True
                         End If
                     Else
                         ilBypassCrf = True
                     End If
                 End If
            
             End If
           
            If (tmCrf.sDay(ilDay) = "Y") And (tmSdf.iLen = tmCrf.iLen) And (Not ilBypassCrf) And (tmCrf.iVefCode = ilCrfVefCode) Then
                gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                llSAsgnDate = gDateValue(slDate)
                gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                llEAsgnDate = gDateValue(slDate)
                If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                    gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
                    llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                    gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
                    llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                    If tmCrf.sAirGameType = "G" Then
                        llSAsgnTime = 999999
                        llEAsgnTime = -1
                        If tmSsf.iType > 0 Then
                            tmCafSrchKey1.lCrfCode = tmCrf.lCode
                            ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                If tmCaf.sType = "G" Then
                                    If tmCaf.iGameNo = tmSsf.iType Then
                                        llSAsgnTime = llSpotTime
                                        llEAsgnTime = llSpotTime
                                        Exit Do
                                    End If
                                End If
                                ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                    ElseIf tmCrf.sAirGameType = "T" Then
                        llSAsgnTime = 999999
                        llEAsgnTime = -1
                        If tmSsf.iType > 0 Then
                            tmCafSrchKey1.lCrfCode = tmCrf.lCode
                            ilRet = btrGetEqual(hmCaf, tmCaf, imCafRecLen, tmCafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While (ilRet = BTRV_ERR_NONE) And (tmCaf.lCrfCode = tmCrf.lCode)
                                If tmCaf.sType = "T" Then
                                    If (tmCaf.iTeamMnfCode = tmGsf.iHomeMnfCode) Or (tmCaf.iTeamMnfCode = tmGsf.iVisitMnfCode) Then
                                        llSAsgnTime = llSpotTime
                                        llEAsgnTime = llSpotTime
                                        Exit Do
                                    End If
                                End If
                                ilRet = btrGetNext(hmCaf, tmCaf, imCafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                    End If
                    If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                        ilAvailOk = True    'Ok to book into
                        If ((tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O")) And (tmSdf.sSpotType <> "C" And tmSdf.sSpotType <> "O") Then      '3-29-05
                            ilEvtIndex = 1
                            Do
                                If ilEvtIndex > tmSsf.iCount Then
                                    imSsfRecLen = Len(tmSsf)
                                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                        ilEvtIndex = 1
                                    Else
                                        mAssignCopyTest = 0
                                        'Region test added: 8-4-00
                                        '6/8/16: Replaced GoSub
                                        'GoSub lRegionTest
                                        mRegionTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
                                        'Region test added: 8-4-00
                                        Exit Function
                                    End If
                                End If
                                'Scan for avail that matches time of spot- then test avail name
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
                                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                                    'Test time-
                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                                    If llSpotTime = llAvailTime Then
                                        If tmCrf.sInOut = "I" Then
                                            If tmCrf.ianfCode <> tmAvail.ianfCode Then
                                                ilAvailOk = False   'No
                                            End If
                                        Else
                                            If tmCrf.ianfCode = tmAvail.ianfCode Then
                                                ilAvailOk = False   'No
                                            End If
                                        End If

                                        '8-3-04 the Named avail property must allow local spots to be included
                                        tmAnfSrchKey.iCode = tmAvail.ianfCode
                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If (ilRet = BTRV_ERR_NONE) Then
                                            If Not ilCntrSpots And tmAnf.sBookLocalFeed = "L" Then      'Local avail requested to be excluded, exclude if avail type = "L"
                                                ilAvailOk = False
                                            End If
                                            If Not ilFeedSpots And tmAnf.sBookLocalFeed = "F" Then      'Network avail requested to be excluded, exclude if avail type = "F"
                                                ilAvailOk = False
                                            End If
                                        End If
                                        Exit Do
                                    ElseIf (llSpotTime < llAvailTime) And (tmSdf.sXCrossMidnight <> "Y") Then
                                        'Spot missing from Ssf
                                        mAssignCopyTest = 0
                                        'Region test added: 8-4-00
                                        '6/8/16: Replaced GoSub
                                        'GoSub lRegionTest
                                        mRegionTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
                                        'Region test added: 8-4-00
                                        Exit Function
                                    End If
                                End If
                                ilEvtIndex = ilEvtIndex + 1
                            Loop
                        End If
                        'Region test added: 8-4-00 (Start Point 2)
                        If ilAvailOk Then
                            If Trim$(tmCrf.sZone) <> "R" Then
                                ilNonRegionDefined = True
                            Else
                                If ilRegionRotNo = -1 Then
                                    ilRegionRotNo = tmCrf.iRotNo
                                End If
                                ilRegionMissing = False
                                ilAvailOk = False
                            End If
                        End If
                        'Region test added: 8-4-00 (End Point 2)
                        If ilAvailOk Then
                            If ilRotNo = -1 Then
                                ilRotNo = tmCrf.iRotNo
                            End If
                            If Trim$(tmCrf.sZone) = "" Then 'All zones
                                imNoZones = imNoZones + 1
                                tmRotNo(imNoZones).iRotNo = tmCrf.iRotNo
                                tmRotNo(imNoZones).sZone = "Oth"
                                'Add supersede test
                                If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
                                    mAssignCopyTest = 4
                                    'Test if superseded
                                    If tmSdf.sPtType = "1" Then
                                        If imNoZones = 1 Then
                                            If tmRotNo(1).iRotNo > tmSdf.iRotNo Then
                                                mAssignCopyTest = 2
                                            End If
                                        Else
                                            mAssignCopyTest = 3
                                        End If
                                    ElseIf tmSdf.sPtType = "2" Then
                                    Else    'Zones defined
                                        If imNoZones = 1 Then
                                            mAssignCopyTest = 2
                                        Else
                                            tmTzfSrchKey.lCode = tmSdf.lCopyCode
                                            ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                For illoop = 1 To imNoZones Step 1
                                                    ilMatch = False
                                                    For ilIndex = 1 To 6 Step 1
                                                        If tmTzf.lCifZone(ilIndex - 1) > 0 Then
                                                            If StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), Trim$(tmRotNo(illoop).sZone), 1) = 0 Then
                                                                ilMatch = True
                                                                If tmRotNo(illoop).iRotNo > tmTzf.iRotNo(ilIndex - 1) Then
                                                                    mAssignCopyTest = 2
                                                                    'Region test added: 8-4-00
                                                                    '6/8/16: Replaced GoSub
                                                                    'GoSub lRegionTest
                                                                    mRegionTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
                                                                    'Region test added: 8-4-00
                                                                    Exit Function
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilIndex
                                                    If Not ilMatch Then
                                                        mAssignCopyTest = 3
                                                    End If
                                                Next illoop
                                            Else
                                                mAssignCopyTest = 2
                                            End If
                                        End If
                                    End If
                                Else
                                    mAssignCopyTest = 1
                                End If
                                'Region test added: 8-4-00
                                '6/8/16: Replaced GoSub
                                'GoSub lRegionTest
                                mRegionTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
                                'Region test added: 8-4-00
                                Exit Function
                            End If
                            For illoop = 1 To 6 Step 1
                                If Trim$(tmRotNo(illoop).sZone) = "" Then
                                    tmRotNo(illoop).iRotNo = tmCrf.iRotNo
                                    tmRotNo(illoop).sZone = tmCrf.sZone
                                    imNoZones = imNoZones + 1
                                    Exit For
                                End If
                                If StrComp(tmCrf.sZone, tmRotNo(illoop).sZone, 1) = 0 Then
                                    Exit For
                                End If
                            Next illoop
                        End If
                    End If
                End If
            End If
            'ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Next ilLoopOnCrf
        'Loop
    End If
    'Region test added: 8-4-00
    '6/8/16: Replaced GoSub
    'GoSub lRegionTest
    mRegionTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
    'Region test added: 8-4-00
    If (imNoZones = 0) Or (ilVpfIndex = -1) Then
        mAssignCopyTest = 0
    Else
        'Test if all zones specified
        For ilIndex = LBound(tgVpf(ilVpfIndex).sGZone) To LBound(tgVpf(ilVpfIndex).sGZone) Step 1
            If Trim$(tgVpf(ilVpfIndex).sGZone(ilIndex)) <> "" Then
                ilFound = False
                For illoop = 1 To imNoZones Step 1
                    If StrComp(Trim$(tgVpf(ilVpfIndex).sGZone(ilIndex)), Trim$(tmRotNo(illoop).sZone), 1) = 0 Then
                        ilFound = True
                        Exit For
                    End If
                Next illoop
                If Not ilFound Then
                    If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
                        mAssignCopyTest = 3
                    Else
                        mAssignCopyTest = 1
                    End If
                    Exit Function
                End If
            End If
        Next ilIndex
        'Test if superseded
        If (tmSdf.sPtType = "1") Or (tmSdf.sPtType = "2") Or (tmSdf.sPtType = "3") Then
            mAssignCopyTest = 4
            If tmSdf.sPtType = "1" Then
                mAssignCopyTest = 2
            ElseIf tmSdf.sPtType = "2" Then
            Else
                tmTzfSrchKey.lCode = tmSdf.lCopyCode
                ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    For illoop = 1 To imNoZones Step 1
                        ilMatch = False
                        For ilIndex = 1 To 6 Step 1
                            If tmTzf.lCifZone(ilIndex - 1) > 0 Then
                                If StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), Trim$(tmRotNo(illoop).sZone), 1) = 0 Then
                                    ilMatch = True
                                    If tmRotNo(illoop).iRotNo > tmTzf.iRotNo(ilIndex - 1) Then
                                        mAssignCopyTest = 2
                                        Exit Function
                                    End If
                                End If
                            End If
                        Next ilIndex
                        If Not ilMatch Then
                            mAssignCopyTest = 3
                        End If
                    Next illoop
                Else
                    mAssignCopyTest = 2
                End If
            End If
        Else
            mAssignCopyTest = 1
        End If
    End If
    Exit Function
'Region test added: 8-4-00 (Start Point 3)
'lRegionTest:
'
'    If ilRegionMissing Then
'        'Test if any region copy defined
'        tmRafSrchKey1.iAdfCode = tmSdf.iAdfCode
'        ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            '5-1-15 change to support to new key
''            tmCrfSrchKey4.sRotType = slType
''            tmCrfSrchKey4.iEtfCode = 0
''            tmCrfSrchKey4.iEnfCode = 0
''            tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
''            tmCrfSrchKey4.lChfCode = tmSdf.lChfCode         '5-18-05 set the starting point of the search
''            tmCrfSrchKey4.lFsfCode = tmSdf.lFsfCode              'feed code
''            'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
''            tmCrfSrchKey4.iRotNo = 32000
''            ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get last current record to obtain date
''            '5-1-15 remove vehicle test
''            Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode)
'            lgRegionCvfCount = lgRegionCvfCount + lgRegionCvfCount
'            For ilLoopOnCrf = LBound(tmCRFByCntr) To UBound(tmCRFByCntr) - 1
'                tmCrf = tmCRFByCntr(ilLoopOnCrf).tCrf
'            'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode)
'                '5-19-05 place matching test with the dowhile to prevent too many reads
'                'If (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode) Then
'                    'Test date, time, day and zone
'
'                'ilBypassCrf = False
''                If Not gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf) Then     '5-1-15
''                    ilBypassCrf = True
''                End If
'                'If (tmCrf.sState <> "D") And (Not ilBypassCrf) Then
'                    gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
'                    llSAsgnDate = gDateValue(slDate)
'                    gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
'                    llEAsgnDate = gDateValue(slDate)
'                    If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) And (tmCrf.iVefCode = ilCrfVefCode) Then
'                        If Trim$(tmCrf.sZone) = "R" Then
'                            ilRegionMissing = True
'                            'Exit Do
'                            Exit For
'                        End If
'                    End If
'                'End If
'
'                'End If
'               ' ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'               Next ilLoopOnCrf
'           'Loop
'        Else
'            ilRegionMissing = False
'        End If
'    End If
'    '6/8/16: Replaced GoSub
'    'GoSub lRegionSupersedeTest
'    mRegionSupersedeTest ilRegionSuperseded, ilAvailOK, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
'    Return
'lRegionSupersedeTest:
'    If ilRegionSuperseded = -1 Then
'        tmRafSrchKey1.iAdfCode = tmSdf.iAdfCode
'        ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            ilAvailOK = True
'            slType = "A"
'            '5-1-15 change to support to new cvf
''            tmCrfSrchKey4.sRotType = slType
''            tmCrfSrchKey4.iEtfCode = 0
''            tmCrfSrchKey4.iEnfCode = 0
''            tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
''            tmCrfSrchKey4.lChfCode = tmSdf.lChfCode     '5-19-05 set starting point of search
''            tmCrfSrchKey4.lFsfCode = tmSdf.lFsfCode 'feed code
''            'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
''            tmCrfSrchKey4.iRotNo = 32000
''            ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get last current record to obtain date
''            Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmSdf.lFsfCode = tmCrf.lFsfCode)
'            For ilLoopOnCrf = LBound(tmCRFByCntr) To UBound(tmCRFByCntr) - 1
'                tmCrf = tmCRFByCntr(ilLoopOnCrf).tCrf
'                lgSupercedeCount = lgSupercedeCount + 1
'
'            'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lFsfCode = tmCrf.lFsfCode)
'                'Test date, time, day and zone
'                ilBypassCrf = False
''                If Not gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf) Then     '5-1-15
''                    ilBypassCrf = True
''                End If
'                If (tmCrf.sDay(ilDay) = "Y") And (tmSdf.iLen = tmCrf.iLen) And (tmCrf.sState <> "D") And (Not ilBypassCrf) And (tmCrf.iVefCode = ilCrfVefCode) Then      '5-1-15
'                    gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
'                    llSAsgnDate = gDateValue(slDate)
'                    gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
'                    llEAsgnDate = gDateValue(slDate)
'                    If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
'                        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
'                        llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
'                        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
'                        llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
'                        If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
'                            ilAvailOK = True    'Ok to book into
'                            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
'                                Do
'                                    If ilEvtIndex > tmSsf.iCount Then
'                                        imSsfRecLen = Len(tmSsf)
'                                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
'                                        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
'                                            ilEvtIndex = 1
'                                        Else
'                                            ilRegionSuperseded = 4
'                                            Return
'                                        End If
'                                    End If
'                                    'Scan for avail that matches time of spot- then test avail name
'                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
'                                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
'                                        'Test time-
'                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
'                                        llAvailTime = CLng(gTimeToCurrency(slTime, False))
'                                        If llSpotTime = llAvailTime Then
'                                            If tmCrf.sInOut = "I" Then
'                                                If tmCrf.ianfCode <> tmAvail.ianfCode Then
'                                                    ilAvailOK = False   'No
'                                                End If
'                                            Else
'                                                If tmCrf.ianfCode = tmAvail.ianfCode Then
'                                                    ilAvailOK = False   'No
'                                                End If
'                                            End If
'                                            Exit Do
'                                        ElseIf llSpotTime < llAvailTime Then
'                                            ilRegionSuperseded = 4
'                                            Return
'                                        End If
'                                    End If
'                                    ilEvtIndex = ilEvtIndex + 1
'                                Loop
'                            End If
'                            If ilAvailOK Then
'                                If Trim$(tmCrf.sZone) <> "R" Then
'                                    If Trim$(tmCrf.sZone) = "" Then
'                                        tmRsfSrchKey1.lCode = tmSdf.lCode
'                                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        If ilRet = BTRV_ERR_NONE Then
'                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
'                                                ilRegionSuperseded = 2
'                                                If ilRegionRotNo = -1 Then
'                                                    ilRegionRotNo = tmRsf.iRotNo
'                                                End If
'                                            End If
'                                        Else
'                                            ilRegionSuperseded = 0
'                                        End If
'                                        Return
'                                    End If
'                                Else
'                                    If ilRegionRotNo = -1 Then
'                                        ilRegionRotNo = tmCrf.iRotNo
'                                    End If
'                                    tmRsfSrchKey1.lCode = tmSdf.lCode
'                                    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'                                    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
'                                        If tmRsf.lRafCode = tmCrf.lRafCode Then
'                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
'                                                ilRegionSuperseded = 2
'                                                Return
'                                            Else
'                                                ilRegionSuperseded = 4
'                                                Return
'                                            End If
'                                        End If
'                                        ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                                    Loop
'                                    ilRegionSuperseded = 1
'                                    Return
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'                'ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                Next ilLoopOnCrf
'            'Loop
'        End If
'        ilRegionSuperseded = 0
'    End If
'    Return
'    'Region test added: 8-4-00 (End Point 3)
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mErrMsg                         *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Translate Micro Help error code *
'*                                                     *
'*******************************************************
Sub mErrMsg(ilError As Integer)
    Dim slMsg As String
    If (ilError = -99) Or (ilError = 0) Then
        Exit Sub
    End If
    Select Case ilError
        Case -1
            slMsg = "Invalid Job Handle"
        Case -2
            slMsg = "Only One Designer Active Allowed"
        Case -3
            slMsg = "Invalid Project Type"
        Case -4
            slMsg = "Print function called without Print Job Opened"
        Case -5
            slMsg = "LLPrintSetBoxText called without Print Job Opened"
        Case -6
            slMsg = "Multi-Printing Not Allowed"
        Case -10
            slMsg = "LLPrintStart or LLPrintWithBoxStart called without Print Job Opened"
        Case -11
            slMsg = "Print Device could not be opened"
        Case -12
            slMsg = "Error while printing (could be insufficient disk space, paper jam or missing DLL)"
        Case -13
            slMsg = "An application can only have one Job Open"
        Case -14
            slMsg = "Visula Basic DLL missing"
        Case -15
            slMsg = "No Printer available"
        Case -16
            slMsg = "No Preview mode set"
        Case -17
            slMsg = "No Preview files found"
        Case -18
            slMsg = "Parameter call error"
        Case -19
            slMsg = "Expression in LLExprEvaluate could not be interpreted"
        Case -20
            slMsg = "Unknown expression-mode in LLSetOption"
        Case -21
            slMsg = "No table defined with LL_PROJECT_LIST"
        Case -22
            slMsg = "Project file not found"
        Case -23
            slMsg = "Expression error"
        Case -24
            slMsg = "Project file has wrong format"
        Case -25
            slMsg = "LLPrintEnableObject- object name is invalid"
        Case -26
            slMsg = "LLPrintEnableObject- project has no objects"
        Case -27
            slMsg = "LLPrintEnableObject- no object with name defined"
        Case -28
            slMsg = "LLPrint...Start- no table in the table mode"
        Case -29
            slMsg = "LLPrint...Start- project has no objects"
        Case -30
            slMsg = "LLPrintGetTextCharsPrinted- no text object"
        Case -31
            slMsg = "Variable does not exist"
        Case -32
            slMsg = "Field function used although the project is not a table object"
        Case -33
            slMsg = "Expression mode error"
        Case -34
            slMsg = "Error code error error"
        Case -35
            slMsg = "Variable not defined"
        Case -36
            slMsg = "Field not defined"
        Case -37
            slMsg = "Sorting order not defined"
        Case -99
            slMsg = "User Aborted printing"
        Case -100
            slMsg = "DLL required are in error"
        Case -101
            slMsg = "Required Language DLL missing"
        Case -102
            slMsg = "Memory problems"
    End Select
    slMsg = "Report Error #" & Trim$(str$(ilError)) & ": " & slMsg
    MsgBox slMsg, vbOKOnly, "Report Error"
End Sub
Function mGetRegionalFlag() As Integer
'
'
'           mGetRegionalFlag - setup the code from the three types of Regional copy
'                               errors or warnings
'
'           imNonRegionDefined -  (currently not tested)
'                               true : non region copy defined that can be assigned to the spot
'                               false: non-region copy not found (error)
'           imRegionMissing - True: rotation defined for contract other than spot contract that
'                                   is valid for date of spot and vehicle  (warning)
'                              false: region defined for spot or no regions required
'           imRegionSuperceded - 1: not assigned (no regional ever assigned)
'                                2: regional superseded (regional copy superseded)
'                                4: Ok or no regional copy
'            Return: integer with one value
'                   0 = none defined
'                   1 = warning : regional copy not defined for another cnt of same advt
'                   2 = regional copy exists but not assigned
'                   3 = regional copy exists , superseded but not assigned
'                   4 = OK
Dim ilFlag As Integer
    mGetRegionalFlag = 0
    ilFlag = 0
    If (imRegionMissing) And imRegionSuperseded = 0 Then    'another cntr of same advt doesnt have region copy
        ilFlag = 1
    ElseIf imRegionSuperseded = 1 Then      'regional copy not assigned
        ilFlag = 2
    ElseIf imRegionSuperseded = 2 Then      'reginal copy superseded
        ilFlag = 3
    ElseIf imRegionSuperseded = 4 Then      'OK
        ilFlag = 4
    End If
    mGetRegionalFlag = ilFlag
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSeqNo                       *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine sequence # for spot   *
'*                                                     *
'*******************************************************
Function mGetSeqNo(tlSdf As SDF) As Integer
    Dim ilSpotSeqNo As Integer
    Dim ilSpot As Integer
    Dim ilEvtIndex As Integer
    'Dim slSsfType As String
    Dim ilSSFType As Integer
    Dim ilRet As Integer
    ''slSsfType = "O" 'On Air
    '11/24/12
    'ilSSFType = 0 'On Air
    ilSSFType = tmSdf.iGameNo
    ilSpotSeqNo = 0
    If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        'If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefCode <> tlSdf.iVefCode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
        If (tmSsf.iType <> ilSSFType) Or (tmSsf.iVefCode <> tlSdf.iVefCode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
            'tmSsfSrchKey.sType = slSsfType
            tmSsfSrchKey.iType = ilSSFType
            tmSsfSrchKey.iVefCode = tlSdf.iVefCode
            tmSsfSrchKey.iDate(0) = tlSdf.iDate(0)
            tmSsfSrchKey.iDate(1) = tlSdf.iDate(1)
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            imSsfRecLen = Len(tmSsf)
            ilRet = gSSFGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get last current record to obtain date
        Else
            ilRet = BTRV_ERR_NONE
        End If
        ilSpotSeqNo = 0
        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
            ilEvtIndex = 1
            Do
                If ilEvtIndex > tmSsf.iCount Then
                    imSsfRecLen = Len(tmSsf)
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
                    If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
                        ilEvtIndex = 1
                    Else
                        Exit Do
                    End If
                End If
                'Scan for avail that matches time of spot- then test avail name
               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                    'Test time-
                    If (tmAvail.iTime(0) = tlSdf.iTime(0)) And (tmAvail.iTime(1) = tlSdf.iTime(1)) Then
                        For ilSpot = ilEvtIndex + 1 To ilEvtIndex + tmAvail.iNoSpotsThis Step 1
                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot)
                            If tmSpot.lSdfCode = tlSdf.lCode Then
                                Exit Do
                            Else
                                ilSpotSeqNo = ilSpotSeqNo + 1
                            End If
                        Next ilSpot
                        Exit Do
                    End If
                End If
                ilEvtIndex = ilEvtIndex + 1
            Loop
        End If
    End If
    mGetSeqNo = ilSpotSeqNo
    Exit Function
End Function
'*******************************************************
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Sub mObtainCopy(slProduct As String, slZone As String, slCart As String, slISCI As String, slChfFsfProduct As String)
'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           slProduct(O)- Product (different zones separated by Chr(10)
'                         first product obtained from tmChf if time zone
'           slZone(O)- Zones
'           slCart(O)- Carts (different zones separated by Chr(10))
'           slISCI(O)- ISCI (different zones separated by Chr(10))
'           slChfFsfProduct(I) - product desc from contract header or feed spot
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    slProduct = ""
    slZone = ""
    slCart = ""
    slISCI = ""
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.lcpfCode > 0 Then
                tmCpfSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmCpf.sISCI = ""
                    tmCpf.sName = ""
                End If
                slISCI = Trim$(tmCpf.sISCI)
                slProduct = Trim$(tmCpf.sName)
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
                slCart = " "
            End If
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        slProduct = Trim$(slChfFsfProduct)
        slZone = " "   'First row leave blank so product and copy line up
        slCart = " "
        slISCI = " "
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ' Look for the first positive lZone value
            For ilIndex = 1 To 6 Step 1
                If (tmTzf.lCifZone(ilIndex - 1) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), "Oth", 1) <> 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If slZone = "" Then
                            slZone = Trim$(tmTzf.sZone(ilIndex - 1))
                        Else
                            slZone = slZone & Chr$(10) & Trim$(tmTzf.sZone(ilIndex - 1))
                        End If
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmCpf.sISCI = ""
                                tmCpf.sName = ""
                            End If
                            If slISCI = "" Then
                                If Trim$(tmCpf.sISCI) = "" Then
                                    slISCI = " "
                                Else
                                    slISCI = Trim$(tmCpf.sISCI)
                                End If
                            Else
                                If Trim$(tmCpf.sISCI) = "" Then
                                    slISCI = slISCI & Chr$(10) & " "
                                Else
                                    slISCI = slISCI & Chr$(10) & Trim$(tmCpf.sISCI)
                                End If
                            End If
                            If slProduct = "" Then
                                slProduct = Trim$(tmCpf.sName)
                            Else
                                slProduct = slProduct & Chr$(10) & "               " & Trim$(tmCpf.sName)
                            End If
                        End If
                        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                            If tmCif.iMcfCode <> tmMcf.iCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmMcf.sName = ""
                                End If
                                If slCart = "" Then
                                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                Else
                                    slCart = slCart & Chr$(10) & Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                End If
                            Else
                                If slCart = "" Then
                                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                Else
                                    slCart = slCart & Chr$(10) & Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                End If
                            End If
                        Else
                            If slCart = "" Then
                                slCart = " "
                            Else
                                slCart = slCart & Chr$(10) & " "
                            End If
                        End If
                    End If
                End If
            Next ilIndex
            For ilIndex = 1 To 6 Step 1
                If (tmTzf.lCifZone(ilIndex - 1) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex - 1)), "Oth", 1) = 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    If slZone = "" Then
                        slZone = Trim$(tmTzf.sZone(ilIndex - 1))
                    Else
                        slZone = slZone & Chr$(10) & "Other"
                    End If
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmCif.lcpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lcpfCode
                            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmCpf.sISCI = ""
                                tmCpf.sName = ""
                            End If
                            If slISCI = "" Then
                                If Trim$(tmCpf.sISCI) = "" Then
                                    slISCI = " "
                                Else
                                    slISCI = Trim$(tmCpf.sISCI)
                                End If
                            Else
                                If Trim$(tmCpf.sISCI) = "" Then
                                    slISCI = slISCI & Chr$(10) & " "
                                Else
                                    slISCI = slISCI & Chr$(10) & Trim$(tmCpf.sISCI)
                                End If
                            End If
                            If slProduct = "" Then
                                slProduct = Trim$(tmCpf.sName)
                            Else
                                slProduct = slProduct & Chr$(10) & "               " & Trim$(tmCpf.sName)
                            End If
                        End If
                        If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                            If tmCif.iMcfCode <> tmMcf.iCode Then
                                tmMcfSrchKey.iCode = tmCif.iMcfCode
                                ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmMcf.sName = ""
                                End If
                                If slCart = "" Then
                                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                Else
                                    slCart = slCart & Chr$(10) & Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                End If
                            Else
                                If slCart = "" Then
                                    slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                Else
                                    slCart = slCart & Chr$(10) & Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                End If
                            End If
                        Else
                            If slCart = "" Then
                                slCart = " "
                            Else
                                slCart = slCart & Chr$(10) & " "
                            End If
                        End If
                    End If
                End If
            Next ilIndex
        End If
    End If
End Sub
'********************************************************************
'*
'*      Procedure Name:mObtainCopyCntr
'*
'*             Created:10/09/93      By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:Obtain the Sdf records to be
'*                     reported
'*          9/25/98 - Test spot & pkg vehicles for
'*                    rotations
'*          8/23/99 Show contracts missing copy for
'                   airing vehicles
'           8/26/99 Include/Exclude fill spots
'*      dh - 8-18-00 Add regional copy feature
'*      dh 3-16-07 fix subscript out of range when flagging Missed,
'*      dh 6-22-07 For line option, show the start date as the requested start date or the
'          scheduline line start date, which ever is later.
'          i.e. if requested date is Fri-Sun (for a daily buy), user wants to see
'               the line start date as the Fri (or later), not the Monday of thatweek
'               or the true start date of the line
'       DH 7-10-08 show first day that copy is missing when line option
'***********************************************************************
Sub mObtainCopyCntr(ilVefCode As Integer, slName As String, ilVpfIndex As Integer, slStartDate As String, slEndDate As String, ilCntrSpots As Integer, ilFeedSpots As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilUpper As Integer          'Upper bound for tmCopyCntr, do not use for other purposes
    Dim ilFound As Integer
                                    '1 spot in M, U, or R column is missed
    Dim illoop As Integer
    Dim ilLoopAir As Integer
    Dim ilAiringVeh As Integer      'true if airing vehicle that is tested, otherwise false
    Dim ilAssign As Integer 'True=Spot can be assigned; False=No Copy
    Dim ilAsgnVefCode As Integer
    Dim ilPkgRot As Integer
    Dim ilPkgVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim ilRotNo As Integer
    Dim ilPkgRotNo As Integer
    Dim ilLnVefCode As Integer
    Dim ilLnRotNo As Integer
    Dim ilLnRet As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim ilRegionalFlag As Integer       '8-11-00
    Dim llAirDate As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim tlVef As VEF
    Dim slCntrNo As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE      'type long field
    '8-11-00
    Dim ilPkgNonRegionDefined As Integer
    Dim ilPkgRegionMissing As Integer
    Dim ilPkgRegionSuperseded As Integer
    Dim ilLnNonRegionDefined As Integer
    Dim ilLnRegionMissing As Integer
    Dim ilLnRegionSuperseded As Integer
    Dim ilRegionRotNo As Integer
    Dim ilPkgRegionRotNo As Integer
    Dim ilLnRegionRotNo As Integer
    Dim ilIncludeUnAssg As Integer           '5-20-05
    Dim ilIncludeReassg As Integer
    Dim slStr As String
    Dim ilSaveRdfCode As Integer
    Dim ilGameNo As Integer
    Dim ilIncludeLine As Integer            'include each line separately rather than combining all lines for the contract
    Dim ilLineNo As Integer
    Dim ilLoopDays As Integer
    Dim llDate As Long
    Dim ilMonRequest(0 To 1) As Integer
    Dim ilDay As Integer
    Dim llStartOfFlt As Long
    Dim ilAirToSellList() As Integer        '7-16-14
    Dim ilTZAdj As Integer
    Dim llTZAdjDate As Long
    Dim slTZAdjDate As String
    Dim ilLoopOnZone As Integer
    Dim llAirTime As Long
    Dim llSellTimesToTest() As Long  'list of selling times to look for adjusted zone times
    Dim llTemp As Long
    Dim llLatestSellTime As Long        'latest selling time to test for cross midnite and time zones
    Dim ilUpperTemp As Integer
    Dim llTempAirTime As Long
    Dim blFoundSell As Boolean
    Dim ilInx As Integer
    Dim slType As String * 1
    Dim slDate As String


    ilIncludeUnAssg = gSetCheck(RptSel!ckcSelC3(0).Value)   '5-20-05
    ilIncludeReassg = gSetCheck(RptSel!ckcSelC3(1).Value)   '5-20-05
    ilIncludeLine = False
    If RptSel!rbcSelCInclude(1).Value Then      'show by line (vs contract)
        ilIncludeLine = True
    End If

    '6-22-07 no longer backup date to Monday to determine show the valid start date day for the sched line
    'i.e. 6-22-07 if requested date is Fri-Sun (for a daily buy), user wants to see
    'the line start date as the Fri (or later), not the Monday of thatweek
    'or the true start date of the line
    llDate = gDateValue(slStartDate)
    'ilDay = gWeekDayLong(llDate)
    'Do While ilDay <> 0          'loop while day is not a monday
    '    llDate = llDate - 1
    '    ilDay = gWeekDayLong(llDate)
    'Loop
    gPackDateLong llDate, ilMonRequest(0), ilMonRequest(1)

    lmTodayDate = gDateValue(gNow())        'used in mAssignCopyTest
    
    ilTZAdj = 0
    ilUpper = UBound(tmCopyCntr)
    'Retrieve the vehicle
    tmVefSrchKey.iCode = ilVefCode      'retrieve vehicle to determine if airing or not
    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If tlVef.sType <> "A" Then              'anything other than airing
        '7-16-14 need to determine if the next day needs to be examined becuase of time zones and getting spots for earlier time zones (i.e. pst sun 9p-12m needs to get the spots from Mon 12m-3a)
        ReDim ilAirToSellList(0 To 0) As Integer
        If tlVef.sType <> "S" Then          'not selling, either conventional or game
            ilAirToSellList(0) = tlVef.iCode
            ReDim Preserve ilAirToSellList(0 To 1) As Integer
        Else
            gBuildLinkArray hmVLF, tlVef, slStartDate, ilAirToSellList()
        End If
        For illoop = LBound(ilAirToSellList) To UBound(ilAirToSellList) - 1        'loop thru the airing vehicles that are associated with the selling vehicle selected
            'see if any zones that need to backup time (negative offset) defined for airing vehicle.  If not, no need to process it
            ilVpfIndex = gBinarySearchVpf(ilAirToSellList(illoop))
            If ilVpfIndex >= 0 Then         'no vpf for this vehicle, ignore the airing vehicle
                'find the most negative time zone adjustment
                For ilLoopOnZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                    If tgVpf(ilVpfIndex).iGLocalAdj(ilLoopOnZone) < ilTZAdj Then
                        ilTZAdj = tgVpf(ilVpfIndex).iGLocalAdj(ilLoopOnZone)
                        
                    End If
                Next ilLoopOnZone
            End If
        Next illoop
        
        ReDim llSellTimesToTest(0 To 0) As Long
        
        llTZAdjDate = gDateValue(slEndDate)
        slTZAdjDate = slEndDate
        If ilTZAdj < 0 Then
            llTZAdjDate = llTZAdjDate + 1
            slTZAdjDate = Format(llTZAdjDate, "m/d/yy")
'slStartDate = slTZAdjDate   'debug to test the partial cross over date
            'at least 1 airing vehicle has time zone; get the airing vehicle links if selling
            If tlVef.sType = "S" Then
                'get airing links for last date +1
                ilUpperTemp = 0
                llLatestSellTime = 0
                'ReDim tmVlf(1 To 1) As VLF
                ReDim tmVlf(0 To 0) As VLF
                gObtainVlf "S", hmVLF, ilVefCode, llTZAdjDate, tmVlf()        'get the associated airing vehicles links with this selling  vehicle
                For ilLoopOnZone = LBound(tmVlf) To UBound(tmVlf) - 1
                    gUnpackTimeLong tmVlf(ilLoopOnZone).iAirTime(0), tmVlf(ilLoopOnZone).iAirTime(1), False, llAirTime
                    'find airing times less than 3am, then get the selling time to test
                    If llAirTime < -(ilTZAdj * 3600) Then
                        ilInx = gBinarySearchVpf(tmVlf(ilLoopOnZone).iAirCode)
                        'If tgVpf(ilInx).sCopyOnAir = "Y" Then
                            gUnpackTimeLong tmVlf(ilLoopOnZone).iSellTime(0), tmVlf(ilLoopOnZone).iSellTime(1), False, llTemp
                            llSellTimesToTest(ilUpperTemp) = llTemp
                            ilUpperTemp = ilUpperTemp + 1
                            ReDim Preserve llSellTimesToTest(LBound(llSellTimesToTest) To ilUpperTemp) As Long
                            If llLatestSellTime < llTemp Then
                                llLatestSellTime = llTemp
                            End If
                        'End If
                    End If
                Next ilLoopOnZone
            Else                            'conventional or game
                llLatestSellTime = -(ilTZAdj * 3600)
            End If
        Else
            llLatestSellTime = 86400             '12m, end of day
        End If
        
        
        btrExtClear hmSdf   'Clear any previous extend operation
        ilExtLen = Len(tmSdf)
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
        btrExtClear hmSdf   'Clear any previous extend operation
        tmSdfSrchKey1.iVefCode = ilVefCode
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


            '8-2-04  Exclude/include contract/feed spots
            tlLongTypeBuff.lCode = 0
            If Not ilCntrSpots Or Not ilFeedSpots Then           'either local or feed spots are to be excluded
                If ilCntrSpots Then                         'include local only
                    ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                    If lmCntrCode = 0 Then          '11-16-05 option for single contract
                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
                    Else        'extended reads for single contract #
                        tlLongTypeBuff.lCode = lmCntrCode
                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
                    End If
                ElseIf ilFeedSpots Then                      'include feed only
                    ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
                End If
            Else                'include everything, check for selective contract
                If lmCntrCode <> 0 Then
                    tlLongTypeBuff.lCode = lmCntrCode
                    ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
                End If
            End If



            tlIntTypeBuff.iType = ilVefCode
            ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            If slStartDate <> "" Then
                gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("Sdf", "SdfDate")
                'If slEndDate <> "" Then
                If slTZAdjDate <> "" Then       '7-16-14
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                Else
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
            End If
            'If slEndDate <> "" Then
            If slTZAdjDate <> "" Then       '7-16-14
                'gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                gPackDate slTZAdjDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
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
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llAirTime            '7-14-14
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llAirDate            '7-14-14
                    '10-11-18 Ignore hidden and cancelled spots
                    If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "M" And RptSel!ckcSelC3(2).Value = vbChecked)) And (llAirDate <> llTZAdjDate) Or ((llAirTime <= llLatestSellTime And llAirDate = llTZAdjDate)) And ((tmSdf.sSchStatus <> "C") Or (tmSdf.sSchStatus = "H")) Then
                        '8/23/99 show contracts missing copy for airing vehicles
                        '8/26/99 option to include fill spots
                        If (tmSdf.sSpotType <> "X") Or (tmSdf.sSpotType = "X" And RptSel!ckcSelC5(0).Value = vbChecked) Then    'include if not a fill or its a fill that user wants included
                            If llAirDate = llTZAdjDate And ilTZAdj < 0 And tlVef.sType = "S" Then
                                'any links that need to be tested for requested end date +1, crossing mid
                                For ilLoopOnZone = LBound(llSellTimesToTest) To UBound(llSellTimesToTest) - 1
                                    If llAirTime = llSellTimesToTest(ilLoopOnZone) Then
                                        '6/8/16: Replaced GoSub
                                        'GoSub CopyCntr:
                                        mCopyCntr llAirDate, llAirTime, llLatestSellTime, llDate, slStartDate, slEndDate, slLive, ilRdfCode, ilVpfIndex, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, ilAssign, ilUpper, ilIncludeLine, slName, ilMonRequest(), ilAiringVeh
                                        Exit For
                                    End If
                                Next ilLoopOnZone
                                ilRet = ilRet
                            Else
                                ilAiringVeh = False         'need to know in copycntr subroutine if this is an airing vehicle
                                '6/8/16: Replaced GoSub
                                'GoSub CopyCntr:
                                mCopyCntr llAirDate, llAirTime, llLatestSellTime, llDate, slStartDate, slEndDate, slLive, ilRdfCode, ilVpfIndex, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, ilAssign, ilUpper, ilIncludeLine, slName, ilMonRequest(), ilAiringVeh
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
    Else                'airing vehicles only
        llTZAdjDate = gDateValue(slEndDate)
        slTZAdjDate = slEndDate
        ilTZAdj = 0
        ilVpfIndex = gBinarySearchVpf(ilVefCode)
        If ilVpfIndex >= 0 Then         'no vpf for this vehicle, ignore the airing vehicle
            'find the most negative time zone adjustment
            For ilLoopOnZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                If tgVpf(ilVpfIndex).iGLocalAdj(ilLoopOnZone) < ilTZAdj Then
                    ilTZAdj = tgVpf(ilVpfIndex).iGLocalAdj(ilLoopOnZone)
                End If
            Next ilLoopOnZone
        End If
        llLatestSellTime = 86400        '12m
        If ilTZAdj < 0 Then
            llTZAdjDate = llTZAdjDate + 1
            slTZAdjDate = Format(llTZAdjDate, "m/d/yy")
            llLatestSellTime = -(ilTZAdj * 3600)            'latest time to test on extra day processing when zones exist
        End If

        'For llAirDate = gDateValue(slStartDate) To gDateValue(slEndDate) Step 1   'loop on dates for each selling vehicle for the selected airing veh
'slStartDate = slTZAdjDate      'debug to test the partial cross over date
         For llAirDate = gDateValue(slStartDate) To gDateValue(slTZAdjDate) Step 1  'loop on dates air vehicle
 
            'ReDim tmVlf(1 To 1) As VLF
            ReDim tmVlf(0 To 0) As VLF
            gPackDateLong llAirDate, ilDate0, ilDate1
            gObtainVlf "A", hmVLF, ilVefCode, llAirDate, tmVlf()        'get the associated selling vehicles with this airing vehicle
            For ilLoopAir = LBound(tmVlf) To UBound(tmVlf) - 1 Step 1
                 blFoundSell = False
'this commented out code tests for the selling vehicle:  if selected, dont process the this airing vehicle as it would already have been processed with selling
'                 For ilLoop = LBound(imSellVefSelected) To UBound(imSellVefSelected) - 1
'                    If tmVlf(ilLoopAir).iSellCode = imSellVefSelected(ilLoop) Then
'                        blFoundSell = True
'                        Exit For
'                    End If
'                Next ilLoop
                If Not blFoundSell Then                                 'selling selected, dont process twice
                    ilVpfIndex = -1                                         'find the selling vehicles associated options tabel
                    'For ilAssign = 0 To UBound(tgVpf) Step 1
                    '    If tmVlf(ilLoopAir).iSellCode = tgVpf(ilAssign).iVefKCode Then
                        ilAssign = gBinarySearchVpf(tmVlf(ilLoopAir).iSellCode)
                        If ilAssign <> -1 Then
                            ilVpfIndex = ilAssign
                    '        Exit For
                        End If
                    'Next ilAssign
                   
    
                    If ilVpfIndex <> -1 Then                                'if neg, veh options not found
                        tmSdfSrchKey1.iVefCode = tmVlf(ilLoopAir).iSellCode 'setup for first spot access
                        tmSdfSrchKey1.iDate(0) = ilDate0
                        tmSdfSrchKey1.iDate(1) = ilDate1
                        tmSdfSrchKey1.iTime(0) = tmVlf(ilLoopAir).iSellTime(0)
                        tmSdfSrchKey1.iTime(1) = tmVlf(ilLoopAir).iSellTime(1)
                        tmSdfSrchKey1.sSchStatus = "G"
                        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf(ilLoopAir).iSellCode = tmSdf.iVefCode) And (ilDate0 = tmSdf.iDate(0)) And (ilDate1 = tmSdf.iDate(1)) And (tmVlf(ilLoopAir).iSellTime(0) = tmSdf.iTime(0)) And (tmVlf(ilLoopAir).iSellTime(1) = tmSdf.iTime(1))
                            If (tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O" Or tmSdf.sSchStatus = "S") Then
                                '8/26/99 option to include fill spots
                                '8-2-04 Test inclusion/exclusion of contract/feed spots
                                If (tmSdf.sSpotType <> "X") Or (tmSdf.sSpotType = "X" And RptSel!ckcSelC5(0).Value = vbChecked) And ((tmSdf.lChfCode = 0 And ilCntrSpots) Or (tmSdf.lFsfCode = 0 And ilFeedSpots)) Then   'include if not a fill or its a fill that user wants included
                                    If (tmSdf.lChfCode = lmCntrCode) Or lmCntrCode = 0 Then     '2-2-10 test for selective cntr on airing vehicles
                                        ilAiringVeh = True
                                        gUnpackTimeLong tmVlf(ilLoopAir).iAirTime(0), tmVlf(ilLoopAir).iAirTime(1), False, llAirTime            '7-16-14                                    GoSub CopyCntr:                             'do the copy testing
                                        '6/8/16: Replaced GoSub
                                        'GoSub CopyCntr:
                                        mCopyCntr llAirDate, llAirTime, llLatestSellTime, llDate, slStartDate, slEndDate, slLive, ilRdfCode, ilVpfIndex, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, ilAssign, ilUpper, ilIncludeLine, slName, ilMonRequest(), ilAiringVeh
                                    End If
                                End If
                            End If
                            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                End If
            Next ilLoopAir
        Next llAirDate
    End If
    Exit Sub
'CopyCntr:               'all types of vehicles come thru here to check if copy exists
'    '7-16-14 test for 13m-3a selling
'     If llAirDate = gDateValue(slEndDate) + 1 Then          'processing day+1 because of time zones
'        ' llTempAirTime = llAirTime + (ilTZAdj * 3600)
'        If llAirTime > llLatestSellTime Then
'         'If llTempAirTime >= 0 Then
'             Return
'         End If
'    End If
'
'    ilGameNo = tmSdf.iGameNo
'    'ilAssign = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode)
'    ilSchPkgVefCode = 0
'    ilAsgnVefCode = 0
'    ilLnVefCode = 0
'    ilPkgVefCode = 0
'    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilAsgnVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
'    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSpotType = "X") Then
'        slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, hmSmf, hmCrf)
'        If (slStr = "S") Or (slStr = "B") Then
'            ilSchPkgVefCode = gGetMGPkgVefCode(hmClf, tmSdf)
'        End If
'        If slStr = "O" Then
'            ilAsgnVefCode = ilLnVefCode
'            ilLnVefCode = 0
'        ElseIf slStr = "S" Then
'            ilPkgVefCode = ilSchPkgVefCode
'            ilSchPkgVefCode = 0
'            ilLnVefCode = 0
'        Else
'            If ilPkgVefCode = ilSchPkgVefCode Then
'                ilSchPkgVefCode = 0
'            End If
'        End If
'    Else
'        ilLnVefCode = 0
'    End If
'
'    '6-3-15 see if copy defined for billboards
'    If tmSdf.sSpotType = "C" Then
'        slType = "C"
'    ElseIf tmSdf.sSpotType = "O" Then
'        slType = "O"
'    Else
'        slType = "A"
'    End If
'    '6-3-15
'    'build array of Rotation headers based package vehicle, line vehicle, sched package vehicle so that the crf & cvf do not have to be constantly reread
'    'mBuildCRFByCntr ilCrfVefCode, slSpotDate, slType, slLive
'    slDate = Format$(llAirDate, "m/d/yy")
'    mBuildCRFByCntr ilSchPkgVefCode, ilAsgnVefCode, ilLnVefCode, ilPkgVefCode, slDate, slType, slLive
'
'    'ilRet = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo)
'    ilRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
'    ilAssign = ilRet
'    If ilPkgVefCode > 0 Then
'        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
'        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
'        If ilPkgRegionRotNo > ilRegionRotNo Then
'            '8-17-00
'            imNonRegionDefined = ilPkgNonRegionDefined
'            imRegionMissing = ilPkgRegionMissing
'            imRegionSuperseded = ilPkgRegionSuperseded
'        End If
'
'        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
'            If ilPkgRotNo > ilRotNo Then
'                ilRotNo = ilPkgRotNo
'                ilAsgnVefCode = ilPkgVefCode
'                ilAssign = ilPkgRot
'            End If
'        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
'            ilRotNo = ilPkgRotNo
'            ilAsgnVefCode = ilPkgVefCode
'            ilAssign = ilPkgRot
'        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
'            ilRotNo = ilPkgRotNo
'            ilAsgnVefCode = ilPkgVefCode
'            ilAssign = ilPkgRot
'        End If
'    End If
'    If ilSchPkgVefCode > 0 Then
'        ilPkgVefCode = ilSchPkgVefCode
'        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
'        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
'        If ilPkgRegionRotNo > ilRegionRotNo Then
'            '8-17-00
'            imNonRegionDefined = ilPkgNonRegionDefined
'            imRegionMissing = ilPkgRegionMissing
'            imRegionSuperseded = ilPkgRegionSuperseded
'        End If
'
'        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
'            If ilPkgRotNo > ilRotNo Then
'                ilRotNo = ilPkgRotNo
'                ilAsgnVefCode = ilPkgVefCode
'                ilAssign = ilPkgRot
'            End If
'        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
'            ilRotNo = ilPkgRotNo
'            ilAsgnVefCode = ilPkgVefCode
'            ilAssign = ilPkgRot
'        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
'            ilRotNo = ilPkgRotNo
'            ilAsgnVefCode = ilPkgVefCode
'            ilAssign = ilPkgRot
'        End If
'    End If
'
'    If (ilAsgnVefCode <> ilLnVefCode) And (ilLnVefCode > 0) Then
'        ilLnRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilLnVefCode, ilLnRotNo, ilLnNonRegionDefined, ilLnRegionMissing, ilLnRegionSuperseded, ilLnRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
'        If ilLnRegionRotNo > ilRegionRotNo Then
'            '8-17-00
'            imNonRegionDefined = ilLnNonRegionDefined
'            imRegionMissing = ilLnRegionMissing
'            imRegionSuperseded = ilLnRegionSuperseded
'        End If
'        If (ilAssign <> 0) And (ilLnRet <> 0) Then
'            If ilLnRotNo > ilRotNo Then
'                ilRotNo = ilLnRotNo
'                ilAssign = ilLnRet
'                ilAsgnVefCode = ilLnVefCode
'            End If
'        ElseIf (ilAssign = 0) And (ilLnRet <> 0) Then
'            ilRotNo = ilLnRotNo
'            ilAssign = ilLnRet
'            ilAsgnVefCode = ilLnVefCode
'        End If
'    End If
'
'
'    ilRegionalFlag = mGetRegionalFlag()     '8-18-00 Convert the different regional copy warnings/errors into 1 flag
'
'    '12-15-04 change to test for ilRegionalFlag <> 0 (instead of 4) for regional copy OK
'    If ilAssign <> 4 Or ilRegionalFlag <> 0 Then    '8-18-00 4 = copy OK
'        ilFound = False
'        If RptSel!ckcTrans.Value = vbUnchecked Then 'prevent lines from being separated if there are different dayparts for same vehicle
'                                                    'and user doesnt want the dp to print
'            ilSaveRdfCode = 0
'        Else
'            ilSaveRdfCode = ilRdfCode
'        End If
'        For ilLoop = 0 To ilUpper Step 1
'
'            If ilIncludeLine Then
'                ilLineNo = tmSdf.iLineNo
'            Else
'                ilLineNo = 0
'            End If
'            If (tmCopyCntr(ilLoop).lChfCode = tmSdf.lChfCode) And (tmCopyCntr(ilLoop).iVefCode = tmSdf.iVefCode) And (tmCopyCntr(ilLoop).iLen = tmSdf.iLen) And (tmCopyCntr(ilLoop).iAsgnVefCode = ilAsgnVefCode) And (tmCopyCntr(ilLoop).lFsfCode = tmSdf.lFsfCode) And (tmCopyCntr(ilLoop).iRdfcode = ilSaveRdfCode) And (tmCopyCntr(ilLoop).iLineNo = ilLineNo) Then
'                ilFound = True
'
'
'                If (ilAssign = 0) Or (ilAssign = 3) Then 'Copy not defined
'                    tmCopyCntr(ilLoop).iNoSpots = tmCopyCntr(ilLoop).iNoSpots + 1
'                    '7-10-08 if by line, show the first date that a spot has missing copy
'                    If tmCopyCntr(ilLoop).iNoSpots = 1 And ilIncludeLine Then         '1st spot without copy
'                        tmCopyCntr(ilLoop).iStartDate(0) = tmSdf.iDate(0)
'                        tmCopyCntr(ilLoop).iStartDate(1) = tmSdf.iDate(1)
'                    End If
'                    If tmSdf.sSchStatus = "M" Then
'                        tmCopyCntr(ilLoop).iNoSpotsMiss = 1             'at least 1 spot missed missing copy
'                    End If
'                    '11-16-05 show the live flag for the missing copy only
'                    'if array is blank, nothing has been set yet
'                    'l = live, r = recorded, m = mixture of live/recorded
'                    If tmCopyCntr(ilLoop).sLiveFlag <> "X" Then         'if the flag is already mixed, dont touch it
'                        If tmCopyCntr(ilLoop).sLiveFlag <> slLive Then      'must have both live and recorded across the lines
'                            tmCopyCntr(ilLoop).sLiveFlag = "X"
'                        End If
'                    End If
'                ElseIf (ilAssign = 1) Then    'Not assigned
'                    tmCopyCntr(ilLoop).iNoUnAssg = tmCopyCntr(ilLoop).iNoUnAssg + 1
'                    If tmSdf.sSchStatus = "M" Then
'                        tmCopyCntr(ilLoop).iNoUnAssgMiss = 1            'atleast 1 spot missed unassigned
'                    End If
'                Else    'Supersede
'                    tmCopyCntr(ilLoop).iNoToReassg = tmCopyCntr(ilLoop).iNoToReassg + 1
'                    If tmSdf.sSchStatus = "M" Then
'                        tmCopyCntr(ilLoop).iNoToReassgMiss = 1          'at least 1 spot missed to reassign
'                    End If
'                End If
'
'                '8-18-00 Accum errors for regional copy
'                If ilRegionalFlag = 1 Then      'not assigned
'                    tmCopyCntr(ilLoop).iRegionNoUnAssg = tmCopyCntr(ilLoop).iRegionNoUnAssg + 1
'                    If tmSdf.sSchStatus = "M" Then
'                        tmCopyCntr(ilLoop).iRegionNoUnAssgMiss = 1            'atleast 1 spot missed unassigned
'                    End If
'                ElseIf ilRegionalFlag = 2 Then      'superseded
'                    tmCopyCntr(ilLoop).iRegionNoToReassg = tmCopyCntr(ilLoop).iRegionNoToReassg + 1
'                    If tmSdf.sSchStatus = "M" Then
'                        tmCopyCntr(ilLoop).iRegionNoToReassgMiss = 1          'at least 1 spot missed to reassign
'                    End If
'                End If
'
'                Exit For
'            End If
'        Next ilLoop
'        If Not ilFound Then
'            If tmSdf.lChfCode = 0 Then
'                tmFSFSrchKey.lCode = tmSdf.lFsfCode
'                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'                slCntrNo = Trim$(tmFsf.sRefID)
'                Do While Len(slCntrNo) < 8
'                    slCntrNo = "0" & slCntrNo
'                Loop
'            Else
'                If tmChf.lCode <> tmSdf.lChfCode Then
'                    tmChfSrchKey.lCode = tmSdf.lChfCode
'                    ilRet = btrGetEqual(hmCHF, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'                End If
'                slCntrNo = Trim$(str$(tmChf.lCntrNo))
'                Do While Len(slCntrNo) < 8
'                    slCntrNo = "0" & slCntrNo
'                Loop
'            End If
'
'
'            If tmAdf.iCode <> tmSdf.iAdfCode Then
'                tmAdfSrchKey.iCode = tmSdf.iAdfCode
'                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'            End If
'            tmCopyCntr(ilUpper).sKey = tmAdf.sName & slCntrNo & slName & Trim$(str$(tmSdf.iLen))
'
'            If (ilAssign = 0) Or (ilAssign = 3) Then    'Copy not defined
'                tmCopyCntr(ilUpper).iNoSpots = 1
'                '7-10-08 if by line, show the first date that a spot has missing copy
'                If ilIncludeLine Then         '1st spot without copy
'                    tmCopyCntr(ilUpper).iStartDate(0) = tmSdf.iDate(0)
'                    tmCopyCntr(ilUpper).iStartDate(1) = tmSdf.iDate(1)
'                End If
'
'                tmCopyCntr(ilUpper).iNoUnAssg = 0
'                tmCopyCntr(ilUpper).iNoToReassg = 0
'                If tmSdf.sSchStatus = "M" Then
'                    tmCopyCntr(ilUpper).iNoSpotsMiss = 1        '3-16-07 (fix subscript out of range)
'                End If
'            ElseIf (ilAssign = 1) Then    'Not assigned
'                tmCopyCntr(ilUpper).iNoSpots = 0
'                tmCopyCntr(ilUpper).iNoUnAssg = 1
'                tmCopyCntr(ilUpper).iNoToReassg = 0
'                If tmSdf.sSchStatus = "M" Then
'                    'tmCopyCntr(ilLoop).iNoUnAssgMiss = 1
'                    tmCopyCntr(ilUpper).iNoUnAssgMiss = 1       '3-16-07 (fix subscript out of range)
'                End If
'            Else    'Supersede
'                tmCopyCntr(ilUpper).iNoSpots = 0
'                tmCopyCntr(ilUpper).iNoUnAssg = 0
'                tmCopyCntr(ilUpper).iNoToReassg = 1
'                If tmSdf.sSchStatus = "M" Then
'                    'tmCopyCntr(ilLoop).iNoToReassgMiss = 1
'                    tmCopyCntr(ilUpper).iNoToReassgMiss = 1     '3-16-07 (fix subscript out of range)
'                End If
'            End If
'            '8-18-00 Accum errors for regional copy
'            If ilRegionalFlag = 1 Then      'not assigned
'                tmCopyCntr(ilUpper).iRegionNoUnAssg = 1
'                tmCopyCntr(ilUpper).iRegionNoToReassg = 0
'                If tmSdf.sSchStatus = "M" Then
'                    'tmCopyCntr(ilLoop).iRegionNoUnAssgMiss = 1
'                    tmCopyCntr(ilUpper).iRegionNoUnAssgMiss = 1     '3-16-07 (fix subscript out of range)
'                End If
'            ElseIf ilRegionalFlag = 2 Then      'superseded
'                tmCopyCntr(ilUpper).iRegionNoUnAssg = 0
'                tmCopyCntr(ilUpper).iRegionNoToReassg = 1
'                If tmSdf.sSchStatus = "M" Then
'                    'tmCopyCntr(ilLoop).iRegionNoToReassgMiss = 1
'                    tmCopyCntr(ilUpper).iRegionNoToReassgMiss = 1       '3-16-07 (fix subscript out of range)
'                End If
'            End If
'            tmCopyCntr(ilUpper).lChfCode = tmSdf.lChfCode
'            tmCopyCntr(ilUpper).lFsfCode = tmSdf.lFsfCode
'            tmCopyCntr(ilUpper).iVefCode = tmSdf.iVefCode
'            tmCopyCntr(ilUpper).iAsgnVefCode = ilAsgnVefCode
'            tmCopyCntr(ilUpper).iRdfcode = ilSaveRdfCode                '1-05-06
'            tmCopyCntr(ilUpper).iLineNo = 0
'            If ilIncludeLine Then
'                tmCopyCntr(ilUpper).iLineNo = ilLineNo                 '2-28-07
'                tmClfSrchKey.lChfCode = tmSdf.lChfCode
'                tmClfSrchKey.iLine = tmSdf.iLineNo
'                tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
'                tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
'                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
'                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
'                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                Loop
'
'                '***** 7-10-08 First date of spot missing copy has been determine by the
'                'by the missing copy counts: 1st time, save the date
'                'If not showing missing copy, default to the start date of the flight or requested report start date, whichever is later
'
'                If tmCopyCntr(ilUpper).iStartDate(0) = 0 And tmCopyCntr(ilUpper).iStartDate(1) = 0 Then
'                    ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)
'
'                    gPackDate slStartDate, ilDate0, ilDate1
'                    'default to start date of requested report
'                    tmCopyCntr(ilUpper).iStartDate(0) = ilDate0
'                    tmCopyCntr(ilUpper).iStartDate(1) = ilDate1
'
'                    gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llStartOfFlt
'                    'use the start of flight or start date of requested report, whichever is later since a line cant start earlier than requested date
'                    gUnpackDateLong ilMonRequest(0), ilMonRequest(1), llDate
'
'                    If llStartOfFlt < llDate Then
'                        llStartOfFlt = llDate
'                    End If
'
'                    If (tmCff.lChfCode = tmClf.lChfCode) And (tmCff.iClfLine = tmClf.iLine) Then
'                        'ilDay = gWeekDayLong(llDate)
'                        ilDay = gWeekDayLong(llStartOfFlt)
'                        'start looping with the requested start date
'                        For ilLoopDays = 0 To 7
'                            If tmCff.iDay(ilDay) <> 0 Then      'valid first day of week
'                                'gPackDateLong llStartOfFlt, tmCopyCntr(ilUpper).iStartDate(0), tmCopyCntr(ilUpper).iStartDate(1)
'                                Exit For
'                            Else
'                                'llDate = llDate + 1
'                                llStartOfFlt = llStartOfFlt + 1
'                                'ilDay = gWeekDayLong(llDate)
'                                ilDay = gWeekDayLong(llStartOfFlt)
'                            End If
'                        Next ilLoopDays
'                    End If
'                End If
'
'                'OK to use line end date
'                tmCopyCntr(ilUpper).iEndDate(0) = tmClf.iEndDate(0)
'                tmCopyCntr(ilUpper).iEndDate(1) = tmClf.iEndDate(1)
'            End If
'
'
'            If tmSdf.iVefCode <> ilAsgnVefCode Or ilAiringVeh Then  'if the rot and airing vehicle arent the same, get the rot vehicle.
'                'When testing airing vehicles, the spot (airing veh) will be the same as the selling
'                'vehicle since its going back to do spots based on the selling vehicle.  The ilAiringVeh flag
'                'to determine if airing veh to setup correct rotation & airing vehicles on the report
'                'tmVefSrchKey.iCode = ilAsgnVefCode
'                'ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'                ilRet = gBinarySearchVef(ilAsgnVefCode)
'                If ilRet >= 0 Then
'                'If ilRet = BTRV_ERR_NONE Then
'                    'tmCopyCntr(ilUpper).sVehName = Trim$(tlVef.sName)
'                    tmCopyCntr(ilUpper).sVehName = Trim$(tgMVef(ilRet).sName)
'                Else
'                    tmCopyCntr(ilUpper).sVehName = slName
'                End If
'            Else
'                tmCopyCntr(ilUpper).sVehName = slName
'            End If
'            tmCopyCntr(ilUpper).sAirVehName = slName
'            tmCopyCntr(ilUpper).iLen = tmSdf.iLen
'            tmCopyCntr(ilUpper).sLiveFlag = slLive          '11-16-05  live or recorded spot
'
'            ReDim Preserve tmCopyCntr(0 To ilUpper + 1) As COPYCNTRSORT
'            ilUpper = ilUpper + 1
'        End If
'    End If
'Return
End Sub
'*****************************************************************
'*
'*      Procedure Name:mObtainCopyDate
'*
'*             Created:10/09/93      By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:Obtain the Sdf records to be
'*                     reported
'*
'*          dh 8/18/00 Add Regional copy feature
'*          7-27-04 Test to include/exclude contract/feed spots
'*******************************************************************
Sub mObtainCopyDate(ilRptType As Integer, ilVefCode As Integer, ilVpfIndex As Integer, slStartDate As String, slEndDate As String, ilSpotType As Integer, ilIncludeUnassigned As Integer, ilCntrSpots As Integer, ilFeedSpots As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'
'   Where
'       ilRptType(I)- 0= Sort for by Date; 1= Sort for by Advertiser
'       ilVefCode(I) - vehicle code
'       ilVpfIndex(I) - VPF index to vehicle processing
'       slStartDate(I) - start of spots to gather
'       slEndDAte(I) - end date of spots to gather
'       ilSpotType(I)- 0= All spots; 1=with copy; 2=without copy
'       ilIncludeUnassigned(I) - include unassigned copy (true/false)
'       ilContrSpots (I) - include contract spots (true/false)
'       ilFeedSpots (I) - include feed spots (true/false)
'
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slCode As String
    Dim llUpper As Long      '7-17-09 changed from integer to long
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilCntLoop As Integer
    Dim slCntrNo As String
    Dim slSdfDate As String
    Dim slSdfTime As String
    Dim llSdfTime As Long
    Dim ilAssign As Integer
    Dim ilSpotSeqNo As Integer
    Dim ilAsgnVefCode As Integer
    Dim slNameCode As String
    Dim ilPkgRot As Integer
    Dim ilPkgVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim ilRotNo As Integer
    Dim ilPkgRotNo As Integer
    Dim ilLnVefCode As Integer
    Dim ilLnRotNo As Integer
    Dim ilLnRet As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim ilRegionalFlag As Integer       '8-11-00
    Dim tlVef As VEF
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim ilPkgNonRegionDefined As Integer    '8-11-00
    Dim ilPkgRegionMissing As Integer
    Dim ilPkgRegionSuperseded As Integer
    Dim ilLnNonRegionDefined As Integer    '8-11-00
    Dim ilLnRegionMissing As Integer
    Dim ilLnRegionSuperseded As Integer
    Dim llSaveCopyCode As Long           '8-16-00 saved SDF code when processing regional copy
    Dim ilRegionRotNo As Integer
    Dim ilPkgRegionRotNo As Integer
    Dim ilLnRegionRotNo As Integer
    Dim ilIncludeUnAssg As Integer  '5-20-05
    Dim ilIncludeReassg As Integer
    Dim slStr As String
    Dim slType As String
    
    Dim ilGameNo As Integer
    ilIncludeUnAssg = True          '5-20-05 include unassigned, force to include both unassigned and to be reassigned,
                                    'for common subroutine
    ilIncludeReassg = True

    lmTodayDate = gDateValue(gNow())        'used in mAssignCopyTest
    llUpper = UBound(tmCopy)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilVefCode
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

         '7-21-04  Exclude/include contract/feed spots
        tlLongTypeBuff.lCode = 0
        If Not ilCntrSpots Or Not ilFeedSpots Then           'either local or feed spots are to be excluded
            If ilCntrSpots Then                         'include local only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
            ElseIf ilFeedSpots Then                      'include feed only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
            End If
        End If

        tlIntTypeBuff.iType = ilVefCode
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
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                If ilRptType = 0 Then   'By Date
                    ilFound = True
                Else                    'status by advt
                    If RptSel!ckcAll.Value = vbChecked Then         'if all is set, no need to loop
                        ilFound = True
                    Else
                        For illoop = LBound(tmSelAdvt) To UBound(tmSelAdvt) - 1 Step 1
                            If tmSelAdvt(illoop) = tmSdf.iAdfCode Then
                                'advertisers match, if spot is feed, then found valid spot;
                                'otherwise test the selection of contracts
                                If tmSdf.lChfCode = 0 Then
                                    ilFound = True
                                Else
                                    For ilCntLoop = 0 To RptSel!lbcSelection(5).ListCount - 1 Step 1
                                        If RptSel!lbcSelection(5).Selected(ilCntLoop) Then
                                            slNameCode = RptSel!lbcSelection(3).List(ilCntLoop)
                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                            If tmSdf.lChfCode = Val(slCode) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilCntLoop
                                End If
                            End If
                        Next illoop

                    End If
                End If
                If ilFound Then
                    gUnpackDateForSort tmSdf.iDate(0), tmSdf.iDate(1), slSdfDate
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llSdfTime
                    slSdfTime = Trim$(str$(llSdfTime))
                    Do While Len(slSdfTime) < 6
                        slSdfTime = "0" & slSdfTime
                    Loop
                    ilSpotSeqNo = mGetSeqNo(tmSdf)
                    If ilSpotSeqNo < 10 Then
                        slSdfTime = slSdfTime & "0" & Trim$(str$(ilSpotSeqNo)) & "0"    '8-16-00 nonregion record (vs region = 1)
                    Else
                        slSdfTime = slSdfTime & Trim$(str$(ilSpotSeqNo)) & "0"     '8-16-00 nonregion record (vs region = 1)
                    End If
                    ilGameNo = tmSdf.iGameNo
                    'ilAssign = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode)
                    ilSchPkgVefCode = 0
                    ilAsgnVefCode = 0
                    ilLnVefCode = 0
                    ilPkgVefCode = 0
                    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilAsgnVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSpotType = "X") Then
                        slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, hmSmf, hmCrf)
                        If (slStr = "S") Or (slStr = "B") Then
                            ilSchPkgVefCode = gGetMGPkgVefCode(hmClf, tmSdf)
                        End If
                        If slStr = "O" Then
                            ilAsgnVefCode = ilLnVefCode
                            ilLnVefCode = 0
                        ElseIf slStr = "S" Then
                            ilPkgVefCode = ilSchPkgVefCode
                            ilSchPkgVefCode = 0
                            ilLnVefCode = 0
                        Else
                            If ilPkgVefCode = ilSchPkgVefCode Then
                                ilSchPkgVefCode = 0
                            End If
                        End If
                    Else
                        ilLnVefCode = 0
                    End If
                    
                    '6/15/15: added the mBuildCRFByCntr as mAssignCopyTest uses the array tmCRFByCntr build by mBuildCRFByCntr
                    If tmSdf.sSpotType = "C" Then
                        slType = "C"
                    ElseIf tmSdf.sSpotType = "O" Then
                        slType = "O"
                    Else
                        slType = "A"
                    End If
                    mBuildCRFByCntr ilSchPkgVefCode, ilAsgnVefCode, ilLnVefCode, ilPkgVefCode, slSdfDate, slType, slLive
                    '6/15/15: end of change
                    
                    'ilRet = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo)
                    ilRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
                    ilAssign = ilRet
                    If ilPkgVefCode > 0 Then
                        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
                        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
                        If ilPkgRegionRotNo > ilRegionRotNo Then
                            '8-17-00
                            imNonRegionDefined = ilPkgNonRegionDefined
                            imRegionMissing = ilPkgRegionMissing
                            imRegionSuperseded = ilPkgRegionSuperseded
                        End If

                        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
                            If ilPkgRotNo > ilRotNo Then
                                ilRotNo = ilPkgRotNo
                                ilAsgnVefCode = ilPkgVefCode
                                ilAssign = ilPkgRot
                            End If
                        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
                            ilRotNo = ilPkgRotNo
                            ilAsgnVefCode = ilPkgVefCode
                            ilAssign = ilPkgRot
                        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
                            ilRotNo = ilPkgRotNo
                            ilAsgnVefCode = ilPkgVefCode
                            ilAssign = ilPkgRot
                        End If
                    End If
                    If ilSchPkgVefCode > 0 Then
                        ilPkgVefCode = ilSchPkgVefCode
                        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
                        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
                        If ilPkgRegionRotNo > ilRegionRotNo Then
                            '8-17-00
                            imNonRegionDefined = ilPkgNonRegionDefined
                            imRegionMissing = ilPkgRegionMissing
                            imRegionSuperseded = ilPkgRegionSuperseded
                        End If

                        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
                            If ilPkgRotNo > ilRotNo Then
                                ilRotNo = ilPkgRotNo
                                ilAsgnVefCode = ilPkgVefCode
                                ilAssign = ilPkgRot
                            End If
                        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
                            ilRotNo = ilPkgRotNo
                            ilAsgnVefCode = ilPkgVefCode
                            ilAssign = ilPkgRot
                        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
                            ilRotNo = ilPkgRotNo
                            ilAsgnVefCode = ilPkgVefCode
                            ilAssign = ilPkgRot
                        End If
                    End If
                    If (ilAsgnVefCode <> ilLnVefCode) And (ilLnVefCode > 0) Then
                        ilLnRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilLnVefCode, ilLnRotNo, ilLnNonRegionDefined, ilLnRegionMissing, ilLnRegionSuperseded, ilLnRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
                        If ilLnRegionRotNo > ilRegionRotNo Then
                            '8-17-00
                            imNonRegionDefined = ilLnNonRegionDefined
                            imRegionMissing = ilLnRegionMissing
                            imRegionSuperseded = ilLnRegionSuperseded
                        End If
                        If (ilAssign <> 0) And (ilLnRet <> 0) Then
                            If ilLnRotNo > ilRotNo Then
                                ilRotNo = ilLnRotNo
                                ilAssign = ilLnRet
                                ilAsgnVefCode = ilLnVefCode
                            End If
                        ElseIf (ilAssign = 0) And (ilLnRet <> 0) Then
                            ilRotNo = ilLnRotNo
                            ilAssign = ilLnRet
                            ilAsgnVefCode = ilLnVefCode
                        End If
                    End If

                    ilRegionalFlag = mGetRegionalFlag()     '8-11-00 Convert the different regional copy warnings/errors into 1 flag
                    tmCopy(llUpper).iRegionalStatus = ilRegionalFlag   'error flag for regional copy
                    tmCopy(llUpper).iRegionalSort = 0                 'normal spot line comes before regional copy line (if any)

                    If ilAssign <> 4 Then
                        'Bypass spots that have Rotation that can be assigned
                        If ilAssign <> 0 Then
                            If (ilIncludeUnassigned) Or (ilSpotType = 0) Then
                                tmCopy(llUpper).iCopyStatus = ilAssign + 1
                                '6/8/16: Replaced GoSub
                                'GoSub lProcMakeRec
                                mProcMakeRec llUpper, ilRptType, slSdfDate, slSdfTime, ilAsgnVefCode
                            'Else
                            '    If (ilSpotType = 0) Then    'Or (ilSpotType = 1) Or (ilSpotType = 2) Then     '0=All spots; 1=With copy; 2=Without copy
                            '        tmCopy(llUpper).iCopyStatus = 2
                            '        GoSub lProcMakeRec
                            '    End If
                            End If
                        Else
                            'No copy
                            If (ilSpotType = 0) Or (ilSpotType = 2) Then
                                tmCopy(llUpper).iCopyStatus = 0
                                '6/8/16: Replaced GoSub
                                'GoSub lProcMakeRec
                                mProcMakeRec llUpper, ilRptType, slSdfDate, slSdfTime, ilAsgnVefCode
                            End If
                        End If
                    Else
                        If (ilSpotType = 0) Or (ilSpotType = 1) Then
                            tmCopy(llUpper).iCopyStatus = 1
                            '6/8/16: Replaced GoSub
                            'GoSub lProcMakeRec
                            mProcMakeRec llUpper, ilRptType, slSdfDate, slSdfTime, ilAsgnVefCode
                        End If
                    End If
                'End If     '8-16-00
                    '8-16-00 See if theres any regional copy to show
                    '8-27-09 Currently, the 2 reports that call this filter out the data.
                    'go around this code to prevent creating them in prepass.
                    'Create prepass for anything other report that comes thru mObtaincopydate
                    If ilRptType <> 0 And ilRptType <> 1 Then   'do not create for Copy STatus by date or advertiser

                        tmRsfSrchKey1.lCode = tmSdf.lCode
                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        Do While ilRet = BTRV_ERR_NONE And tmSdf.lCode = tmRsf.lSdfCode
                            llSaveCopyCode = tmSdf.lCopyCode                 'fake out the spot copy pointer so that subrtn are common to retrieve copy information
                            tmSdf.lCopyCode = tmRsf.lCopyCode
                            slSdfTime = Trim$(str$(llSdfTime))
                            Do While Len(slSdfTime) < 6
                                slSdfTime = "0" & slSdfTime
                            Loop
                            ilSpotSeqNo = mGetSeqNo(tmSdf)
                            If ilSpotSeqNo < 10 Then
                                slSdfTime = slSdfTime & "0" & Trim$(str$(ilSpotSeqNo)) & "1"    '8-16-00 nonregion record (vs region = 1)
                            Else
                                slSdfTime = slSdfTime & Trim$(str$(ilSpotSeqNo)) & "1"          '8-16-00 nonregion record (vs region = 1)
                            End If
                            tmCopy(llUpper).iRegionalSort = 1                 'regional copy line follows spot line
                            '6/8/16: Replaced GoSub
                            'GoSub lProcMakeRec
                            mProcMakeRec llUpper, ilRptType, slSdfDate, slSdfTime, ilAsgnVefCode
                            tmSdf.lCopyCode = llSaveCopyCode     'restor the orig SDF code to continue processing the remaining spots
                            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                End If
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
'lProcMakeRec:
'    tmCopy(llUpper).tSdf = tmSdf
'    If ilRptType = 0 Then   'By Date
'        If tmVef.iCode <> tmSdf.iVefCode Then
'            tmVefSrchKey.iCode = tmSdf.iVefCode
'            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'        End If
'        tmCopy(llUpper).sKey = tmVef.sName & slSdfDate & slSdfTime
'    Else                        'by advt
'        If tmSdf.lChfCode = 0 Then              'feed spot
'            tmFSFSrchKey.lCode = tmSdf.lFsfCode
'            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'            slCntrNo = Trim$(tmFsf.sRefID)
'            Do While Len(slCntrNo) < 10
'                slCntrNo = "0" & slCntrNo
'            Loop
'
'        Else                                    'contract spot
'            If tmChf.lCode <> tmSdf.lChfCode Then
'                tmChfSrchKey.lCode = tmSdf.lChfCode
'                ilRet = btrGetEqual(hmCHF, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'            End If
'            slCntrNo = Trim$(str$(tmChf.lCntrNo))
'            Do While Len(slCntrNo) < 10         'chged from 8 to 10 for Feed spot ref ID
'                slCntrNo = "0" & slCntrNo
'            Loop
'        End If
'
'        ilLoop = gBinarySearchAdf(tmSdf.iAdfCode)
'        slCode = ""
'        If ilLoop <> -1 Then
'            ''slCode = Trim$(tgCommAdf(ilLoop).sName)
'            'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
'            '    slCode = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
'            'Else
'                slCode = Trim$(tgCommAdf(ilLoop).sName)
'            'End If
'        End If
'        tmCopy(llUpper).sKey = slCode & slCntrNo & slSdfDate & slSdfTime
'    End If
'    If ilAsgnVefCode <> tmVef.iCode Then
'        tmVefSrchKey.iCode = ilAsgnVefCode
'        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'        tmCopy(llUpper).sVehName = tlVef.sName
'    Else
'        tmCopy(llUpper).sVehName = tmVef.sName
'    End If
'    ReDim Preserve tmCopy(0 To llUpper + 1) As COPYSORT
'    llUpper = llUpper + 1
'    Return
End Sub
'*******************************************************
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSdf                      *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified: 11/20/96     By:d.h.           *
'*                                                     *
'*            Comments:Obtain the Sdf records to be    *
'*                     reported                        *
'*                                                     *
'*      7-21-04 include/exclude contract/feed spots    *
'*******************************************************
Sub mObtainSdf(ilVefCode As Integer, slStartDate As String, slEndDate As String, ilSpotType As Integer, ilBillType As Integer, ilIncludePSA As Integer, ilMissedType As Integer, ilISCIOnly As Integer, ilCostType As Integer, ilByOrderOrAir As Integer, ilCntrSpots As Integer, ilFeedSpots As Integer)
'
'
'   where:
'       ilSpotType(I)-1=Scheduled only, 2=Missed Only; 3=Both
'       ilBillType(I)-1=Billed only, 2=Unbilled Only; 3=Both; 0=Neither
'       ilMissedType(I)-Bits 15-0 (left to right)
'                       Bit   Meaning
'                         0   Missed (U, M, R)
'                         1   Cancelled (C)
'                         2   Hidden (H)
'       ilCosttype(I) - if negative, ignore spot type tests.  Otherwise,
'                       bit string of spots types to include
'                       bit 0 = charged, 1 = .00, 2 = adu, 3 = bonus, 4 = extra,
'                       5 = fill, 6 = n/c, 7 = mg, 8 = recapturable, 9 = spinoff
'       ilByOrderOrAir(I)- 0=Order; 1=Aired
'       ilCntrSpots(I) - true if include contract spots
'       ilFeedSpots (I) - true if include feed spots

    Dim slDate As String
    Dim llTime As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilOk As Integer
    Dim ilSpotSeqNo As Integer
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim ilVpfIndex As Integer
    Dim slLLDate As String
    Dim llLLDate As Long

    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    
    
    'ignore  BB from days in future
    ilVpfIndex = gBinarySearchVpf(ilVefCode)
    If ilVpfIndex <> -1 Then
        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slLLDate
        If slLLDate = "" Then
            slLLDate = Format(Now, "m/d/yy")
        Else
            If gDateValue(slLLDate) < gDateValue(Format(Now, "m/d/yy")) Then
                slLLDate = Format(Now, "m/d/yy")
            End If
        End If
        slLLDate = gIncOneDay(slLLDate)
    Else
        slLLDate = gIncOneDay(Format(Now, "m/d/yy"))
    End If
    llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater
    
    ilUpper = UBound(tmPLSdf)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilVefCode
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

        '7-32-04  include/Exclude cntr/network spots
        tlLongTypeBuff.lCode = 0
        If Not ilCntrSpots Or Not ilFeedSpots Then           'either local or feed spots are to be excluded
            If ilCntrSpots Then                         'include contract spots only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
            ElseIf ilFeedSpots Then                      'include feed only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
            End If
        End If

        tlIntTypeBuff.iType = ilVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        If (slStartDate <> "") Or (slEndDate <> "") Then
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        Else
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        End If
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
        ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'Build sort key
                ilOk = True
                If ilByOrderOrAir = 0 Then
                    'Schedule and Missed only
                    If (tmPLSdf(ilUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "O") Then
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus <> "S") Then
                            ilOk = False
                        End If
                    Else
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus = "H") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "C") Then
                            ilOk = False
                        End If
                    End If
                Else
                    If ilSpotType = 1 Then  'Scheduled only
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus <> "S") And (tmPLSdf(ilUpper).tSdf.sSchStatus <> "G") And (tmPLSdf(ilUpper).tSdf.sSchStatus <> "O") Then
                            ilOk = False
'                        Else
'                            If ilBillType = 1 Then  'Billed only
'                                If (tmPLSdf(ilUpper).tSdf.sBill <> "Y") Then
'                                    ilOk = False
'                                End If
'                            ElseIf ilBillType = 2 Then  'Unbilled only
'                                If (tmPLSdf(ilUpper).tSdf.sBill = "Y") Then
'                                    ilOk = False
'                                End If
'                            ElseIf ilBillType = 0 Then  'Neither
'                                ilOk = False
'                            End If
                        End If
                    ElseIf ilSpotType = 2 Then  'Missed only
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "O") Then
                            ilOk = False
                        Else
                            If (tmPLSdf(ilUpper).tSdf.sSchStatus = "H") Then
                                If (ilMissedType And &H4) <> &H4 Then
                                    ilOk = False
                                End If
                            ElseIf (tmPLSdf(ilUpper).tSdf.sSchStatus = "C") Then
                                If (ilMissedType And &H2) <> &H2 Then
                                    ilOk = False
                                End If
                            Else
                                If (ilMissedType And &H1) <> &H1 Then
                                    ilOk = False
                                End If
                            End If
                        End If
                    Else
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus <> "S") And (tmPLSdf(ilUpper).tSdf.sSchStatus <> "G") And (tmPLSdf(ilUpper).tSdf.sSchStatus <> "O") Then
                            If (tmPLSdf(ilUpper).tSdf.sSchStatus = "H") Then
                                If (ilMissedType And &H4) <> &H4 Then
                                    ilOk = False
                                End If
                            ElseIf (tmPLSdf(ilUpper).tSdf.sSchStatus = "C") Then
                                If (ilMissedType And &H2) <> &H2 Then
                                    ilOk = False
                                End If
                            Else
                                If (ilMissedType And &H1) <> &H1 Then
                                    ilOk = False
                                End If
                            End If
                        End If
                    End If
                    If ilBillType = 1 Then  'Billed only
                        If (tmPLSdf(ilUpper).tSdf.sBill <> "Y") Then
                            ilOk = False
                        End If
                    ElseIf ilBillType = 2 Then  'Unbilled only
                        If (tmPLSdf(ilUpper).tSdf.sBill = "Y") Then
                            ilOk = False
                        End If
                    ElseIf ilBillType = 0 Then  'Neither
                        ilOk = False
                    End If
                        'End If
                    'End If
                End If
                If (ilOk) And (Not ilIncludePSA) Then
                    If tmPLSdf(ilUpper).tSdf.lChfCode > 0 Then          'only test for psa spot if its a contract spot (vs network feed spot)
                        tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                            ilOk = False
                        End If
                    End If
                End If
                If (ilOk) And (ilISCIOnly) Then
                    If tmPLSdf(ilUpper).tSdf.lChfCode > 0 Then          'copy only applies to contract spots
                  '**********  NEED TO IMPLEMENT NETWORK COPY   **************
                        tmSdf = tmPLSdf(ilUpper).tSdf
                        mObtainCopy slProduct, slZone, slCart, slISCI, slProduct
                        If Len(slISCI) <= 0 Then
                            tmAdfSrchKey.iCode = tmPLSdf(ilUpper).tSdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            'If tmAdf.sShowISCI <> "Y" Then
                            If tmAdf.sShowISCI = "N" Then           '3-10-15 Y or T= show isci
                                tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                If tmChf.iAgfCode > 0 Then     'agency exists
                                    tmAgfSrchKey.iCode = tmChf.iAgfCode
                                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    'If tmAgf.sShowISCI <> "Y" Then
                                    If tmAgf.sShowISCI = "N" Then           '3-10-15 Y or T = Show ISCI
                                        ilOk = False
                                    End If
                                Else
                                    ilOk = False
                                End If
                            End If
                        Else                'isci defined
                            ilOk = False    'spot has copy & isci codes, ignore it
                        End If
                    End If
                End If
                
                'Test if Open or Close BB, ignore if in the future
                If tmPLSdf(ilUpper).tSdf.sSpotType = "O" Or tmPLSdf(ilUpper).tSdf.sSpotType = "C" Then
                    gUnpackDate tmPLSdf(ilUpper).tSdf.iDate(0), tmPLSdf(ilUpper).tSdf.iDate(1), slDate
                    If gDateValue(slDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
                        ilOk = False
                    End If
                End If
                
                If ilOk Then
                    '************ CHECK TO SEE IF THIS READ CAN COME OUT BECAUSE GGETSPOT PRICE GETS LINE *********
                    'get line first, to send to filter routine
                    tmPLSdf(ilUpper).sLiveCopy = ""
                    If tmPLSdf(ilUpper).tSdf.lChfCode > 0 Then          'only test for spot costs if its a contract spot (vs network feed spot)
                        tmClfSrchKey.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode
                        tmClfSrchKey.iLine = tmPLSdf(ilUpper).tSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(ilUpper).tSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(ilUpper).tSdf.iLineNo) Then
                            ilRet = gGetSpotPrice(tmPLSdf(ilUpper).tSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, tmPLSdf(ilUpper).sCostType)
                            tmPLSdf(ilUpper).iVefCode = tmClf.iVefCode
                            If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                                mTestCostType ilOk, ilCostType, tmPLSdf(ilUpper).sCostType
                            End If
                            'If Not ilOk Then
                                'ilOk = False
                            'End If
                        Else
                            ilOk = False
                        End If
                        tmPLSdf(ilUpper).sLiveCopy = tmClf.sLiveCopy    '5-31-12
                    Else
                        tmPLSdf(ilUpper).sCostType = "Feed"
                    End If

                End If
                If ilOk Then
                    ilSpotSeqNo = mGetSeqNo(tmPLSdf(ilUpper).tSdf)
                    tmPLSdf(ilUpper).sKey = tmVef.sName
                    gUnpackDateForSort tmPLSdf(ilUpper).tSdf.iDate(0), tmPLSdf(ilUpper).tSdf.iDate(1), slDate
                    tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slDate
                    If (tmPLSdf(ilUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "O") Then
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|A"
                    Else
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|Z"
                    End If
                    gUnpackTimeLong tmPLSdf(ilUpper).tSdf.iTime(0), tmPLSdf(ilUpper).tSdf.iTime(1), False, llTime
                    slStr = Trim$(str$(llTime))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    If ilSpotSeqNo < 10 Then
                        slStr = slStr & "0" & Trim$(str$(ilSpotSeqNo))
                    Else
                        slStr = slStr & Trim$(str$(ilSpotSeqNo))
                    End If
                    tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slStr
                    ReDim Preserve tmPLSdf(0 To ilUpper + 1) As SPOTTYPESORT
                    ilUpper = ilUpper + 1
                End If
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub

    'ilRet = Err.Number
    'Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCefRec                     *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Read in comment record         *
'*                                                     *
'*******************************************************
Sub mReadCefRec(llCefCode As Long)
'
'   iRet = mReadCefRec()
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    tmCefSrchKey.lCode = llCefCode
    If llCefCode <> 0 Then
        tmCef.sComment = ""
        imCefRecLen = Len(tmCef)    '1009
        ilRet = btrGetEqual(hmCef, tmCef, imCefRecLen, tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmCef.lCode = 0
            'tmCef.iStrLen = 0
            tmCef.sComment = ""
        End If
    Else
        tmCef.lCode = 0
        'tmCef.iStrLen = 0
        tmCef.sComment = ""
    End If
    Exit Sub
End Sub
'
'                   sub mTestCostType - for Spots by Advt and spots by Date & time
'                   Include different types of spots test
'                   <input> ilCosttype - bit string based on user request of
'                           types of spots to include
'                           slStrCost - string defining cost of spot ($ value as string
'                                       or text such as ADU , bonus, etc.
'                   <output> ilOk - false if not a spot to report
'
Sub mTestCostType(ilOk As Integer, ilCostType As Integer, slStrCost As String)
    'look for inclusion of charge spots with
    If (InStr(slStrCost, ".") <> 0) Then        'found spot cost
        'is it a .00?
        If gCompNumberStr(slStrCost, "0.00") = 0 Then       'its a .00 spot
            If (ilCostType And SPOT_00) <> SPOT_00 Then      'include .00?
                ilOk = False
            End If
        Else
            If (ilCostType And SPOT_CHARGE) <> SPOT_CHARGE Then    'include charged spots?
                ilOk = False                                            'no
            End If
        End If
    ElseIf Trim$(slStrCost) = "ADU" And (ilCostType And SPOT_ADU) <> SPOT_ADU Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Bonus" And (ilCostType And SPOT_BONUS) <> SPOT_BONUS Then
            ilOk = False
    'ElseIf Trim$(slStrCost) = "Extra" And (ilCostType And SPOT_EXTRA) <> SPOT_EXTRA Then
    ElseIf Trim$(slStrCost) = "+ Fill" And (ilCostType And SPOT_EXTRA) <> SPOT_EXTRA Then
            ilOk = False
    'ElseIf Trim$(slStrCost) = "Fill" And (ilCostType And SPOT_FILL) <> SPOT_FILL Then
    ElseIf Trim$(slStrCost) = "- Fill" And (ilCostType And SPOT_FILL) <> SPOT_FILL Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "N/C" And (ilCostType And SPOT_NC) <> SPOT_NC Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "MG" And (ilCostType And SPOT_MG) <> SPOT_MG Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Recapturable" And (ilCostType And SPOT_RECAP) <> SPOT_RECAP Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Spinoff" And (ilCostType And SPOT_SPINOFF) <> SPOT_SPINOFF Then
            ilOk = False
    End If
End Sub
'
'
'           Open all files for Copy Status by Date
'           and Copy Status by Advertiser
'       <input>  None
'       <output>  Return: 0 = no error
'           7-27-04
'
Public Function mOpenCopyStatusFiles() As String
Dim ilRet As Integer
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Chf"
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Clf"
        Exit Function
    End If
    imClfRecLen = Len(tmClf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Adf"
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Vef"
        Exit Function
    End If
    imVefRecLen = Len(tmVef)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Sdf"
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)

    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCrf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Crf"
        Exit Function
    End If
    imCrfRecLen = Len(tmCrf)

    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmTzf
        btrDestroy hmCrf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Tzf"
        Exit Function
    End If
    imTzfRecLen = Len(tmTzf)

    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmCrf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "SSF"
        Exit Function
    End If
    imSsfRecLen = Len(tmSsf)

    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "VSf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmCrf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Vsf"
        Exit Function
    End If

    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmCrf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Sif"
        Exit Function
    End If
    imSifRecLen = Len(tmSif)

    '8-11-2000
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Raf"
        Exit Function
    End If
    imRafRecLen = Len(tmRaf)

    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Rsf"
        Exit Function
    End If
    imRsfRecLen = Len(tmRsf)

    hmCpr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Cpr"
        Exit Function
    End If
    imCprRecLen = Len(tmCpr)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Fsf"
        Exit Function
    End If
    imFsfRecLen = Len(tmFsf)

    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Prf"
        Exit Function
    End If
    imPrfRecLen = Len(tmPrf)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAnf
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Anf"
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)

    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGsf
        btrDestroy hmAnf
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Gsf"
        Exit Function
    End If
    imGsfRecLen = Len(tmGsf)

    hmCaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCaf, "", sgDBPath & "Caf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCaf)
        ilRet = btrClose(hmGsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCaf
        btrDestroy hmGsf
        btrDestroy hmAnf
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Caf"
        Exit Function
    End If
    imCafRecLen = Len(tmCaf)
    
    hmCvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCvf)
        ilRet = btrClose(hmCaf)
        ilRet = btrClose(hmGsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmRsf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCvf
        btrDestroy hmCaf
        btrDestroy hmGsf
        btrDestroy hmAnf
        btrDestroy hmPrf
        btrDestroy hmFsf
        btrDestroy hmCpr
        btrDestroy hmRsf
        btrDestroy hmRaf
        btrDestroy hmSif
        btrDestroy hmVsf
        btrDestroy hmSsf
        btrDestroy hmTzf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        mOpenCopyStatusFiles = "Cvf"
        Exit Function
    End If
    imCvfRecLen = Len(tmCvf)
    
    mOpenCopyStatusFiles = ""           'no error
    Exit Function
End Function
'
'           mCloseCopyFiles - close common copy files for Copy Status by
'           Date,Copy Status by Advertiser,  & Contracts Missing Copy
'
Public Sub mCloseCopyFiles()
Dim ilRet As Integer

    ilRet = btrClose(hmCaf)
    ilRet = btrClose(hmGsf)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmRsf)
    ilRet = btrClose(hmRaf)
    ilRet = btrClose(hmSif)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmTzf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCrf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmCaf
    btrDestroy hmGsf
    btrDestroy hmAnf
    btrDestroy hmPrf
    btrDestroy hmFsf
    btrDestroy hmCpr
    btrDestroy hmRsf
    btrDestroy hmRaf
    btrDestroy hmSif
    btrDestroy hmVsf
    btrDestroy hmSsf
    btrDestroy hmTzf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmAdf
    btrDestroy hmCrf
    btrDestroy hmClf
    btrDestroy hmCHF
    Exit Sub
End Sub
'
'           mGetEarliestCopyDate - test the date that copy is being tested and
'           determine if the copy missed, unassigned or to be assigned is earlier
'           than the earliest found so far.
'           <input>  assume spot date in tmSdf
'           <output> tlCopyCnt(0 to 1) new spot date if earlier
Public Sub mGetEarliestCopyDate(ilDate1 As Integer, ilDate2 As Integer) 'VBC NR
Dim llCurrentEArliestDate As Long 'VBC NR
Dim llSpotDate As Long 'VBC NR


        gUnpackDateLong ilDate1, ilDate2, llCurrentEArliestDate 'VBC NR
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSpotDate 'VBC NR
        If llCurrentEArliestDate = 0 Then 'VBC NR
            ilDate1 = tmSdf.iDate(0) 'VBC NR
            ilDate2 = tmSdf.iDate(1) 'VBC NR
        Else 'VBC NR
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSpotDate 'VBC NR
            If llSpotDate < llCurrentEArliestDate Then 'VBC NR
                gPackDateLong llSpotDate, ilDate1, ilDate2 'VBC NR
            End If 'VBC NR
        End If 'VBC NR
End Sub 'VBC NR

'Public Sub mBuildCRFByCntr(ilVefCode As Integer, slActiveDate As String, slType As String, slLive As String)
Public Sub mBuildCRFByCntr(ilSchPkgVefCode As Integer, ilAsgnVefCode As Integer, ilLnVefCode As Integer, ilPkgVefCode As Integer, slActiveDate As String, slType As String, slLive As String)
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim blFound As Boolean
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim tlIntTypeBuff As POPICODE   'Type field record
    Dim slDate As String
    Dim blBypassCrf As Boolean

    ReDim tmCRFByCntr(0 To 0) As CRFBYCNTR
    
    imCrfRecLen = Len(tmCrf)
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    tmCrfSrchKey4.sRotType = slType
    tmCrfSrchKey4.iEtfCode = 0
    tmCrfSrchKey4.iEnfCode = 0
    tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
    tmCrfSrchKey4.lChfCode = tmSdf.lChfCode
    tmCrfSrchKey4.lFsfCode = 0
    'tmCrfSrchKey4.iVefCode = 0
    tmCrfSrchKey4.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (tmCrf.sRotType = slType) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
       
        'filter out dormant rotations
        tlCharTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Crf", "CrfState")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
        'filter out type of rotations
        tlCharTypeBuff.sType = slType
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If

        'filter etf code
        tlIntTypeBuff.iCode = 0
        ilOffSet = gFieldOffset("Crf", "crfetfcode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
        'filter enf code
        tlIntTypeBuff.iCode = 0
        ilOffSet = gFieldOffset("Crf", "crfenfcode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If


        'filter out adv code
        tlIntTypeBuff.iCode = tmSdf.iAdfCode
        ilOffSet = gFieldOffset("Crf", "crfadfCode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
        'filter out Contract code
        tlLongTypeBuff.lCode = tmSdf.lChfCode
        ilOffSet = gFieldOffset("Crf", "crfChfCode")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
        'filter out lower rotation
        tlIntTypeBuff.iCode = tmSdf.iRotNo
        ilOffSet = gFieldOffset("Crf", "crfrotno")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_GTE, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        
         'filter out spot length
        tlIntTypeBuff.iCode = tmSdf.iLen
        ilOffSet = gFieldOffset("Crf", "crfLen")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
       
        'filter rotations out based on spot date
        If slActiveDate = "" Then
            slDate = "1/1/1970"
        Else
            slDate = slActiveDate
        End If
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfStartDate")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
 
            ilExtLen = Len(tmCrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                blBypassCrf = False
                If slLive = "L" Then
                    If tmCrf.sLiveCopy <> "L" Then
                        blBypassCrf = True
                    End If
                ElseIf slLive = "M" Then
                    If tmCrf.sLiveCopy <> "M" Then
                        blBypassCrf = True
                    End If
                ElseIf slLive = "S" Then
                    If tmCrf.sLiveCopy <> "S" Then
                        blBypassCrf = True
                    End If
                ElseIf slLive = "P" Then
                    If tmCrf.sLiveCopy <> "P" Then
                        blBypassCrf = True
                    End If
                ElseIf slLive = "Q" Then
                    If tmCrf.sLiveCopy <> "Q" Then
                        blBypassCrf = True
                    End If
                Else
                    If (tmCrf.sLiveCopy = "L") Or (tmCrf.sLiveCopy = "M") Or (tmCrf.sLiveCopy = "S") Or (tmCrf.sLiveCopy = "P") Or (tmCrf.sLiveCopy = "Q") Then
                        blBypassCrf = True
                    End If
                End If
                If Not blBypassCrf Then
                    '6-3-15 check to see if any of the vehicles to process is valid for this spots rotation
                    If ilSchPkgVefCode > 0 Then
                        If gCheckCrfVehicle(ilSchPkgVefCode, tmCrf, hmCvf) Then
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf = tmCrf           'save entire record
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf.iVefCode = ilSchPkgVefCode
                            ReDim Preserve tmCRFByCntr(0 To UBound(tmCRFByCntr) + 1) As CRFBYCNTR
                        End If
                    End If
                    
                    If (ilAsgnVefCode > 0) And (ilAsgnVefCode <> ilSchPkgVefCode) Then       'if same as previous already ade
                        If gCheckCrfVehicle(ilAsgnVefCode, tmCrf, hmCvf) Then
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf = tmCrf           'save entire record
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf.iVefCode = ilAsgnVefCode
                            ReDim Preserve tmCRFByCntr(0 To UBound(tmCRFByCntr) + 1) As CRFBYCNTR
                        End If
                    End If
                    
                    If (ilLnVefCode > 0) And ((ilLnVefCode <> ilSchPkgVefCode) And (ilLnVefCode <> ilAsgnVefCode)) Then
                        If gCheckCrfVehicle(ilLnVefCode, tmCrf, hmCvf) Then
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf = tmCrf           'save entire record
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf.iVefCode = ilLnVefCode
                           ReDim Preserve tmCRFByCntr(0 To UBound(tmCRFByCntr) + 1) As CRFBYCNTR
                        End If
                    End If
                    
                    If (ilPkgVefCode > 0) And ((ilPkgVefCode <> ilSchPkgVefCode) And (ilPkgVefCode <> ilAsgnVefCode) And (ilPkgVefCode <> ilLnVefCode)) Then
                        If gCheckCrfVehicle(ilPkgVefCode, tmCrf, hmCvf) Then
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf = tmCrf           'save entire record
                            tmCRFByCntr(UBound(tmCRFByCntr)).tCrf.iVefCode = ilPkgVefCode
                            ReDim Preserve tmCRFByCntr(0 To UBound(tmCRFByCntr) + 1) As CRFBYCNTR
                        End If
                    End If
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
              
        btrExtClear hmCrf   'Clear any previous extend operation
    End If
End Sub

Private Sub mRegionSupersedeTest(ilRegionSuperseded As Integer, ilAvailOk As Integer, slType As String, ilSSFType As Integer, slDate As String, llDate As Long, llSAsgnDate As Long, llEAsgnDate As Long, llSAsgnTime As Long, llEAsgnTime As Long, ilRegionMissing As Integer, ilBypassCrf As Integer, ilDay As Integer, ilCrfVefCode As Integer, slTime As String, ilAsgnDate0 As Integer, ilAsgnDate1 As Integer, llAvailTime As Long, llSpotTime As Long, ilEvtIndex As Integer, ilRegionRotNo As Integer)
    Dim ilLoopOnCrf As Integer
    Dim ilRet As Integer
    
    If ilRegionSuperseded = -1 Then
        tmRafSrchKey1.iAdfCode = tmSdf.iAdfCode
        ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilAvailOk = True
            slType = "A"
            '5-1-15 change to support to new cvf
'            tmCrfSrchKey4.sRotType = slType
'            tmCrfSrchKey4.iEtfCode = 0
'            tmCrfSrchKey4.iEnfCode = 0
'            tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
'            tmCrfSrchKey4.lChfCode = tmSdf.lChfCode     '5-19-05 set starting point of search
'            tmCrfSrchKey4.lFsfCode = tmSdf.lFsfCode 'feed code
'            'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
'            tmCrfSrchKey4.iRotNo = 32000
'            ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get last current record to obtain date
'            Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmSdf.lFsfCode = tmCrf.lFsfCode)
            For ilLoopOnCrf = LBound(tmCRFByCntr) To UBound(tmCRFByCntr) - 1
                tmCrf = tmCRFByCntr(ilLoopOnCrf).tCrf
                lgSupercedeCount = lgSupercedeCount + 1
 
            'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lFsfCode = tmCrf.lFsfCode)
                'Test date, time, day and zone
                ilBypassCrf = False
'                If Not gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf) Then     '5-1-15
'                    ilBypassCrf = True
'                End If
                If (tmCrf.sDay(ilDay) = "Y") And (tmSdf.iLen = tmCrf.iLen) And (tmCrf.sState <> "D") And (Not ilBypassCrf) And (tmCrf.iVefCode = ilCrfVefCode) Then      '5-1-15
                    gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                    llSAsgnDate = gDateValue(slDate)
                    gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                    llEAsgnDate = gDateValue(slDate)
                    If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) Then
                        gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slTime
                        llSAsgnTime = CLng(gTimeToCurrency(slTime, False))
                        gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slTime
                        llEAsgnTime = CLng(gTimeToCurrency(slTime, True)) - 1
                        If (llSpotTime >= llSAsgnTime) And (llSpotTime <= llEAsgnTime) Then
                            ilAvailOk = True    'Ok to book into
                            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
                                Do
                                    If ilEvtIndex > tmSsf.iCount Then
                                        imSsfRecLen = Len(tmSsf)
                                        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tmSdf.iVefCode) And (tmSsf.iDate(0) = ilAsgnDate0) And (tmSsf.iDate(1) = ilAsgnDate1) Then
                                            ilEvtIndex = 1
                                        Else
                                            ilRegionSuperseded = 4
                                            'Return
                                            Exit Sub
                                        End If
                                    End If
                                    'Scan for avail that matches time of spot- then test avail name
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvtIndex)
                                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                                        'Test time-
                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                        llAvailTime = CLng(gTimeToCurrency(slTime, False))
                                        If llSpotTime = llAvailTime Then
                                            If tmCrf.sInOut = "I" Then
                                                If tmCrf.ianfCode <> tmAvail.ianfCode Then
                                                    ilAvailOk = False   'No
                                                End If
                                            Else
                                                If tmCrf.ianfCode = tmAvail.ianfCode Then
                                                    ilAvailOk = False   'No
                                                End If
                                            End If
                                            Exit Do
                                        ElseIf llSpotTime < llAvailTime Then
                                            ilRegionSuperseded = 4
                                            'Return
                                            Exit Sub
                                        End If
                                    End If
                                    ilEvtIndex = ilEvtIndex + 1
                                Loop
                            End If
                            If ilAvailOk Then
                                If Trim$(tmCrf.sZone) <> "R" Then
                                    If Trim$(tmCrf.sZone) = "" Then
                                        tmRsfSrchKey1.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
                                                ilRegionSuperseded = 2
                                                If ilRegionRotNo = -1 Then
                                                    ilRegionRotNo = tmRsf.iRotNo
                                                End If
                                            End If
                                        Else
                                            ilRegionSuperseded = 0
                                        End If
                                        'Return
                                        Exit Sub
                                    End If
                                Else
                                    If ilRegionRotNo = -1 Then
                                        ilRegionRotNo = tmCrf.iRotNo
                                    End If
                                    tmRsfSrchKey1.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
                                        If tmRsf.lRafCode = tmCrf.lRafCode Then
                                            If tmRsf.iRotNo < tmCrf.iRotNo Then
                                                ilRegionSuperseded = 2
                                                'Return
                                                Exit Sub
                                            Else
                                                ilRegionSuperseded = 4
                                                'Return
                                                Exit Sub
                                            End If
                                        End If
                                        ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                    ilRegionSuperseded = 1
                                    'Return
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                'ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Next ilLoopOnCrf
            'Loop
        End If
        ilRegionSuperseded = 0
    End If
End Sub

Private Sub mRegionTest(ilRegionSuperseded As Integer, ilAvailOk As Integer, slType As String, ilSSFType As Integer, slDate As String, llDate As Long, llSAsgnDate As Long, llEAsgnDate As Long, llSAsgnTime As Long, llEAsgnTime As Long, ilRegionMissing As Integer, ilBypassCrf As Integer, ilDay As Integer, ilCrfVefCode As Integer, slTime As String, ilAsgnDate0 As Integer, ilAsgnDate1 As Integer, llAvailTime As Long, llSpotTime As Long, ilEvtIndex As Integer, ilRegionRotNo As Integer)
    Dim ilLoopOnCrf As Integer
    Dim ilRet As Integer
    
    If ilRegionMissing Then
        'Test if any region copy defined
        tmRafSrchKey1.iAdfCode = tmSdf.iAdfCode
        ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            '5-1-15 change to support to new key
'            tmCrfSrchKey4.sRotType = slType
'            tmCrfSrchKey4.iEtfCode = 0
'            tmCrfSrchKey4.iEnfCode = 0
'            tmCrfSrchKey4.iAdfCode = tmSdf.iAdfCode
'            tmCrfSrchKey4.lChfCode = tmSdf.lChfCode         '5-18-05 set the starting point of the search
'            tmCrfSrchKey4.lFsfCode = tmSdf.lFsfCode              'feed code
'            'tmCrfSrchKey1.iVefCode = ilCrfVefCode   'tmSdf.iVefCode
'            tmCrfSrchKey4.iRotNo = 32000
'            ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get last current record to obtain date
'            '5-1-15 remove vehicle test
'            Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode)
            lgRegionCvfCount = lgRegionCvfCount + lgRegionCvfCount
            For ilLoopOnCrf = LBound(tmCRFByCntr) To UBound(tmCRFByCntr) - 1
                tmCrf = tmCRFByCntr(ilLoopOnCrf).tCrf
            'Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iEtfCode = 0) And (tmCrf.iEnfCode = 0) And (tmCrf.iadfCode = tmSdf.iadfCode) And (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode)
                '5-19-05 place matching test with the dowhile to prevent too many reads
                'If (tmCrf.iVefCode = ilCrfVefCode) And (tmSdf.lChfCode <> tmCrf.lChfCode Or tmSdf.lFsfCode <> tmCrf.lFsfCode) Then
                    'Test date, time, day and zone
                    
                'ilBypassCrf = False
'                If Not gCheckCrfVehicle(ilCrfVefCode, tmCrf, hmCvf) Then     '5-1-15
'                    ilBypassCrf = True
'                End If
                'If (tmCrf.sState <> "D") And (Not ilBypassCrf) Then
                    gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slDate
                    llSAsgnDate = gDateValue(slDate)
                    gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                    llEAsgnDate = gDateValue(slDate)
                    If (llDate >= llSAsgnDate) And (llDate <= llEAsgnDate) And (tmCrf.iVefCode = ilCrfVefCode) Then
                        If Trim$(tmCrf.sZone) = "R" Then
                            ilRegionMissing = True
                            'Exit Do
                            Exit For
                        End If
                    End If
                'End If
                
                'End If
               ' ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
               Next ilLoopOnCrf
           'Loop
        Else
            ilRegionMissing = False
        End If
    End If
    '6/8/16: Replaced GoSub
    'GoSub lRegionSupersedeTest
    mRegionSupersedeTest ilRegionSuperseded, ilAvailOk, slType, ilSSFType, slDate, llDate, llSAsgnDate, llEAsgnDate, llSAsgnTime, llEAsgnTime, ilRegionMissing, ilBypassCrf, ilDay, ilCrfVefCode, slTime, ilAsgnDate0, ilAsgnDate1, llAvailTime, llSpotTime, ilEvtIndex, ilRegionRotNo
    Exit Sub
End Sub

Private Sub mCopyCntr(llAirDate As Long, llAirTime As Long, llLatestSellTime As Long, llDate As Long, slStartDate As String, slEndDate As String, slLive As String, ilRdfCode As Integer, ilVpfIndex As Integer, ilCntrSpots As Integer, ilFeedSpots As Integer, ilIncludeUnAssg As Integer, ilIncludeReassg As Integer, ilAssign As Integer, ilUpper As Integer, ilIncludeLine As Integer, slName As String, ilMonRequest() As Integer, ilAiringVeh As Integer)
    Dim ilRet As Integer
    Dim ilGameNo As Integer
    Dim ilAsgnVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim ilLnRotNo As Integer
    Dim ilLnNonRegionDefined As Integer
    Dim ilLnRegionMissing As Integer
    Dim ilLnRegionSuperseded As Integer
    Dim ilLnRegionRotNo As Integer
    Dim ilLnRet As Integer
    Dim slStr As String
    Dim slType As String * 1
    Dim slDate As String
    Dim ilRotNo As Integer
    Dim ilRegionRotNo As Integer
    Dim ilPkgRot As Integer
    Dim ilPkgRotNo As Integer
    Dim ilPkgNonRegionDefined As Integer
    Dim ilPkgRegionMissing As Integer
    Dim ilPkgRegionSuperseded As Integer
    Dim ilPkgRegionRotNo As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRegionalFlag As Integer
    Dim ilFound As Integer
    Dim ilSaveRdfCode As Integer
    Dim illoop As Integer
    Dim ilLineNo As Integer
    Dim slCntrNo As String
    Dim llStartOfFlt As Long
    Dim ilLoopDays As Integer
    Dim ilDay As Integer
    Dim ilInclLineDate0 As Integer
    Dim ilInclLineDate1 As Integer

    
    '7-16-14 test for 13m-3a selling
     If llAirDate = gDateValue(slEndDate) + 1 Then          'processing day+1 because of time zones
        ' llTempAirTime = llAirTime + (ilTZAdj * 3600)
        If llAirTime > llLatestSellTime Then
         'If llTempAirTime >= 0 Then
'             Return
            Exit Sub
         End If
    End If

    ilGameNo = tmSdf.iGameNo
    'ilAssign = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode)
    ilSchPkgVefCode = 0
    ilAsgnVefCode = 0
    ilLnVefCode = 0
    ilPkgVefCode = 0
    ilRet = gGetCrfVefCode(hmClf, tmSdf, ilAsgnVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSpotType = "X") Then
        slStr = gGetMGCopyAssign(tmSdf, ilPkgVefCode, ilLnVefCode, slLive, hmSmf, hmCrf)
        If (slStr = "S") Or (slStr = "B") Then
            ilSchPkgVefCode = gGetMGPkgVefCode(hmClf, tmSdf)
        End If
        If slStr = "O" Then
            ilAsgnVefCode = ilLnVefCode
            ilLnVefCode = 0
        ElseIf slStr = "S" Then
            ilPkgVefCode = ilSchPkgVefCode
            ilSchPkgVefCode = 0
            ilLnVefCode = 0
        Else
            If ilPkgVefCode = ilSchPkgVefCode Then
                ilSchPkgVefCode = 0
            End If
        End If
    Else
        ilLnVefCode = 0
    End If
    
    '6-3-15 see if copy defined for billboards
    If tmSdf.sSpotType = "C" Then
        slType = "C"
    ElseIf tmSdf.sSpotType = "O" Then
        slType = "O"
    Else
        slType = "A"
    End If
    '6-3-15
    'build array of Rotation headers based package vehicle, line vehicle, sched package vehicle so that the crf & cvf do not have to be constantly reread
    'mBuildCRFByCntr ilCrfVefCode, slSpotDate, slType, slLive
    slDate = Format$(llAirDate, "m/d/yy")
    mBuildCRFByCntr ilSchPkgVefCode, ilAsgnVefCode, ilLnVefCode, ilPkgVefCode, slDate, slType, slLive

    'ilRet = mAssignCopyTest("O", ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo)
    ilRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilAsgnVefCode, ilRotNo, imNonRegionDefined, imRegionMissing, imRegionSuperseded, ilRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
    ilAssign = ilRet
    If ilPkgVefCode > 0 Then
        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
        If ilPkgRegionRotNo > ilRegionRotNo Then
            '8-17-00
            imNonRegionDefined = ilPkgNonRegionDefined
            imRegionMissing = ilPkgRegionMissing
            imRegionSuperseded = ilPkgRegionSuperseded
        End If

        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
            If ilPkgRotNo > ilRotNo Then
                ilRotNo = ilPkgRotNo
                ilAsgnVefCode = ilPkgVefCode
                ilAssign = ilPkgRot
            End If
        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
            ilRotNo = ilPkgRotNo
            ilAsgnVefCode = ilPkgVefCode
            ilAssign = ilPkgRot
        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
            ilRotNo = ilPkgRotNo
            ilAsgnVefCode = ilPkgVefCode
            ilAssign = ilPkgRot
        End If
    End If
    If ilSchPkgVefCode > 0 Then
        ilPkgVefCode = ilSchPkgVefCode
        'ilPkgRot = mAssignCopyTest("O", ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo)
        ilPkgRot = mAssignCopyTest(ilGameNo, ilVpfIndex, ilPkgVefCode, ilPkgRotNo, ilPkgNonRegionDefined, ilPkgRegionMissing, ilPkgRegionSuperseded, ilPkgRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
        If ilPkgRegionRotNo > ilRegionRotNo Then
            '8-17-00
            imNonRegionDefined = ilPkgNonRegionDefined
            imRegionMissing = ilPkgRegionMissing
            imRegionSuperseded = ilPkgRegionSuperseded
        End If

        If (ilAssign <> 0) And (ilPkgRot <> 0) Then
            If ilPkgRotNo > ilRotNo Then
                ilRotNo = ilPkgRotNo
                ilAsgnVefCode = ilPkgVefCode
                ilAssign = ilPkgRot
            End If
        ElseIf (ilAssign = 0) And (ilPkgRot = 0) Then
            ilRotNo = ilPkgRotNo
            ilAsgnVefCode = ilPkgVefCode
            ilAssign = ilPkgRot
        ElseIf (ilAssign = 0) And (ilPkgRot <> 0) Then
            ilRotNo = ilPkgRotNo
            ilAsgnVefCode = ilPkgVefCode
            ilAssign = ilPkgRot
        End If
    End If

    If (ilAsgnVefCode <> ilLnVefCode) And (ilLnVefCode > 0) Then
        ilLnRet = mAssignCopyTest(ilGameNo, ilVpfIndex, ilLnVefCode, ilLnRotNo, ilLnNonRegionDefined, ilLnRegionMissing, ilLnRegionSuperseded, ilLnRegionRotNo, ilCntrSpots, ilFeedSpots, ilIncludeUnAssg, ilIncludeReassg, slLive)
        If ilLnRegionRotNo > ilRegionRotNo Then
            '8-17-00
            imNonRegionDefined = ilLnNonRegionDefined
            imRegionMissing = ilLnRegionMissing
            imRegionSuperseded = ilLnRegionSuperseded
        End If
        If (ilAssign <> 0) And (ilLnRet <> 0) Then
            If ilLnRotNo > ilRotNo Then
                ilRotNo = ilLnRotNo
                ilAssign = ilLnRet
                ilAsgnVefCode = ilLnVefCode
            End If
        ElseIf (ilAssign = 0) And (ilLnRet <> 0) Then
            ilRotNo = ilLnRotNo
            ilAssign = ilLnRet
            ilAsgnVefCode = ilLnVefCode
        End If
    End If


    ilRegionalFlag = mGetRegionalFlag()     '8-18-00 Convert the different regional copy warnings/errors into 1 flag

    '12-15-04 change to test for ilRegionalFlag <> 0 (instead of 4) for regional copy OK
    If ilAssign <> 4 Or ilRegionalFlag <> 0 Then    '8-18-00 4 = copy OK
        ilFound = False
        If RptSel!ckcTrans.Value = vbUnchecked Then 'prevent lines from being separated if there are different dayparts for same vehicle
                                                    'and user doesnt want the dp to print
            ilSaveRdfCode = 0
        Else
            ilSaveRdfCode = ilRdfCode
        End If
        For illoop = 0 To ilUpper Step 1

            If ilIncludeLine Then
                ilLineNo = tmSdf.iLineNo
            Else
                ilLineNo = 0
            End If
            If (tmCopyCntr(illoop).lChfCode = tmSdf.lChfCode) And (tmCopyCntr(illoop).iVefCode = tmSdf.iVefCode) And (tmCopyCntr(illoop).iLen = tmSdf.iLen) And (tmCopyCntr(illoop).iAsgnVefCode = ilAsgnVefCode) And (tmCopyCntr(illoop).lFsfCode = tmSdf.lFsfCode) And (tmCopyCntr(illoop).iRdfCode = ilSaveRdfCode) And (tmCopyCntr(illoop).iLineNo = ilLineNo) Then
                ilFound = True


                If (ilAssign = 0) Or (ilAssign = 3) Then 'Copy not defined
                    tmCopyCntr(illoop).iNoSpots = tmCopyCntr(illoop).iNoSpots + 1
                    '7-10-08 if by line, show the first date that a spot has missing copy
                    If tmCopyCntr(illoop).iNoSpots = 1 And ilIncludeLine Then         '1st spot without copy
                        tmCopyCntr(illoop).iStartDate(0) = tmSdf.iDate(0)
                        tmCopyCntr(illoop).iStartDate(1) = tmSdf.iDate(1)
                    End If
                    If tmSdf.sSchStatus = "M" Then
                        tmCopyCntr(illoop).iNoSpotsMiss = 1             'at least 1 spot missed missing copy
                    End If
                    '11-16-05 show the live flag for the missing copy only
                    'if array is blank, nothing has been set yet
                    'l = live, r = recorded, m = mixture of live/recorded
                    If tmCopyCntr(illoop).sLiveFlag <> "X" Then         'if the flag is already mixed, dont touch it
                        If tmCopyCntr(illoop).sLiveFlag <> slLive Then      'must have both live and recorded across the lines
                            tmCopyCntr(illoop).sLiveFlag = "X"
                        End If
                    End If
                ElseIf (ilAssign = 1) Then    'Not assigned
                    tmCopyCntr(illoop).iNoUnAssg = tmCopyCntr(illoop).iNoUnAssg + 1
                    If tmSdf.sSchStatus = "M" Then
                        tmCopyCntr(illoop).iNoUnAssgMiss = 1            'atleast 1 spot missed unassigned
                    End If
                Else    'Supersede
                    tmCopyCntr(illoop).iNoToReassg = tmCopyCntr(illoop).iNoToReassg + 1
                    If tmSdf.sSchStatus = "M" Then
                        tmCopyCntr(illoop).iNoToReassgMiss = 1          'at least 1 spot missed to reassign
                    End If
                End If

                '8-18-00 Accum errors for regional copy
                If ilRegionalFlag = 1 Then      'not assigned
                    tmCopyCntr(illoop).iRegionNoUnAssg = tmCopyCntr(illoop).iRegionNoUnAssg + 1
                    If tmSdf.sSchStatus = "M" Then
                        tmCopyCntr(illoop).iRegionNoUnAssgMiss = 1            'atleast 1 spot missed unassigned
                    End If
                ElseIf ilRegionalFlag = 2 Then      'superseded
                    tmCopyCntr(illoop).iRegionNoToReassg = tmCopyCntr(illoop).iRegionNoToReassg + 1
                    If tmSdf.sSchStatus = "M" Then
                        tmCopyCntr(illoop).iRegionNoToReassgMiss = 1          'at least 1 spot missed to reassign
                    End If
                End If

                Exit For
            End If
        Next illoop
        If Not ilFound Then
            If tmSdf.lChfCode = 0 Then
                tmFSFSrchKey.lCode = tmSdf.lFsfCode
                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                slCntrNo = Trim$(tmFsf.sRefID)
                Do While Len(slCntrNo) < 8
                    slCntrNo = "0" & slCntrNo
                Loop
            Else
                If tmChf.lCode <> tmSdf.lChfCode Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                slCntrNo = Trim$(str$(tmChf.lCntrNo))
                Do While Len(slCntrNo) < 8
                    slCntrNo = "0" & slCntrNo
                Loop
            End If


            If tmAdf.iCode <> tmSdf.iAdfCode Then
                tmAdfSrchKey.iCode = tmSdf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            End If
            tmCopyCntr(ilUpper).sKey = tmAdf.sName & slCntrNo & slName & Trim$(str$(tmSdf.iLen))

            If (ilAssign = 0) Or (ilAssign = 3) Then    'Copy not defined
                tmCopyCntr(ilUpper).iNoSpots = 1
                '7-10-08 if by line, show the first date that a spot has missing copy
                If ilIncludeLine Then         '1st spot without copy
                    tmCopyCntr(ilUpper).iStartDate(0) = tmSdf.iDate(0)
                    tmCopyCntr(ilUpper).iStartDate(1) = tmSdf.iDate(1)
                End If

                tmCopyCntr(ilUpper).iNoUnAssg = 0
                tmCopyCntr(ilUpper).iNoToReassg = 0
                If tmSdf.sSchStatus = "M" Then
                    tmCopyCntr(ilUpper).iNoSpotsMiss = 1        '3-16-07 (fix subscript out of range)
                End If
            ElseIf (ilAssign = 1) Then    'Not assigned
                tmCopyCntr(ilUpper).iNoSpots = 0
                tmCopyCntr(ilUpper).iNoUnAssg = 1
                tmCopyCntr(ilUpper).iNoToReassg = 0
                If tmSdf.sSchStatus = "M" Then
                    'tmCopyCntr(ilLoop).iNoUnAssgMiss = 1
                    tmCopyCntr(ilUpper).iNoUnAssgMiss = 1       '3-16-07 (fix subscript out of range)
                End If
            Else    'Supersede
                tmCopyCntr(ilUpper).iNoSpots = 0
                tmCopyCntr(ilUpper).iNoUnAssg = 0
                tmCopyCntr(ilUpper).iNoToReassg = 1
                If tmSdf.sSchStatus = "M" Then
                    'tmCopyCntr(ilLoop).iNoToReassgMiss = 1
                    tmCopyCntr(ilUpper).iNoToReassgMiss = 1     '3-16-07 (fix subscript out of range)
                End If
            End If
            '8-18-00 Accum errors for regional copy
            If ilRegionalFlag = 1 Then      'not assigned
                tmCopyCntr(ilUpper).iRegionNoUnAssg = 1
                tmCopyCntr(ilUpper).iRegionNoToReassg = 0
                If tmSdf.sSchStatus = "M" Then
                    'tmCopyCntr(ilLoop).iRegionNoUnAssgMiss = 1
                    tmCopyCntr(ilUpper).iRegionNoUnAssgMiss = 1     '3-16-07 (fix subscript out of range)
                End If
            ElseIf ilRegionalFlag = 2 Then      'superseded
                tmCopyCntr(ilUpper).iRegionNoUnAssg = 0
                tmCopyCntr(ilUpper).iRegionNoToReassg = 1
                If tmSdf.sSchStatus = "M" Then
                    'tmCopyCntr(ilLoop).iRegionNoToReassgMiss = 1
                    tmCopyCntr(ilUpper).iRegionNoToReassgMiss = 1       '3-16-07 (fix subscript out of range)
                End If
            End If
            tmCopyCntr(ilUpper).lChfCode = tmSdf.lChfCode
            tmCopyCntr(ilUpper).lFsfCode = tmSdf.lFsfCode
            tmCopyCntr(ilUpper).iVefCode = tmSdf.iVefCode
            tmCopyCntr(ilUpper).iAsgnVefCode = ilAsgnVefCode
            tmCopyCntr(ilUpper).iRdfCode = ilSaveRdfCode                '1-05-06
            tmCopyCntr(ilUpper).iLineNo = 0
            If ilIncludeLine Then
                tmCopyCntr(ilUpper).iLineNo = ilLineNo                 '2-28-07
                tmClfSrchKey.lChfCode = tmSdf.lChfCode
                tmClfSrchKey.iLine = tmSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

                '***** 7-10-08 First date of spot missing copy has been determine by the
                'by the missing copy counts: 1st time, save the date
                'If not showing missing copy, default to the start date of the flight or requested report start date, whichever is later

                If tmCopyCntr(ilUpper).iStartDate(0) = 0 And tmCopyCntr(ilUpper).iStartDate(1) = 0 Then
                    ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)
                    
                    '12-7-16 using ildate0 and ildate1 wipes out the looping air date for airing vehicles tests
'                    gPackDate slStartDate, ilDate0, ilDate1
'                    'default to start date of requested report
'                    tmCopyCntr(ilUpper).iStartDate(0) = ilDate0
'                    tmCopyCntr(ilUpper).iStartDate(1) = ilDate1

                    gPackDate slStartDate, ilInclLineDate0, ilInclLineDate1
                    'default to start date of requested report
                    tmCopyCntr(ilUpper).iStartDate(0) = ilInclLineDate0
                    tmCopyCntr(ilUpper).iStartDate(1) = ilInclLineDate1
                    
                    gUnpackDateLong tmCff.iStartDate(0), tmCff.iStartDate(1), llStartOfFlt
                    'use the start of flight or start date of requested report, whichever is later since a line cant start earlier than requested date
                    gUnpackDateLong ilMonRequest(0), ilMonRequest(1), llDate

                    If llStartOfFlt < llDate Then
                        llStartOfFlt = llDate
                    End If

                    If (tmCff.lChfCode = tmClf.lChfCode) And (tmCff.iClfLine = tmClf.iLine) Then
                        'ilDay = gWeekDayLong(llDate)
                        ilDay = gWeekDayLong(llStartOfFlt)
                        'start looping with the requested start date
                        For ilLoopDays = 0 To 7
                            If tmCff.iDay(ilDay) <> 0 Then      'valid first day of week
                                'gPackDateLong llStartOfFlt, tmCopyCntr(ilUpper).iStartDate(0), tmCopyCntr(ilUpper).iStartDate(1)
                                Exit For
                            Else
                                'llDate = llDate + 1
                                llStartOfFlt = llStartOfFlt + 1
                                'ilDay = gWeekDayLong(llDate)
                                ilDay = gWeekDayLong(llStartOfFlt)
                            End If
                        Next ilLoopDays
                    End If
                End If

                'OK to use line end date
                tmCopyCntr(ilUpper).iEndDate(0) = tmClf.iEndDate(0)
                tmCopyCntr(ilUpper).iEndDate(1) = tmClf.iEndDate(1)
            End If


            If tmSdf.iVefCode <> ilAsgnVefCode Or ilAiringVeh Then  'if the rot and airing vehicle arent the same, get the rot vehicle.
                'When testing airing vehicles, the spot (airing veh) will be the same as the selling
                'vehicle since its going back to do spots based on the selling vehicle.  The ilAiringVeh flag
                'to determine if airing veh to setup correct rotation & airing vehicles on the report
                'tmVefSrchKey.iCode = ilAsgnVefCode
                'ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                ilRet = gBinarySearchVef(ilAsgnVefCode)
                If ilRet >= 0 Then
                'If ilRet = BTRV_ERR_NONE Then
                    'tmCopyCntr(ilUpper).sVehName = Trim$(tlVef.sName)
                    tmCopyCntr(ilUpper).sVehName = Trim$(tgMVef(ilRet).sName)
                Else
                    tmCopyCntr(ilUpper).sVehName = slName
                End If
            Else
                tmCopyCntr(ilUpper).sVehName = slName
            End If
            tmCopyCntr(ilUpper).sAirVehName = slName
            tmCopyCntr(ilUpper).iLen = tmSdf.iLen
            tmCopyCntr(ilUpper).sLiveFlag = slLive          '11-16-05  live or recorded spot

            ReDim Preserve tmCopyCntr(0 To ilUpper + 1) As COPYCNTRSORT
            ilUpper = ilUpper + 1
        End If
    End If
End Sub

Private Sub mProcMakeRec(llUpper As Long, ilRptType As Integer, slSdfDate As String, slSdfTime As String, ilAsgnVefCode As Integer)
    Dim ilRet As Integer
    Dim slCntrNo As String
    Dim illoop As Integer
    Dim slCode As String
    Dim tlVef As VEF
    
    tmCopy(llUpper).tSdf = tmSdf
    If ilRptType = 0 Then   'By Date
        If tmVef.iCode <> tmSdf.iVefCode Then
            tmVefSrchKey.iCode = tmSdf.iVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        End If
        tmCopy(llUpper).sKey = tmVef.sName & slSdfDate & slSdfTime
    Else                        'by advt
        If tmSdf.lChfCode = 0 Then              'feed spot
            tmFSFSrchKey.lCode = tmSdf.lFsfCode
            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            slCntrNo = Trim$(tmFsf.sRefID)
            Do While Len(slCntrNo) < 10
                slCntrNo = "0" & slCntrNo
            Loop

        Else                                    'contract spot
            If tmChf.lCode <> tmSdf.lChfCode Then
                tmChfSrchKey.lCode = tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            End If
            slCntrNo = Trim$(str$(tmChf.lCntrNo))
            Do While Len(slCntrNo) < 10         'chged from 8 to 10 for Feed spot ref ID
                slCntrNo = "0" & slCntrNo
            Loop
        End If

        illoop = gBinarySearchAdf(tmSdf.iAdfCode)
        slCode = ""
        If illoop <> -1 Then
            ''slCode = Trim$(tgCommAdf(ilLoop).sName)
            'If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
            '    slCode = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
            'Else
                slCode = Trim$(tgCommAdf(illoop).sName)
            'End If
        End If
        tmCopy(llUpper).sKey = slCode & slCntrNo & slSdfDate & slSdfTime
    End If
    If ilAsgnVefCode <> tmVef.iCode Then
        tmVefSrchKey.iCode = ilAsgnVefCode
        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        tmCopy(llUpper).sVehName = tlVef.sName
    Else
        tmCopy(llUpper).sVehName = tmVef.sName
    End If
    ReDim Preserve tmCopy(0 To llUpper + 1) As COPYSORT
    llUpper = llUpper + 1
End Sub
