Attribute VB_Name = "RPTGENCB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptgencb.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: rptgencb.Bas
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
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
'Global tmsRec As LPOPREC
Type SMFINFO
    tSmf As SMF
    iAdfCode As Integer
    iSlfCode As Integer
    lChfCode As Long
    lCntrNo As Long
End Type
Dim tmSmfInfo() As SMFINFO
Type SPOTSALE
    sKey As String * 100    'Vehicle|sSofName|AdvtName|Date or 99999 if total
    iVefCode As Integer
    sVehName As String * 20
    sSOFName As String * 20
    sAdvtName As String * 30
    lCntrNo As Long
    lDate As Long
    sDate As String * 8
'    iCNoSpots As Long            'chged to long 5-12-99 from integer
    lCNoSpots As Long             '5-5-17 name change, designation of "L" for long
    sCGross As String * 12
    sCCommission As String * 12
    sCNet As String * 12
'    iTNoSpots As Integer
    lTNoSpots As Long           '5-5-17 changed to long
    sTGross As String * 12
    sTCommission As String * 12
    sTNet As String * 12
    iSofCode As Integer
End Type
Type VEHICLELLD
    iVefCode As Integer             'vehicle code
    iLLD(0 To 1) As Integer         'vehicles last log date
End Type
Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbf As CBF
Dim tmSort() As TYPESORT
Dim tmPLSdf() As SPOTTYPESORT
Dim tmPLPcf() As PCFTYPESORT
Dim tmSpotSOF() As SPOTTYPESORT
Dim tmSeqSortType() As SEQSORTTYPE
'Dim tmCopyCntr() As COPYCNTRSORTCB
'Dim tmCopy() As COPYSORTCB
'Dim tmSelAdvt() As Integer
Dim tmSelChf() As Long
Dim tmSelAgf() As Integer
Dim tmSelSlf() As Integer
Dim tmSelVef() As Integer                   'selective vehicle list when single advt selection (retrieval will be by advt, not vehicle.  So the vehicle isnt filtered out)
Dim imSpotSaleVefCode() As Integer
Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0            'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract line flight file handle
Dim tmCffSrchKey As CFFKEY0            'CFF record image
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF

Dim hmFsf As Integer            'Feed spot file handle
Dim tmFSFSrchKey As LONGKEY0     'FSF record image
Dim imFsfRecLen As Integer       'FSF record length
Dim tmFsf As FSF

Dim hmVsf As Integer            'Vehicle combo file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmVpf As Integer            'Vehicle options file handle
Dim tmVpf As VPF                'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer            'Agency name file handle
Dim tmAgf As AGF                'AGF record image
Dim tmAgfSrchKey As INTKEY0            'AGF record image
Dim imAgfRecLen As Integer        'AGF record length
Dim hmRdf As Integer            'Rate card program/time file handle
Dim tmRdf As RDF                'RDF record image
Dim tmRdfSrchKey As INTKEY0     'RDF record image
Dim imRdfRecLen As Integer      'RdF record length
Dim hmSmf As Integer            'MG and outside Times file handle
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey As SMFKEY0     'SMF record image
Dim tmSmfSrchKey2 As LONGKEY0
Dim tmSmfSrchKey5 As SMFKEY5    '4-5-10 speed up spot sales
Dim imSmfRecLen As Integer      'SMF record length
Dim hmStf As Integer            'MG and outside Times file handle
Dim tmStf As STF                'STF record image
Dim imStfRecLen As Integer      'STF record length
Dim tmAStf() As STF
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim tmSdfSrchKey4 As SDFKEY4
Dim tmSdfSrchKey7 As SDFKEY7
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF

Dim hmPcf As Integer            'Digital Line file handle
Dim tmPcf As PCF
Dim tmPcfSrchKey1 As PCFKEY1    'PCF record image
Dim tmPcfSrchKey2 As PCFKEY2    'PCF record image

'Short Title
'Copy rotation
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
'  Library calendar File
Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVLF As Integer            'Vehiclee file handle
Dim tmVlf As VLF                'VEF record image
Dim tmVlfSrchKey As VLFKEY0            'VEF record image
Dim imVlfRecLen As Integer        'VEF record length
Dim hmSlf As Integer            'Salesoerson file handle
Dim tmSlf As SLF                'SLF record image
Dim tmSlfSrchKey As INTKEY0            'SLF record image
Dim imSlfRecLen As Integer        'SLF record length
Dim hmSof As Integer            'Sales Office file handle
Dim tmSof As SOF                'SOF record image
Dim tmSofSrchKey As INTKEY0            'SOF record image
Dim imSofRecLen As Integer        'SOF record length
Dim hmSpf As Integer            'Site file handle
Dim tmSpf As SPF                'SPF record image
Dim imSpfRecLen As Integer        'SPF record length
Dim imUpdateCntrNo As Integer
Dim lmSpfRecPos As Long
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image

Dim hmAirSSF As Integer         '2nd SSF handle


Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim smOrdered() As String
Dim smAired() As String
Dim imUsingBarters As Integer

Dim tmVehLLD() As VEHICLELLD       'array of each vehicles last log dates
Dim tmSelAdf() As Integer    'array of advt to select matching network feeds by adv
Type PKGINFO
    iLineAsPkg As Integer
    iPkgVefCode As Integer
End Type
Dim tmPkgInfo() As PKGINFO '   Determine the Package line and Vehicle for a hidden line
'   Array tmPkgInfo contains the package lines along with its vehicle reference
Type SPOTSEQ
    lSdfCode As Long
    iSeq As Integer
End Type
Dim tmSpotSeq() As SPOTSEQ

Dim tmChfAdvtExt() As CHFADVTEXT

'5-9-17
Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF

Dim imProdPct() As Integer            '5-9-17  Participant share
Dim imMnfGroup() As Integer           '5-9-17  Participants
Dim imMnfSSCode() As Integer          '5-9-17  Particpant sales source
Dim lmStartDates() As Long       'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines

'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
Dim smClientName As String
Dim tmMnfSrchKey As INTKEY0
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmExport As Integer
Dim smLastCLF As String         'Prevent repeated Lookup of Same Contract Line
Dim smLineComment As String     'Line Comment from CXF
Dim smFormulaComment As String  'Digital Line Formula Commnet
Dim hmCxf As Integer
Dim tmCxf As CXF                'CXF Image
Dim imCxfRecLen As Integer      'CXF record length
Dim tmCxfSrchKey As LONGKEY0    'CXF key record image

Const LBONE = 1

'               Cycle thru the VSF for all vehicles used on the contract
'               For each vehicle, remove all bb in future
Public Sub mRemoveBBSpotSetup(slStartDate As String, slEndDate As String, tlChf As CHF)
    Dim ilVefCode As Integer
    Dim ilVsf As Integer
    Dim ilRet As Integer

    If tlChf.lVefCode < 0 Then
        tmVsfSrchKey.lCode = -tlChf.lVefCode
        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilVsf) > 0 Then
                    ilVefCode = tmVsf.iFSCode(ilVsf)
                    'do not game #s or line ids
                    ilRet = gRemoveBBSpots(hmSdf, ilVefCode, 0, slStartDate, slEndDate, tlChf.lCode, 0)
                End If
            Next ilVsf
            
            If tmVsf.lLkVsfCode <= 0 Then
                Exit Do
            End If
            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
End Sub

'           Filter out contract / spot types for HiLo Rate report
'           Created:  DH 6/9/10
Private Function mFilterHiLoRate(tlCntTypes As CNTTYPES) As Integer
    Dim ilFoundSpot As Integer

    ilFoundSpot = True
    If tmChf.iPctTrade = 100 Then
        ilFoundSpot = False
    End If
    If tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "H" Then        'ignore cancel and hidden
        ilFoundSpot = False
    End If
    If tmSdf.sSpotType = "X" Then           'fill
        ilFoundSpot = False
    End If
    
    '12-28-17 filter out the contract types
    mFilterCntTypes tmChf, tlCntTypes, ilFoundSpot
    mFilterHiLoRate = ilFoundSpot
End Function

'           Generate Hi-Lo Spot Rates by DP and vehicle
'           used to get the lowest rate for Politicals
'           Always exclude 100% trades, N/C or any kind ($0,fill, n/c, etc), hidden & cancelled spots
'           Selectivity for contract types (holds, orders, std, remnant, etc),
'           using DP overrides instead of DP name.
'           Created:  DH 6/9/10
Public Sub gGenHiLoRate()
    Dim slName As String
    Dim llRecNo As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llUpper As Long
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim ilMissedType As Integer
    Dim ilCostType As Integer
    Dim ilByOrderOrAir As Integer   '0=Order; 1=Aired
    Dim llContrCode As Long             'selective contr code
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer

    Dim llStartTime As Long             'start time filter entered
    Dim llEndTime As Long               'end time filter entered
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slCash As String                '8-24-01
    Dim slDollar As String              '8-24-01
    Dim ilLocal As Integer              'true to include contracts spots
    Dim ilFeed As Integer               'true to include network feed spots
    Dim ilIncludeCodes As Integer
    Dim ilFoundSpot As Integer          'for option by slsp, include /exclude
    Dim ilSlspOption As Integer         'true orfalse
    Dim ilNet As Integer                'true if net
    Dim ilCommPct As Integer            '% of agency comm
    Dim slSharePct As String
    Dim slAmount As String
    Dim ilPropPrice As Integer
    Dim llSpotSeq As Long
    Dim ilListIndex As Integer
    Dim ilBillType As Integer
    Dim ilSpotSelType As Integer
    Dim slDescription As String
    Dim ilLoop2 As Integer
    Dim ilLoop3 As Integer
    Dim ilShowOVDays As Integer
    Dim ilShowOVTimes As Integer
    Dim slOVStartTime As String
    Dim slOVEndTime As String
    Dim ilXMid As Integer
    Dim llrunningStartTime As Long
    Dim llrunningEndtime As Long
    Dim slTempDays As String
    Dim slTemp As String
    Dim slDay As String
    Dim slSpotCount As String
    Dim ilMajorSet As Integer
    Dim ilMinorSet As Integer
    Dim tlCntTypes As CNTTYPES
    
    ilListIndex = RptSelCb!lbcRptType.ListIndex     'report option

    ilPropPrice = False         'this report uses the actual spot price (never proposal price)

    slStartTime = "12M"        'start time
    llStartTime = gTimeToLong(slStartTime, False)
    slEndTime = "12M"           'end time
    llEndTime = gTimeToLong(slEndTime, True)

    'default all days selected, common code (mobtainsdf) used
    For illoop = 0 To 6
        RptSelCb!ckcSelC8(illoop) = vbChecked
    Next illoop
 
'    slEndDate = RptSelCb!edcSelCFrom.Text   'Start date
    slEndDate = RptSelCb!CSI_CalFrom.Text   'Start date, use csi calendar control
    illoop = Val(RptSelCb!edcSelCTo.Text)        '# days back
    llDate = gDateValue(slEndDate) - illoop + 1
    slStartDate = Format$(llDate, "m/d/yy")
    
    'mSetCostType ilCostType             'set bit pattern of types of spots to include
    ilCostType = 0
    ilCostType = ilCostType Or SPOT_CHARGE          'bit 0, only include charged spots
    ilCostType = ilCostType Or SPOT_MG              '10-20-10 include mgs (this include mg spot rates which will be filtered out later,
                                                    'and charged mgs
    gObtainVirtVehList

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
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        btrDestroy hmCHF
        btrDestroy hmClf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmCff
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmFsf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()

    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmRdf
        btrDestroy hmFsf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)

    llContrCode = 0                     'assume all contracts to be output (else contr code #)
    slStr = RptSelCb!edcSet3.Text  'see if selective contract entred
    If slStr <> "" Then
        llRecNo = Val(slStr)
        tmChfSrchKey1.lCntrNo = llRecNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (llRecNo = tmChf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")            'set the selective contr code only if no errors
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If llRecNo = tmChf.lCntrNo Then
            llContrCode = tmChf.lCode
        Else
            llContrCode = -1                    'get nothing, invalid contr #
        End If
    End If
    
    tlCntTypes.iHold = gSetCheck(RptSelCb!ckcSelC3(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelCb!ckcSelC3(1).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelCb!ckcSelC3(2).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelCb!ckcSelC3(3).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelCb!ckcSelC3(4).Value)
    tlCntTypes.iDR = gSetCheck(RptSelCb!ckcSelC3(5).Value)
    tlCntTypes.iPI = gSetCheck(RptSelCb!ckcSelC3(6).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelCb!ckcSelC3(7).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelCb!ckcSelC3(8).Value)


    'Setup mObtainSDF parameters
    ilMissedType = 1             'get missed, ignore cancel or hidden
    ilSpotSelType = 3            'scheduled & missed
    ilBillType = 3               'billed and unbilled
    ilByOrderOrAir = 1                  'by Aired
    ilLocal = True       'include local spots
    ilFeed = False       'exclude network (feed) vs local
    
    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0
    
    illoop = RptSelCb!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())

    'set generated date and time only once
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)

    For ilVehicle = 0 To RptSelCb!lbcSelection(6).ListCount - 1 Step 1
        If RptSelCb!lbcSelection(6).Selected(ilVehicle) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelCb!lbcCSVNameCode.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)

            ilVpfIndex = -1
            illoop = gBinarySearchVpf(ilVefCode)
            If illoop <> -1 Then
                ilVpfIndex = illoop
            End If
            
            ReDim tmPLSdf(0 To 0) As SPOTTYPESORT
            mObtainSdf ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilSpotSelType, ilBillType, True, ilMissedType, False, ilCostType, ilByOrderOrAir, False, llContrCode, ilLocal, ilFeed, ilPropPrice, ilListIndex, tlCntTypes
            'get the vehicle group once per vehicle
            'gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, ilLoop, tmGrf.iPerGenl(1)   'illoop = minor sort code (unused in report), genl(1) = major sort code
            gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, illoop, tmGrf.iPerGenl(0)   'illoop = minor sort code (unused in report), genl(1) = major sort code
   
            For llUpper = LBound(tmPLSdf) To UBound(tmPLSdf) - 1
                tmSdf = tmPLSdf(llUpper).tSdf
                ilFoundSpot = True
                If tmSdf.lChfCode <> tmChf.lCode Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    On Error GoTo gGenHiLoRateErr:
                    gBtrvErrorMsg ilRet, "gGenHiLoRate (btrGetEqual: CHF)", RptSelCb
                    On Error GoTo 0
                End If
                
                If tmSdf.lChfCode <> tmChf.lChfCode Or tmSdf.iLineNo <> tmClf.iLine Then
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    On Error GoTo gGenHiLoRateErr:
                    gBtrvErrorMsg ilRet, "gGenHiLoRate (btrGetGreaterEqual: CLF)", RptSelCb
                    On Error GoTo 0
                End If
                
                tmGrf.iAdfCode = tmSdf.iAdfCode
                tmGrf.lChfCode = tmChf.lCode           'Contr code
                tmGrf.iCode2 = tmSdf.iLen               'spot length
                tmGrf.iVefCode = tmSdf.iVefCode         'vehicle code
                'tmGrf.lDollars(1) = 0                   'init # spots
                'tmGrf.lDollars(2) = 0                   'gross rate
                tmGrf.lDollars(0) = 0                   'init # spots
                tmGrf.lDollars(1) = 0                   'gross rate
                'tmGrf.iPerGenl(2) = tmSdf.iLineNo       'schedule #
                'tmGrf.iPerGenl(3) = tmSdf.iGameNo       '4-12-06 show game info for vehicle
                tmGrf.iPerGenl(1) = tmSdf.iLineNo       'schedule #
                tmGrf.iPerGenl(2) = tmSdf.iGameNo       '4-12-06 show game info for vehicle
                
                ilFoundSpot = mFilterHiLoRate(tlCntTypes)         'filter contract types, trades, hidden, cancel spots, fills
                If Trim$(tmPLSdf(llUpper).sCostType) = "MG" Then    '10-20-10 spot rate is a mg, it $0, ignore
                    ilFoundSpot = False
                End If
                If ilFoundSpot Then
                    'setup the valid days of the flight
                    If tmPLSdf(llUpper).sDyWk = "W" Then
                        slTempDays = gDayNames(tmPLSdf(llUpper).iDay(), tmPLSdf(llUpper).sXDay(), 2, slStr)            'slstr not needed when returned
                        slStr = ""
                        For illoop = 1 To Len(slTempDays) Step 1
                            slDay = Mid$(slTempDays, illoop, 1)
                            If slDay <> " " And slDay <> "," Then
                                slStr = Trim$(slStr) & Trim$(slDay)
                            End If
                        Next illoop
                    Else
                        'Setup # spots/day
                        slStr = ""
                        For illoop = 0 To 6
                            slSpotCount = Trim$(str$(tmPLSdf(llUpper).iDay(illoop)))
                            Do While Len(slSpotCount) < 3
                                slSpotCount = " " & slSpotCount
                            Loop
                            slStr = slStr & " " & slSpotCount
                        Next illoop
                    End If
                    
                    'slStr contains days
                    tmRdfSrchKey.iCode = tmClf.iRdfCode
                    ilRet = btrGetEqual(hmRdf, tmRdf, Len(tmRdf), tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                    If ilRet <> BTRV_ERR_NONE Then
                        slDescription = "Missing DP"
                    End If

                    If RptSelCb!ckcSelC10(0).Value = vbChecked Then           'use daypart with overrides
                        ilShowOVDays = False
                        ilShowOVTimes = False
                        For ilLoop2 = 0 To 6 Step 1
                            'If tmRdf.sWkDays(7, ilLoop2 + 1) = "Y" Then             'is DP is a valid day
                            If tmRdf.sWkDays(6, ilLoop2) = "Y" Then             'is DP is a valid day

                                If tmPLSdf(llUpper).iDay(ilLoop2) = 0 Then          'is flight a valid day? 0=invalid day, >1 = daily # spots
                                    ilShowOVDays = True
                                    Exit For
                                Else
                                    ilShowOVDays = False
                                End If
                            End If
                        Next ilLoop2
                        'Times
                        If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slOVStartTime       '7-8-05
                            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slOVEndTime
                            ilShowOVTimes = True
                        Else
                            'Add times
                            ilXMid = False
                            'if there are multiple segments and it cross midnight, show the earliest start time and xmidnight end time;
                            'otherwise the first segments start and end times are shown
                            For ilLoop2 = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                                   If (tmRdf.iStartTime(0, ilLoop2) <> 1) Or (tmRdf.iStartTime(1, ilLoop2) <> 0) Then
                                    gUnpackTime tmRdf.iStartTime(0, ilLoop2), tmRdf.iStartTime(1, ilLoop2), "A", "1", slOVStartTime '7-8-05
                                    gUnpackTime tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), "A", "1", slOVEndTime
                                    gUnpackTimeLong tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), True, llrunningEndtime
                                    'If llrunningEndtime = 86400 And ilLoop2 <> 7 Then    'its 12M end of day, but not the first entry (which means only 1 time period, no multi-segments)
                                    If llrunningEndtime = 86400 And ilLoop2 <> UBound(tmRdf.iStartTime, 2) Then    'its 12M end of day, but not the first entry (which means only 1 time period, no multi-segments)
                                        ilXMid = True
                                    End If
                                    For ilLoop3 = ilLoop2 + 1 To UBound(tmRdf.iStartTime, 2)
                                        gUnpackTimeLong tmRdf.iStartTime(0, ilLoop3), tmRdf.iStartTime(1, ilLoop3), False, llrunningStartTime
                                         If llrunningStartTime = 0 And llrunningEndtime = 86400 Then
                                            If ilXMid Then
                                                gUnpackTime tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), "A", "1", slOVEndTime
                                                Exit For
                                            End If
                                        Else
                                            gUnpackTimeLong tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), True, llrunningEndtime
                                        End If
                                    Next ilLoop3
                                    Exit For
                                End If
                            Next ilLoop2
                        End If

                        If ilShowOVDays Or ilShowOVTimes Then
                            slDescription = RTrim$(slStr) & " " & Trim$(slOVStartTime) & "-" & Trim$(slOVEndTime)
                        Else
                            slDescription = Trim$(tmRdf.sName)
                        End If
                    Else            'use daypart name only; no overrides
                        slDescription = Trim$(tmRdf.sName)
                    End If

                    slStr = Trim$(tmPLSdf(llUpper).sCostType)       'check to make sure its not a zero charge
                    If InStr(slStr, ".") <> 0 Then          'its an amount, not N/c or any other text
                        'tmGrf.lDollars(2) = gStrDecToLong(slStr, 2) 'convert string decimal to long value
                        tmGrf.lDollars(1) = gStrDecToLong(slStr, 2) 'convert string decimal to long value
                    End If

                    'If tmGrf.lDollars(2) <> 0 Then      'make sure the 0 $ arent created
                    If tmGrf.lDollars(1) <> 0 Then      'make sure the 0 $ arent created
                        tmGrf.sGenDesc = slDescription      'dp (with or without overrides)
                        tmGrf.iRdfCode = tmRdf.iSortCode     'dp sort #
                        
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If
            Next llUpper
        End If                      'vehicle selected
    Next ilVehicle
    
    Screen.MousePointer = vbDefault
    Erase tmPLSdf
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmRdf)
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmAdf
    btrDestroy hmVef
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmCff
    btrDestroy hmGrf
    btrDestroy hmFsf
    btrDestroy hmRdf
    
    Exit Sub
gGenHiLoRateErr:
    On Error GoTo 0
    ilFoundSpot = False
    Resume Next
End Sub

'search the array that stores the sdf codes in order so that it can pull off the
'spot sequence within the break to maintain same order as spot screen on the
'spots by date and time report
Public Function mBinarySearchSDF(llSdfCode As Long, tlSpotSeq() As SEQSORTTYPE) As Long
    Dim llMiddle As Long
    Dim llMin As Long
    Dim llMax As Long
    llMin = LBound(tlSpotSeq)
    llMax = UBound(tlSpotSeq)
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llSdfCode = tlSpotSeq(llMiddle).lSdfCode Then
            'found the match
            mBinarySearchSDF = llMiddle
            Exit Function
        ElseIf llSdfCode < tlSpotSeq(llMiddle).lSdfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchSDF = -1
End Function

Public Sub mFindPkgReference()
    Dim ilClf As Integer
    For ilClf = LBound(tmPkgInfo) To UBound(tmPkgInfo) - 1
        tmCbf.iExtra2Byte = 0                   'pkg vehicle reference
        If tmPkgInfo(ilClf).iLineAsPkg = tmClf.iPkLineNo Then   '9-9-15 change the field to test (tmclf.ipklineNo isntead of tm.cbf.iextra4byte) find the matching pkg refernce from the hidden line (if applicable)
            tmCbf.iExtra2Byte = tmPkgInfo(ilClf).iPkgVefCode
            tmCbf.lExtra4Byte = tmPkgInfo(ilClf).iLineAsPkg
            Exit For
        End If
    Next ilClf
    Exit Sub
End Sub

'       mSpotSalesNetNet - if user requests net-net option, find the matching
'           Sales Source from the contracts primary slsp.  Then loop through
'           the vehicle table and find the first matching sales source and
'           calculate the net-net based on the first participant in the sales
'           source set
'
'           <input> slNet - Net value to calculate net-net from
'                   ilMatchSS - sales source from primary slsp of contract
'           <output> slNet - calculated net-net or orig net value if
Sub mSpotSalesNetNet(slNet As String, ilMatchSS As Integer)
    Dim slComm As String
    Dim ilLoopPart As Integer
    Dim ilFound As Integer
    Dim ilUse100pct As Integer          '5-9-17

    If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
        ilFound = False
        '5-9-17 determine owners share
        ilUse100pct = False          ' use 100% coming from recv if mnfgroup exists
        gInitPartGroupAndPcts tmSdf.iVefCode, ilMatchSS, 0, imMnfSSCode(), imMnfGroup(), imProdPct(), tmSdf.iDate(), tmPifKey(), tmPifPct(), ilUse100pct

        For ilLoopPart = 1 To UBound(imMnfSSCode)
            If imMnfSSCode(ilLoopPart) = ilMatchSS Then
                slComm = gIntToStrDec(imProdPct(ilLoopPart), 2)
                slNet = gDivStr(gMulStr(slNet, slComm), "100.00")
                ilFound = True
                Exit For
            End If
        Next ilLoopPart
        If Not ilFound Then
            slNet = ".00"
        End If
    End If
End Sub

'**************************************************************************************
'*                                                                                    *
'*      Procedure Name:gCntrDispRpt                                                   *
'*                                                                                    *
'*             Created:6/16/93       By:D. LeVine                                     *
'*            Modified:              By:linesurv                                      *
'*                                                                                    *
'*                                                                                    *
'*            Comments: Generate Contract discrepancy  report                         *
'*                                                                                    *
'*     DS 10/30/00 Converted to Crystal from Bridge                                   *
'*     DH 12/20/00 Spots scheduled outside ordered week previously not showing        *
'*                 Show M for N for combined MG spots                                 *
'*     DS 01/31/01 Fixed several problems in the report                               *
'*     DS 05/03/01 Corrected
'*     dh 06-10-02 Bypass contracts that are clusters in Spot Discrepancy
'*     dh 3-23-03 If fill/extra, determine from advt to show on invoice
'*     dh 1-19-04 change manner in which fill/extra are shown.  if spot price type
'           isnt a "-" or "+", then use advt for default; otherwise "-" is fill, "+" is extra
'       dh 7-19-04 Exclude network spots
'       dh 11-30-04 change access of smf to use key2 instead of key0 for speed
'       dh 12-1-06 subscript out of range error.  retrieve valid vehicle from tgmvef
'*******************************************************
'Sub gCntrDispRpt (ilDispOnly As Integer, ilPreview As Integer, slName As String, ilUrfCode As Integer)
Sub gCntrDispRpt(ilDispOnly As Integer)
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim llRecsRemaining As Long
    Dim ilDBEof As Integer
    'Dim ilDummy As Integer
    Dim ilLLRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim slNowDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llCntrIndex As Long
    Dim llStartIndex As Long
    Dim ilHeaderInit As Integer
    Dim ilLineInit As Integer
    'Dim ilLbcIndex As Integer
    Dim llLbcIndex As Long
    'Dim ilSdfIndex As Integer
    Dim llSdfIndex As Long
    Dim ilShowSpot As Integer
    Dim ilLineEOF As Integer
    Dim ilClf As Integer
    Dim slNameCode As String
    Dim ilCff As Integer
    Dim ilDay As Integer
    Dim ilSpotsPerWk As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim slStrTime As String
    Dim illoop As Integer
    Dim ilNewPage As Integer
    Dim llNoOrderedRows As Long
    Dim llNoAiredRows As Long
    Dim llNoOrderedRowsPrt As Long
    Dim llNoAiredRowsPrt As Long
    Dim llNoRowsToPrt As Long
    Dim llNoTimes As Long
    Dim llNoFlights As Long
    Dim ilMaxRowPerPage As Integer
    Dim llRowRemainingPerPage As Long
    Dim ilField As Integer
    Dim llRow As Long
    Dim ilNoLinesPerPage As Integer
    Dim llNoTotalLines As Long
    Dim slInvalid As String
    Dim slBBLength As String        '3-28-05
    Dim ilAnyOutput As Integer
    Dim ilWkCount As Integer
    Dim llWkDate As Long
    Dim slLnStartDate As String
    Dim slLnEndDate As String
    Dim slDateRange As String
    Dim ilSpotType As Integer   '0=Regular; 1= Missed of MG; 2=MG
    Dim ilNoDayToSun As Integer
    Dim il1stSpotInWk As Integer
    Dim tlVef As VEF
    ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
    'ReDim tmSdfExt(1 To 1) As SDFEXT
    ReDim tmSdfExt(0 To 0) As SDFEXT
    ReDim tmVehLLD(0 To 0) As VEHICLELLD
    Dim ilUpperVehLLD As Integer            'total vehicles built into tmVehLLD array
    Dim llVehLatestDate As Long                'all vehicles, latest log date
    Dim ilSlfCode As Integer                'slsp code is running report (only looks at own stuff)
    Dim sAdvName As String * 30
    Dim ilContFirstTime As Integer
    Dim ilCBS As Integer
    Dim ilPkgSpot As Integer            '12-20-00 true if spot found that is tied to package line (normally wouldnt occur,
                                        'except for MAI import of M for N spots
    Dim slShowOnInv As String * 1       '3-23-03
    Dim slVefType As String
    Dim ilRepSpot As Integer
    Dim ilSpotsFound As Integer
    Dim slLLDate As String          'last log as vehicle processed
    Dim llLLDate As Long
    Dim slSpotDate As String
    Dim ilLen As Integer
    
    If ilDispOnly Then
        If Not gSetFormula("TitleBanner", "'Spot Discrepancies'") Then 'Spot Placement
            Exit Sub
        End If
    Else
        If Not gSetFormula("TitleBanner", "'Spot Placements'") Then 'Spot Discrepancies
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tgChfCB)
    ReDim tgClfCB(0 To 0) As CLFLIST
    tgClfCB(0).iStatus = -1 'Not Used
    tgClfCB(0).lRecPos = 0
    tgClfCB(0).iFirstCff = -1
    ReDim tgCffCB(0 To 0) As CFFLIST
    tgCffCB(0).iStatus = -1 'Not Used
    tgCffCB(0).lRecPos = 0
    tgCffCB(0).iNextCff = -1
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
    imClfRecLen = Len(tgClfCB(0).ClfRec)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tgCffCB(0).CffRec)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVpf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVpf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVpfRecLen = Len(tmVpf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVpf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmVpf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVpf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVpf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmSpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSpf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSpf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imUpdateCntrNo = False
    imSpfRecLen = Len(tmSpf)
    hmCbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmSpf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCbf
        btrDestroy hmSpf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmRdf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)

    If RptSelCb!ckcAll.Value = vbChecked And ilDispOnly Then
        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            imUpdateCntrNo = True
            ilRet = btrGetPosition(hmSpf, lmSpfRecPos)
            slNowDate = Format$(Now, "m/d/yy")
            Do
                ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, lmSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                'tmSRec = tmSpf
                'ilRet = gGetByKeyForUpdate("Spf", hmSpf, tmSRec)
                'tmSpf = tmSRec
                gPackDate slNowDate, tmSpf.iDiscDateRun(0), tmSpf.iDiscDateRun(1)
                ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    End If
    'obtain Vehicle last log dates
    llVehLatestDate = 0
    ilUpperVehLLD = 0
    ilRet = btrGetFirst(hmVpf, tmVpf, imVpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmVehLLD(0 To ilUpperVehLLD) As VEHICLELLD
        tmVehLLD(ilUpperVehLLD).iVefCode = tmVpf.iVefKCode
        tmVehLLD(ilUpperVehLLD).iLLD(0) = tmVpf.iLLD(0)
        tmVehLLD(ilUpperVehLLD).iLLD(1) = tmVpf.iLLD(1)
        gUnpackDate tmVpf.iLLD(0), tmVpf.iLLD(1), slStr
        llDate = gDateValue(slStr)
        If llDate > llVehLatestDate Then            'find latest vehicles last log date
            llVehLatestDate = llDate
        End If
        ilRet = btrGetNext(hmVpf, tmVpf, imVpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilUpperVehLLD = ilUpperVehLLD + 1
    Loop
    If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then        'Guide or CSI
        ilSlfCode = 0
    Else
        ilSlfCode = tgUrf(0).iSlfCode
    End If
    If ilSlfCode > 0 Then                                   'slsp looking at own stuff, only allow to look at current stuff
                                                            'nothing past LLD of each vehicle
        'use vehicles LLD or user entered end date, whichever is earlier
        If llVehLatestDate < llEndDate Then                     'vehicle LLD is earlier than user's end date
            llEndDate = llVehLatestDate                         'alter the user entered end date
        End If
    End If
'    slDate = RptSelCb!edcSelCFrom.Text   'Start date
    slDate = RptSelCb!CSI_CalFrom.Text   'Start date, 9-11-19 use csi calendar control
   If slDate = "" Then
        slDate = "1/5/1970" 'Monday
        slDateRange = "Start"
    Else
        slDate = gObtainPrevMonday(slDate)
        slDateRange = slDate
    End If
    llStartDate = gDateValue(slDate)
'    slDate = RptSelCb!edcSelCFrom1.Text   'End date
    slDate = RptSelCb!CSI_CalTo.Text   'End date
    If (StrComp(slDate, "TFN", 1) = 0) Or (Len(slDate) = 0) Then
        llEndDate = gDateValue("12/29/2069")    'Sunday
        If slDateRange <> "Start" Then
            If ilSlfCode > 0 Then
                slDateRange = slDateRange & " to " & "Vehicle LLD"
            Else
                slDateRange = slDateRange & " to " & "End"
            End If
        Else
            If ilSlfCode > 0 Then
                slDateRange = "thru Vehicle LLD"
            Else
                slDateRange = "All Dates"
            End If
        End If
    Else
        slDate = gObtainNextSunday(slDate)
        llEndDate = gDateValue(slDate)
        If ilSlfCode > 0 Then
            slDateRange = slDateRange & " to " & "Vehicle LLD"
        Else
            slDateRange = slDateRange & " to " & slDate
        End If
    End If
    'Period in Broadcast weeks that the report spans
    tmCbf.sDemos = slDateRange
    tmAdf.iCode = 0
    tmVef.iCode = 0
    'slFileName = sgRptPath & slName & Chr$(0)
    'define variables for the load check (yes, L&L now checks the definition file!)
    'gCntrDisp
    'initiate printing
    ilAnyOutput = False
    'slLangText = "Printing..."
    'If ilPreview <> 0 Then
    '    ilLLRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, RptSelCB.hWnd, slLangText)
    '    ilDummy = LlPreviewSetTempPath(hdJob, sgRptSavePath)
    '    ilDummy = LlPreviewSetResolution(hdJob, 200)
    'Else
    '    ilLLRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_NORMAL Or LL_PRINT_MULTIPLE_JOBS, LL_BOXTYPE_BRIDGEMETER, RptSelCB.hWnd, slLangText)
    'End If
    DoEvents
    gAnyClustersDef 'any markets defined as clusters?

    If (ilLLRet = 0) Then
        'Compute Time of job by #Weeks*#Hours*#Zones*#Vehicles
        If RptSelCb!ckcAll.Value = vbChecked Then
            'llRecsRemaining = btrRecords(hmCHF)
            ilRet = mObtainCntrForDates(llStartDate, llEndDate)
            llRecsRemaining = UBound(tmChfAdvtExt)
        Else
            llRecsRemaining = 0
            For illoop = 0 To RptSelCb!lbcSelection(0).ListCount - 1 Step 1
                If RptSelCb!lbcSelection(0).Selected(illoop) Then
                    llRecsRemaining = llRecsRemaining + 1
                End If
            Next illoop
        End If
        llRecNo = 0
        ilErrorFlag = 0
        ilDBEof = False
        'slPrinter = LlVBPrintGetPrinter(hdJob)
        'slPort = LlVBPrintGetPort(hdJob)
        If RptSelCb!ckcAll.Value = vbChecked Then
            llCntrIndex = LBound(tmChfAdvtExt)
        Else
            llCntrIndex = 0
        End If
        'ilDummy = LLDefineVariableExt(hdJob, "Logo", sgLogoPath & "RptLogo.Bmp", LL_DRAWING, "")
        'ilDummy = LLDefineVariableExtHandle(hdJob, "CSILogo", Traffic!imcCSILogo, LL_DRAWING_HBITMAP, "")
        If ilDispOnly Then
            'ilDummy = LLDefineVariableExt(hdJob, "ReportName", "Spot Discrepancies", LL_TEXT, "")
        Else
            'ilDummy = LLDefineVariableExt(hdJob, "ReportName", "Spot Placements", LL_TEXT, "")
        End If
            'ilDummy = LLDefineVariableExt(hdJob, "DateRange", slDateRange, LL_TEXT, "")
        Do
            Screen.MousePointer = vbHourglass
            llStartIndex = llCntrIndex
            ilRet = mReadChfRec(llCntrIndex, ilDispOnly, llStartDate, llEndDate, tmSdfExtSort(), tmSdfExt())
            If Not ilRet Then
                Exit Do
            End If
            tmChf = tgChfCB
                tmCbf.iGenDate(0) = igNowDate(0)
                tmCbf.iGenDate(1) = igNowDate(1)
                tmCbf.lGenTime = lgNowTime
                tmCbf.sProduct = tmChf.sProduct
                tmCbf.iAdfCode = tmChf.iAdfCode
                tmCbf.lContrNo = tmChf.lCntrNo
                ilNewPage = True
                ilHeaderInit = False
                'Number of lines in the contract
                'One Loop for each contract line
                ReDim tmPkgInfo(0 To 0) As PKGINFO
                tmCbf.iExtra2Byte = 0       '9-9-15 init pkg vehicle references  (pkg vehicle code)
                tmCbf.lExtra4Byte = 0       'init pkg vehicle line reference
                'build array of the package lines along with their package vehicles so that the hidden lines can show the reference
                For ilClf = LBound(tgClfCB) To UBound(tgClfCB) - 1
                    tmClf = tgClfCB(ilClf).ClfRec
                    If tmClf.sType = "O" Or tmClf.sType = "A" Or tmClf.sType = "E" Then 'find any kind of pkg line
                        tmPkgInfo(UBound(tmPkgInfo)).iLineAsPkg = tmClf.iLine
                        tmPkgInfo(UBound(tmPkgInfo)).iPkgVefCode = tmClf.iVefCode
                        ReDim Preserve tmPkgInfo(0 To UBound(tmPkgInfo) + 1) As PKGINFO
                    End If
                Next ilClf
                For ilClf = LBound(tgClfCB) To UBound(tgClfCB) - 1 Step 1
                    Screen.MousePointer = vbHourglass
                    tmClf = tgClfCB(ilClf).ClfRec

                    'ignore overides for m for n customers like MAI
                    'For ilLoop = 0 To UBound(tgVpf) Step 1
                    '    If tmClf.iVefCode = tgVpf(ilLoop).iVefKCode Then
                        slVefType = ""
                        illoop = gBinarySearchVpf(tmClf.iVefCode)
                        If illoop <> -1 Then
                           If tgVpf(illoop).sGMedium = "S" Then
                               tmClf.iStartTime(0) = 1
                               tmClf.iStartTime(1) = 0
                            End If
                            
                            'need to ignore  BB from days in future, so get the last log date for this vehicle
                            gUnpackDate tgVpf(illoop).iLLD(0), tgVpf(illoop).iLLD(1), slLLDate
                            If slLLDate = "" Then
                                slLLDate = Format(Now, "m/d/yy")
                            Else
                                If gDateValue(slLLDate) < gDateValue(Format(Now, "m/d/yy")) Then
                                    slLLDate = Format(Now, "m/d/yy")
                                End If
                            End If
                            slLLDate = gIncOneDay(slLLDate)
                            llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater

                            illoop = gBinarySearchVef(tmClf.iVefCode)       '12-01-06, determine the vehicle type
                            If illoop <> -1 Then
                                slVefType = tgMVef(illoop).sType
                            Else
                                slVefType = "C"         'fake out to be conventional
                            End If
                    '       Exit For
                        Else
                            slLLDate = gIncOneDay(Format(Now, "m/d/yy"))
                            llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater
                        End If
                    'Next ilLoop

                    '12-20-00 if this is a package line, determine if any spots are tied to it and then print out the line;
                    'otherwise bypass the line.  Spots are not normally tied to package lines.
                    ilPkgSpot = False

                    If (tmClf.sType = "O" Or tmClf.sType = "A" Or tmClf.sType = "E") Then              'ignore package lines
                        'For ilLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                        For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                            'ilSdfIndex = tmSdfExtSort(ilLbcIndex).iSdfExtIndex
                            llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex
                            'If Abs(tmSdfExt(ilSdfIndex).iLineNo) = tmClf.iLine Then
                            If Abs(tmSdfExt(llSdfIndex).iLineNo) = tmClf.iLine Then
                                ilPkgSpot = True
                                Exit For
                            End If
                        Next llLbcIndex
                    End If
                    ilRepSpot = False

                    If slVefType = "R" Then              'ignore Rep lines
                        'For ilLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                        For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                            'ilSdfIndex = tmSdfExtSort(ilLbcIndex).iSdfExtIndex
                            llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex
                            'If Abs(tmSdfExt(ilSdfIndex).iLineNo) = tmClf.iLine Then
                            If Abs(tmSdfExt(llSdfIndex).iLineNo) = tmClf.iLine Then
                                ilRepSpot = True
                                Exit For
                            End If
                        Next llLbcIndex
                    End If
                    If (slVefType <> "R") And (tmClf.sType <> "O" And tmClf.sType <> "A" And tmClf.sType <> "E") Or (ilPkgSpot) Or (ilRepSpot) Then    '12-20-00 ignore package lines only if there wasnt a line tied to a package
                        'determine last log date of the vehicle it's working on
                        If ilSlfCode > 0 Then
                            For illoop = 0 To UBound(tmVehLLD) - 1 Step 1
                                If tmVehLLD(illoop).iVefCode = tmClf.iVefCode Then
                                    gUnpackDate tmVehLLD(illoop).iLLD(0), tmVehLLD(illoop).iLLD(1), slStr
                                    llEndDate = gDateValue(slStr)
                                    If slStr = "" Then
                                        llEndDate = llVehLatestDate
                                    End If
                                    Exit For
                                End If
                            Next illoop
                        End If
                        'Check if any discrepancies
                        If ilDispOnly Then
                            'ilLLRet = LlPrintSetBoxText(hdJob, "Checking" & Str$(tgChfCB.lCntrNo) & " Line" & Str$(tmClf.iLine), (100# * llRecNo / llRecsRemaining))
                            DoEvents
                            'If ilLLRet <> 0 Then
                            '    ilErrorFlag = ilLLRet
                            '    Exit Do
                            'End If
                            ilRet = mCntrSchdSpotChk(ilDispOnly, ilClf, llStartDate, llEndDate, tmSdfExtSort(), tmSdfExt())
                        Else
                            'ilLLRet = LlPrintSetBoxText(hdJob, "Gathering" & Str$(tgChfCB.lCntrNo) & " Line" & Str$(tmClf.iLine), (100# * llRecNo / llRecsRemaining))
                            'DoEvents
                            'If ilLLRet <> 0 Then
                            '    ilErrorFlag = ilLLRet
                            '    Exit Do
                            'End If
                            'Set status flags
                            ilRet = mCntrSchdSpotChk(ilDispOnly, ilClf, llStartDate, llEndDate, tmSdfExtSort(), tmSdfExt())
                        End If
                        DoEvents

                        If ilRet = False Then

                            ilAnyOutput = True
                            'Build output arrays
                            'ReDim smOrdered(1 To 8, 1 To 1) As String
                            ReDim smOrdered(0 To 8, 0 To 1) As String   'Index zero ignored
                            smOrdered(1, 1) = Trim$(str$(tmClf.iLine))
                            tmCbf.lLineNo = tmClf.iLine
                            'ReDim smAired(1 To 3, 1 To 1) As String
                            ReDim smAired(0 To 3, 0 To 1) As String 'Index zero ignored
                            'Vehcile Name
                            If tmClf.iVefCode <> tmVef.iCode Then
                                tmVefSrchKey.iCode = tmClf.iVefCode
                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                If ilRet <> BTRV_ERR_NONE Then
                                    tmVef.sName = "Missing"
                                End If
                            End If
                            'Vehicle code
                            'tmCbf.ivefCode = tmVef.icode
                            'ok
                            tmCbf.sLineSurvey = Trim$(tmVef.sName)
                            smOrdered(2, 1) = Trim$(tmVef.sName)
                            'Times
                            If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
                                gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                                gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                                smOrdered(3, 1) = slStartTime & "-" & slEndTime
                                tmCbf.sDysTms = slStartTime & "-" & slEndTime
                            Else
                                'Add times
                                slStrTime = ""
                                llNoTimes = 0
                                For illoop = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                                    If (tmRdf.iStartTime(0, illoop) <> 1) Or (tmRdf.iStartTime(1, illoop) <> 0) Then
                                        gUnpackTime tmRdf.iStartTime(0, illoop), tmRdf.iStartTime(1, illoop), "A", "1", slStartTime
                                        gUnpackTime tmRdf.iEndTime(0, illoop), tmRdf.iEndTime(1, illoop), "A", "1", slEndTime
                                        llNoTimes = llNoTimes + 1
                                        If llNoTimes > UBound(smOrdered, 2) Then
                                            'ReDim Preserve smOrdered(1 To 8, 1 To llNoTimes) As String
                                            ReDim Preserve smOrdered(0 To 8, 0 To llNoTimes) As String
                                        End If
                                        smOrdered(3, llNoTimes) = Trim$(tmRdf.sName) & " " & slStartTime & "-" & slEndTime
                                        tmCbf.sDysTms = Trim$(tmRdf.sName) & " " & slStartTime & "-" & slEndTime
                                    End If
                                    'Ordered Times

                                Next illoop

                            End If
                            slStr = slStr & Chr$(0)
                            'Aired Times
                            'tmCbf.sDysTms = slStr
                            'ilDummy = LLDefineFieldExt(hdJob, "Times", slStr, LL_TEXT, "")
                            'Length
                            smOrdered(4, 1) = Trim$(str$(tmClf.iLen))
                            'Spot Len
                            tmCbf.iLen = tmClf.iLen
                            slLnStartDate = ""
                            slLnEndDate = ""
                            ilCff = tgClfCB(ilClf).iFirstCff
                            Do While ilCff <> -1
                                If slLnStartDate = "" Then
                                    gUnpackDate tgCffCB(ilCff).CffRec.iStartDate(0), tgCffCB(ilCff).CffRec.iStartDate(1), slLnStartDate
                                    gUnpackDate tgCffCB(ilCff).CffRec.iEndDate(0), tgCffCB(ilCff).CffRec.iEndDate(1), slLnEndDate
                                Else
                                    gUnpackDate tgCffCB(ilCff).CffRec.iEndDate(0), tgCffCB(ilCff).CffRec.iEndDate(1), slLnEndDate
                                End If
                                tmCbf.iStartQtr(0) = tgCffCB(ilCff).CffRec.iStartDate(0)
                                tmCbf.iStartQtr(1) = tgCffCB(ilCff).CffRec.iStartDate(1)
                                ilCff = tgCffCB(ilCff).iNextCff
                            Loop
                            'Dates, days and number of spots per wk
                            llNoFlights = 0
                            llNoAiredRows = 0
                            ilCff = tgClfCB(ilClf).iFirstCff
                            Do While ilCff <> -1
                                'llNoFlights = llNoFlights + 1
                                'If llNoFlights > UBound(smOrdered, 2) Then
                                '    ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                'End If
                                's
                                gUnpackDate tgCffCB(ilCff).CffRec.iStartDate(0), tgCffCB(ilCff).CffRec.iStartDate(1), slSFlightDate
                                gUnpackDate tgCffCB(ilCff).CffRec.iEndDate(0), tgCffCB(ilCff).CffRec.iEndDate(1), slEFlightDate
                                llWkDate = gDateValue(slSFlightDate)
                                ilNoDayToSun = 6 - gWeekDayStr(slSFlightDate)
                                il1stSpotInWk = True
                                ilCBS = False
                                Do
                                    If ((llWkDate + ilNoDayToSun >= llStartDate) And (llWkDate <= llEndDate)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                        llNoFlights = llNoFlights + 1
                                        ilSpotsFound = False     '12-26-07 need to show ordered info if no airing spots found
                                        If llNoFlights > UBound(smOrdered, 2) Then
                                            'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                            ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                        End If
                                        If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                            smOrdered(5, llNoFlights) = "CBS"
                                            tmCbf.sSortField1 = "CBS"
                                            ilCBS = True
                                        Else
                                            If llWkDate < gDateValue(slLnStartDate) Then
                                                smOrdered(5, llNoFlights) = slLnStartDate
                                                tmCbf.lDrfCode = gDateValue(slLnStartDate)
                                            Else
                                                smOrdered(5, llNoFlights) = Format$(llWkDate, "m/d/yy")
                                                tmCbf.lDrfCode = llWkDate

                                            End If
                                            If llWkDate + ilNoDayToSun > gDateValue(slLnEndDate) Then
                                                smOrdered(5, llNoFlights) = slLnEndDate & "-" & Format$(llWkDate + ilNoDayToSun, "m/d/yy")
                                                'Ordered Dates
                                                tmCbf.sSortField1 = slLnEndDate & "-" & Format$(llWkDate + ilNoDayToSun, "m/d/yy")
                                            Else
                                                smOrdered(5, llNoFlights) = smOrdered(5, llNoFlights) & "-" & Format$(llWkDate + ilNoDayToSun, "m/d/yy")
                                                'Ordered Dates
                                                tmCbf.sSortField1 = smOrdered(5, llNoFlights)
                                            End If

                                        End If
                                        slStr = ""
                                        If (tgCffCB(ilCff).CffRec.iSpotsWk > 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk > 0 Or (tgCffCB(ilCff).CffRec.sDyWk = "W")) Then
                                            'slStr = mDayNames(tgCffCB(ilCff).CffRec.iDay(), tgCffCB(ilCff).CffRec.sXDay(), 2)
                                            slStr = gDayNames(tgCffCB(ilCff).CffRec.iDay(), tgCffCB(ilCff).CffRec.sXDay(), 2, slInvalid)
                                            smOrdered(6, llNoFlights) = slStr
                                            'Ordered Days
                                            tmCbf.sSortField2 = slStr
                                            smOrdered(7, llNoFlights) = Trim$(str$(tgCffCB(ilCff).CffRec.iSpotsWk + tgCffCB(ilCff).CffRec.iXSpotsWk))
                                            'Number Per Week
                                            tmCbf.sResort = Trim$(str$(tgCffCB(ilCff).CffRec.iSpotsWk + tgCffCB(ilCff).CffRec.iXSpotsWk))
                                        Else
                                            ilSpotsPerWk = 0
                                            slStr = ""
                                            For ilDay = 0 To 6 Step 1
                                                ilSpotsPerWk = ilSpotsPerWk + tgCffCB(ilCff).CffRec.iDay(ilDay)
                                                slStr = slStr & str$(tgCffCB(ilCff).CffRec.iDay(ilDay))
                                            Next ilDay
                                            smOrdered(6, llNoFlights) = slStr
                                            'Ordered Days
                                            tmCbf.sSortField2 = slStr
                                            smOrdered(7, llNoFlights) = Trim$(str$(ilSpotsPerWk))
                                            'Number Per Week
                                            tmCbf.sResort = Trim$(str$(ilSpotsPerWk))
                                        End If
                                        'Rate
                                        Select Case tgCffCB(ilCff).CffRec.sPriceType
                                            Case "T"    'True
                                                 'gPDNToStr tmClf.sActPrice, 2, smOrdered(8, 1)
                                                smOrdered(8, llNoFlights) = gLongToStrDec(tgCffCB(ilCff).CffRec.lActPrice, 2)
                                                'Ordered Price
                                                tmCbf.lGrImp = tgCffCB(ilCff).CffRec.lActPrice
                                            Case "N"    'No Charge
                                                smOrdered(8, llNoFlights) = "N/C"
                                                'Ordered Price
                                                tmCbf.lGrImp = -1
                                            Case "M"    'MG Line
                                                smOrdered(8, llNoFlights) = "MG"
                                                'Ordered Price
                                                tmCbf.lGrImp = -2
                                            Case "B"    'Bonus
                                                smOrdered(8, llNoFlights) = "Bonus"
                                                'Ordered Price
                                                tmCbf.lGrImp = -3
                                            Case "S"    'Spinoff
                                                smOrdered(8, llNoFlights) = "Spinoff"
                                                'Ordered Price
                                                tmCbf.lGrImp = -4
                                            Case "P"    'Package
                                                'gPDNToStr tmClf.sActPrice, 2, smOrdered(8, llNoFlights)
                                                smOrdered(8, llNoFlights) = gLongToStrDec(tgCffCB(ilCff).CffRec.lActPrice, 2)
                                                'Ordered Price
                                                tmCbf.lGrImp = tgCffCB(ilCff).CffRec.lActPrice
                                            Case "R"    'Recapturable
                                                smOrdered(8, llNoFlights) = "Recapturable"
                                                'Ordered Price
                                                tmCbf.lGrImp = -5
                                            Case "A"    'ADU
                                                smOrdered(8, llNoFlights) = "ADU"
                                                'Ordered Price
                                                tmCbf.lGrImp = -6
                                        End Select
                                        'Build Aired Array
                                        ilWkCount = 0
                                        llLbcIndex = 0
                                        ilSpotType = 0
                                        If (Not ilPkgSpot) And (Not ilRepSpot) Then       '12-20-00
                                            Do While llLbcIndex <= UBound(tmSdfExtSort) - 1 'RptSelCb!lbcSort.ListCount - 1

                                                slNameCode = tmSdfExtSort(llLbcIndex).sKey  'RptSelCb!lbcSort.List(ilLbcIndex)
                                                llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex  'Val(slCode)
                                                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                '************
                                                If Not ilDispOnly Then
                                                  '5/3/01 braketed out line below
                                                    'ilRet = gParseItem(slNameCode, 2, "|", tmCbf.sLineSurvey)
                                                End If
                                                ilShowSpot = False
                                                If Abs(tmSdfExt(llSdfIndex).iLineNo) = tgClfCB(ilClf).ClfRec.iLine Then
                                                    'Only show spots within date span
                                                    '&H1000 = Missed shown; &H2000 = MG Shown; &H3000 = Missed & MG Shown
                                                    tmSmf.iOrigSchVef = tmSdfExt(llSdfIndex).iVefCode
                                                    If (tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Then
                                                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tmSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        'Obtain original dates

                                                        '11-30-04 change access of smf to use key2 instead of key0 for speed
                                                        'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                                                        'tmSmfSrchKey.iLineNo = tmSdf.iLineNo

                                                        'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                                                        'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                                                        'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                        tmSmfSrchKey2.lCode = tmSdf.lCode
                                                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                                            If tmSmf.lSdfCode = tmSdf.lCode Then
                                                                ilRet = ilRet
                                                                Exit Do
                                                            End If
                                                            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                        Loop

                                                        If (tmSdfExt(llSdfIndex).sSchStatus = "O") And (tmSdfExt(llSdfIndex).sSpotType = "X") Then   'if outside on bonus, bypass printing the missed portion by forcing the status to a 2
                                                            tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H1000
                                                        End If

                                                        If (tmSdfExt(llSdfIndex).iStatus And &H1000) = &H1000 Then
                                                            ilSpotType = 2
                                                        End If
                                                        If (ilSpotType = 0) And (tmSdfExt(llSdfIndex).lMdDate >= llStartDate) And (tmSdfExt(llSdfIndex).lMdDate <= llEndDate) Then
                                                            ilSpotType = 1
                                                            If ((tmSdfExt(llSdfIndex).lMdDate >= llWkDate) And (tmSdfExt(llSdfIndex).lMdDate <= llWkDate + ilNoDayToSun)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                                                ilShowSpot = True
                                                                tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H1000
                                                            Else
                                                                ilSpotType = 2
                                                                gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                                If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                                                    If ((gDateValue(slDate) >= llWkDate) And (gDateValue(slDate) <= llWkDate + ilNoDayToSun)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                                                        ilShowSpot = True
                                                                        tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H2000
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            ilSpotType = 2
                                                            gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                            If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                                                If ((gDateValue(slDate) >= llWkDate) And (gDateValue(slDate) <= llWkDate + ilNoDayToSun)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                                                    ilShowSpot = True
                                                                    If (tmSdfExt(llSdfIndex).lMdDate >= llStartDate) And (tmSdfExt(llSdfIndex).lMdDate <= llEndDate) Then
                                                                        tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H2000
                                                                    Else
                                                                        tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H3000
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        ilSpotType = 0
                                                        If ilDispOnly Then                  'discrepancy only,ignore BB spots
                                                            'If tmSdfExt(llSdfIndex).sSpotType <> "O" And tmSdfExt(llSdfIndex).sSpotType <> "C" Then
                                                                gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                                If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                                                    If ((gDateValue(slDate) >= llWkDate) And (gDateValue(slDate) <= llWkDate + ilNoDayToSun)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                                                        ilShowSpot = True
                                                                        tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H3000
                                                                    End If
                                                                End If
                                                            'End If
                                                        Else
                                                            gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                            If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                                                If ((gDateValue(slDate) >= llWkDate) And (gDateValue(slDate) <= llWkDate + ilNoDayToSun)) Or (gDateValue(slEFlightDate) < gDateValue(slSFlightDate)) Then
                                                                    ilShowSpot = True
                                                                    tmSdfExt(llSdfIndex).iStatus = tmSdfExt(llSdfIndex).iStatus Or &H3000
                                                                End If
                                                            End If
                                                        End If

                                                    End If
                                                    'If tmSdfExt(ilSdfIndex).iStatus > 0 Then
                                                    '    ilShowSpot = True
                                                    'End If
                                                End If
                                                '10/24/05:  If idscrepancy, bypass Fill and Billboard spots
                                                If (ilShowSpot) And (ilDispOnly) Then
                                                    If (tmSdfExt(llSdfIndex).sSpotType = "X") Then
                                                        ilShowSpot = False
                                                    End If
                                                    If Not ilCBS Then
                                                        If (tmSdfExt(llSdfIndex).sSpotType = "O") Or (tmSdfExt(llSdfIndex).sSpotType = "C") Then
                                                            ilShowSpot = False
                                                        End If
                                                    End If
                                                End If

                                                '6-26-06 option to show fill spots on Spot Placement
                                                If Not ilDispOnly Then      'spot placement vs Discrep only
                                                    'spot placement, if "Fill" should it be shown based on user option
                                                    If tmSdfExt(llSdfIndex).sSpotType = "X" And RptSelCb!ckcSelC3(1).Value = vbUnchecked Then
                                                        ilShowSpot = False
                                                    End If
                                                    
                                                    '5-8-14 ignore open/close bb in future
                                                    'Test if Open or Close BB, ignore if in the future
                                                    If tmSdfExt(llSdfIndex).sSpotType = "O" Or tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slSpotDate
                                                        If gDateValue(slSpotDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
                                                            ilShowSpot = False
                                                        End If
                                                    End If
                                                                                        
                                                Else            '11-23-11 if discrep, dont show the obb and cbb
                                                    If tmSdfExt(llSdfIndex).sSpotType = "O" Or tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                                        ilShowSpot = False
                                                    End If
                                                End If

                                                If ilShowSpot Then
                                                    ilSpotsFound = True      '12-26-07
                                                    If ((tmSdfExt(llSdfIndex).iStatus And &H3000) = &H3000) Or (ilSpotType = 0) Then
                                                        tmSdfExt(llSdfIndex).iLineNo = 0
                                                    End If
                                                    ilWkCount = ilWkCount + 1
                                                    If tmVef.sType <> "V" Then
                                                        llNoAiredRows = llNoAiredRows + 1
                                                    Else
                                                        If il1stSpotInWk Then
                                                            llNoAiredRows = llNoAiredRows + 2
                                                        Else
                                                            llNoAiredRows = llNoAiredRows + 1
                                                        End If
                                                    End If
                                                    il1stSpotInWk = False
                                                    If llNoAiredRows > UBound(smAired, 2) Then
                                                        'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                        ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                                    End If
                                                    If llNoAiredRows > llNoFlights Then
                                                        llNoFlights = llNoAiredRows
                                                        If llNoFlights > UBound(smOrdered, 2) Then
                                                            'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                            ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                                        End If
                                                    End If
                                                    If tmVef.sType = "V" Then
                                                        If (tmSdfExt(llSdfIndex).iVefCode <> tmSmf.iOrigSchVef) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O")) Then
                                                            tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                                                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                tlVef.sName = "Missing"
                                                            End If
                                                            slStr = Trim$(tlVef.sName)
                                                        Else
                                                            ilRet = gParseItem(slNameCode, 2, "|", slStr)
                                                        End If
                                                        tmCbf.sLineSurvey = Trim$(slStr)
                                                        smOrdered(2, llNoAiredRows) = "  " & Trim$(slStr)
                                                    End If
                                                    If (ilSpotType = 0) Or (ilSpotType = 2) Then
                                                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                        tmCbf.iPropOrdDate(0) = tmSdfExt(llSdfIndex).iDate(0)
                                                        tmCbf.iPropOrdDate(1) = tmSdfExt(llSdfIndex).iDate(1)
                                                    Else
                                                        'Orig MG date here
                                                        slDate = Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                        gPackDate slDate, tmCbf.iPropOrdDate(0), tmCbf.iPropOrdDate(1)
                                                    End If
                                                    llDate = gDateValue(slDate)
                                                    smAired(1, llNoAiredRows) = Format$(llDate, "ddd") & ", " & slDate
                                                    'Aired Day, Date
                                                    'ilDummy = LLDefineFieldExt(hdJob, "AirDay", slStr, LL_TEXT, "")
                                                    If ((ilSpotType = 0) And (tmSdfExt(llSdfIndex).sSchStatus = "S")) Or (ilSpotType = 2) Then
                                                        gUnpackTime tmSdfExt(llSdfIndex).iTime(0), tmSdfExt(llSdfIndex).iTime(1), "A", "1", slTime
                                                        smAired(2, llNoAiredRows) = slTime
                                                        tmCbf.sBuyer = slTime
                                                    End If
                                                    'Aired Time
                                                    'tmCbf.sBuyer = slTime
                                                    'ilDummy = LLDefineFieldExt(hdJob, "AirTime", slTime, LL_TEXT, "")
                                                    Select Case tmSdfExt(llSdfIndex).sSchStatus
                                                        Case "S"
                                                            If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                                slStr = "Scheduled"
                                                            Else                    'fill or extra?

                                                                '1-19-04 change way in which fill/extra are shown
                                                                slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                                If slShowOnInv = "N" Then
                                                                'If tmSdfExt(ilSdfIndex).sPriceType = "N" Then      'fill?
                                                                    slStr = "-Schd Fill"
                                                                Else
                                                                    'slStr = "Schd Extra"
                                                                    slStr = "+Schd Fill"
                                                                End If
                                                            End If
                                                        Case "G"
                                                            If ilSpotType = 2 Then
                                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                                    If tmSmf.lMtfCode > 0 Then   '12-20-00
                                                                        slStr = "MG, M for N"
                                                                    Else
                                                                        slStr = "MG, Missed " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                                    End If
                                                                Else            'fill or extra?
                                                                    '1-19-04 change way in which fill/extra are shown
                                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                                    'If tmSdfExt(ilSdfIndex).sPriceType = "N" Then
                                                                    If slShowOnInv = "N" Then
                                                                        slStr = "-Fill MG"   ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                    Else
                                                                       ' slStr = "Extra MG"  ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                        slStr = "+Fill MG"  ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                    End If
                                                                End If
                                                                'slStr = "MG, Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                            Else
                                                                gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                                    If tmSmf.lMtfCode > 0 Then   '12-20-00
                                                                        slStr = "MG, M for N"
                                                                    Else
                                                                        slStr = "Missed, MG " & slDate
                                                                    End If
                                                                Else
                                                                    '1-19-04 change way in which fill/extra are shown
                                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                                    If slShowOnInv = "N" Then
                                                                    'If tmSdfExt(ilSdfIndex).sPriceType = "N" Then       'fill or extra?
                                                                        slStr = " -Fill Missed, MG " & slDate
                                                                    Else
                                                                        'slStr = " Extra Missed, MG " & slDate
                                                                        slStr = " +Fill Missed, MG " & slDate
                                                                    End If
                                                                End If
                                                                'slStr = "Missed, MG " & slDate
                                                            End If
                                                        Case "O"
                                                            If ilSpotType = 2 Then
                                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                                    slStr = "Outside, Missed " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                                Else
                                                                    '1-19-04 change way in which fill/extra are shown
                                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                                    If slShowOnInv = "N" Then
                                                                    'If tmSdfExt(ilSdfIndex).sPriceType = "N" Then       'fill or extra?
                                                                        slStr = "-Fill Outside"  ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                    Else
                                                                        'slStr = "Extra Outside" ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                        slStr = "+Fill Outside" ', Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                                    End If
                                                                End If
                                                                'slStr = "Outside, Missed " & Format$(tmSdfExt(ilSdfIndex).lMdDate, "m/d/yy")
                                                            Else
                                                                gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                                    slStr = "Missed, Outside " & slDate
                                                                Else
                                                                    '1-19-04 change way in which fill/extra are shown
                                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                                    If slShowOnInv = "N" Then
                                                                    'If tmSdfExt(ilSdfIndex).sPriceType = "N" Then       'fill or extra?
                                                                        slStr = "-Fill Missed, Outside " & slDate
                                                                    Else
                                                                        'slStr = "Extra Missed, Outside " & slDate
                                                                        slStr = "+Fill Missed, Outside " & slDate
                                                                    End If
                                                                End If
                                                                'slStr = "Missed, Outside " & slDate
                                                            End If
                                                        Case "M"
                                                            slStr = "Missed"
                                                        Case "C"
                                                            slStr = "Cancelled"
                                                        Case "H"
                                                            slStr = "Hidden"
                                                        Case "R"
                                                            slStr = "Ready MG"
                                                        Case "U"
                                                            slStr = "Unschedule MG"
                                                    End Select

                                                    slBBLength = ""
                                                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                                        If tmSdfExt(llSdfIndex).sSpotType = "O" Then
                                                            slBBLength = str$(tmClf.iBBOpenLen) & " BB"
                                                        ElseIf tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                                            slBBLength = str$(tmClf.iBBOpenLen) & " BB"
                                                        End If
                                                    Else
                                                        If tmSdfExt(llSdfIndex).sSpotType = "O" Then
                                                            slBBLength = str$(tmClf.iBBOpenLen) & " OBB"
                                                        ElseIf tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                                            slBBLength = str$(tmClf.iBBCloseLen) & " CBB"
                                                        End If
                                                    End If
                                                    If tmSdfExt(llSdfIndex).sBill = "Y" Then
                                                        slStr = slStr & " Billed "
                                                    End If
                                                    If tmSdfExt(llSdfIndex).sTracer = "M" Then  'Mouse move
                                                        slStr = slStr & "(am)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "P" Then  'Post Log
                                                        slStr = slStr & "(ap)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "C" Then  'Line Change
                                                        slStr = slStr & "(ac)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "U" Then  'Unschedule
                                                        slStr = slStr & "(au)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "1" Then  'Mouse move in past
                                                        slStr = slStr & "(bm)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "2" Then  'post log
                                                        slStr = slStr & "(bp)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "3" Then  'Line change
                                                        slStr = slStr & "(bc)"
                                                    ElseIf tmSdfExt(llSdfIndex).sTracer = "4" Then  'Unschedule
                                                        slStr = slStr & "(bu)"
                                                    End If
                                                    slInvalid = ""
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H10) = &H10 Then
                                                        slInvalid = " Invalid Price?"
                                                    End If
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H1) = &H1 Then
                                                        If slInvalid <> "" Then
                                                            slInvalid = slInvalid & ", Date"
                                                        Else
                                                            slInvalid = " Invalid Date"
                                                        End If
                                                    End If
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H2) = &H2 Then
                                                        If slInvalid <> "" Then
                                                            slInvalid = slInvalid & ", Time"
                                                        Else
                                                            slInvalid = " Invalid Time"
                                                        End If
                                                    End If
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H4) = &H4 Then
                                                        If slInvalid <> "" Then
                                                            slInvalid = slInvalid & ", Vehicle"
                                                        Else
                                                            slInvalid = " Invalid Vehicle"
                                                        End If
                                                    End If
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H8) = &H8 Then
                                                        If slInvalid <> "" Then
                                                            slInvalid = slInvalid & ", Length"
                                                        Else
                                                            slInvalid = " Invalid Length"
                                                        End If
                                                    End If
                                                    If (tmSdfExt(llSdfIndex).iStatus And &H20) = &H20 Then
                                                        If slInvalid <> "" Then
                                                            slInvalid = slInvalid & ", M for N"
                                                        Else
                                                            slInvalid = " M for N"
                                                        End If
                                                    End If

                                                    slStr = slStr & slBBLength      'contenate the bb length if applicable
                                                    If slInvalid <> "" Then
                                                        slStr = slStr & slInvalid
                                                    End If
                                                    smAired(3, llNoAiredRows) = slStr
                                                    'If (tmSdfExt(ilSdfIndex).ivefCode <> tmSmf.iOrigSchVef) And (ilSpotType = 2) And ((tmSdfExt(ilSdfIndex).sSchStatus = "G") Or (tmSdfExt(ilSdfIndex).sSchStatus = "O")) Then
                                                    If (tmSdfExt(llSdfIndex).iVefCode <> tmClf.iVefCode) And (ilSpotType = 2) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O")) Then
                                                        llNoAiredRows = llNoAiredRows + 1
                                                        If llNoAiredRows > UBound(smAired, 2) Then
                                                            'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                            ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                                        End If
                                                        If llNoAiredRows > llNoFlights Then
                                                            llNoFlights = llNoAiredRows
                                                            If llNoFlights > UBound(smOrdered, 2) Then
                                                                'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                                ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                                            End If
                                                        End If
                                                        tmVefSrchKey.iCode = tmSdfExt(llSdfIndex).iVefCode
                                                        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tlVef.sName = "Missing"
                                                        End If
                                                        'bad
                                                        'tmCbf.sLineSurvey = Trim$(tlVef.sname)
                                                        smAired(1, llNoAiredRows) = Trim$(tlVef.sName)
                                                        'Aired Day, Date Vehicle Override
                                                        tmCbf.iOurShare = tmSdfExt(llSdfIndex).iVefCode

                                                    End If
                                                    'If (tmSdfExt(ilSdfIndex).iVefCode <> tmSmf.iOrigSchVef) And (ilSpotType = 1) And ((tmSdfExt(ilSdfIndex).sSchStatus = "G") Or (tmSdfExt(ilSdfIndex).sSchStatus = "O")) Then
                                                    If (tmSdfExt(llSdfIndex).iVefCode <> tmClf.iVefCode) And (ilSpotType = 1) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O")) Then
                                                        llNoAiredRows = llNoAiredRows + 1
                                                        If llNoAiredRows > UBound(smAired, 2) Then
                                                            'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                            ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                                        End If
                                                        If llNoAiredRows > llNoFlights Then
                                                            llNoFlights = llNoAiredRows
                                                            If llNoFlights > UBound(smOrdered, 2) Then
                                                                'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                                ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                                            End If
                                                        End If
                                                        tmVefSrchKey.iCode = tmSdfExt(llSdfIndex).iVefCode
                                                        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tlVef.sName = "Missing"
                                                        End If
                                                        'tmCbf.sLineSurvey = Trim$(tlVef.sName)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            tmCbf.iVefCode = tmSdfExt(llSdfIndex).iVefCode
                                                        Else
                                                            tmCbf.iVefCode = -1
                                                        End If
                                                        smAired(3, llNoAiredRows) = Trim$(tlVef.sName)
                                                    End If
                                                    'Aired Status
                                                    'tmCbf.sSurvey = slStr
                                                    ilLen = Len(slStr)
                                                    tmCbf.sSurvey = Mid(slStr, 1, ilLen)
                                                    tmCbf.sPrefDT = ""                          '8-31-16 init previous text in field
                                                    If ilLen - 30 > 0 Then
                                                        tmCbf.sPrefDT = Mid(slStr, 31, ilLen - 30)
                                                    End If

                                                    'ilDummy = LLDefineFieldExt(hdJob, "Status", slStr, LL_TEXT, "")
                                                    '6/29/16: MOved setting when geting contract header to handle case where no spots exist in first week of the requested period
                                                    ''tmCbf.iGenTime(0) = igNowTime(0)
                                                    ''tmCbf.iGenTime(1) = igNowTime(1)
                                                    'tmCbf.iGenDate(0) = igNowDate(0)
                                                    'tmCbf.iGenDate(1) = igNowDate(1)
                                                    'gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                                    'tmCbf.lGenTime = lgNowTime
                                                    tmCbf.lLineNo = tmClf.iLine
                                                    'tmCbf.sProduct = tmChf.sProduct
                                                    'tmCbf.iAdfCode = tmChf.iAdfCode
                                                    'tmCbf.lContrNo = tmChf.lCntrNo
                                                    tmCbf.sType = tmChf.sType       '1-12-03
                                                    'If ilPkgSpot Then
                                                    '    tmCbf.sDysTms = ""
                                                    '    tmCbf.iLen = 0
                                                    '    tmCbf.sSortField1 = ""
                                                    '    tmCbf.sResort = ""
                                                    '    tmCbf.lGrimp = 0
                                                    'End If
                                                    mFindPkgReference
                                                    If ilDispOnly Then
                                                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                                    Else
                                                        If tmCbf.sResort <> "0   " Then ' And tmCbf.iPropOrdDate(0) <> 0 Then
                                                            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                                        End If
                                                    End If

                                                    'Init/Reset Fields
                                                    'tmCbf.sLineSurvey = ""
                                                    tmCbf.sResort = ""
                                                    tmCbf.sSurvey = ""
                                                    tmCbf.sPrefDT = ""          '2nd half of cbfsurvey field for msg field
                                                    tmCbf.iPropOrdDate(0) = 0
                                                    tmCbf.iPropOrdDate(1) = 0
                                                    tmCbf.iOurShare = 0
                                                    tmCbf.iVefCode = 0
                                                    '12-1-06 the daypart descriptions are not showing, comment out this test
                                                    'If ilDispOnly Then
                                                        'tmCbf.sDysTms = ""
                                                        'tmCbf.iLen = 0
                                                    'End If
                                                    tmCbf.lLineNo = 0
                                                    tmCbf.sSortField2 = ""
                                                    tmCbf.sSortField1 = ""
                                                    tmCbf.sResort = ""
                                                    tmCbf.lGRP = 0
                                                    tmCbf.sBuyer = ""
                                                End If
                                                If (Not ilShowSpot) Or (ilSpotType = 0) Or (ilSpotType = 2) Then
                                                    llLbcIndex = llLbcIndex + 1
                                                    ilSpotType = 0
                                                End If
                                            Loop
                                        End If          '12-20-00

                                        If Not ilSpotsFound Then            '12-26-07  Show ordered line when no airing spots were found (missing from system)
                                            smOrdered(1, 1) = Trim$(str$(tmClf.iLine))
                                            tmCbf.lLineNo = tmClf.iLine
                                            tmCbf.lExtra4Byte = tmClf.iPkLineNo
                                            mFindPkgReference
                                            If ilDispOnly Then
                                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                            Else
                                                If tmCbf.sResort <> "0   " Then ' And tmCbf.iPropOrdDate(0) <> 0 Then
                                                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                                End If
                                            End If


                                        End If

                                        'Add total week count
                                        If ilWkCount <> 1 Then
                                            llNoAiredRows = llNoAiredRows + 1
                                            If llNoAiredRows > UBound(smAired, 2) Then
                                                'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                            End If
                                            smAired(1, llNoAiredRows) = "Week Total: " & Trim$(str$(ilWkCount))
                                            If llNoAiredRows < llNoFlights Then
                                                llNoAiredRows = llNoFlights
                                                If llNoAiredRows > UBound(smAired, 2) Then
                                                    'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                    ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                                End If
                                            End If
                                        End If
                                        If llNoAiredRows > llNoFlights Then
                                            llNoFlights = llNoAiredRows
                                            If llNoFlights > UBound(smOrdered, 2) Then
                                                'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                            End If
                                        End If
                                    End If

                                    llWkDate = llWkDate + ilNoDayToSun + 1
                                    ilNoDayToSun = 6
                                    'tmCbf.iGenTime(0) = igNowTime(0)
                                    'tmCbf.iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmCbf.lGenTime = lgNowTime
                                    tmCbf.iGenDate(0) = igNowDate(0)
                                    tmCbf.iGenDate(1) = igNowDate(1)
                                    tmCbf.lLineNo = tmClf.iLine
                                    tmCbf.lExtra4Byte = tmClf.iPkLineNo
                                    tmCbf.sProduct = tmChf.sProduct
                                    tmCbf.iAdfCode = tmChf.iAdfCode
                                    tmCbf.lContrNo = tmChf.lCntrNo
                                    tmCbf.sType = tmChf.sType       '1-13-03

                                    'If ilPkgSpot Then   'Don't show ordered info for package spots
                                    '    tmCbf.sDysTms = ""
                                    '    tmCbf.iLen = 0
                                    '    tmCbf.sSortField1 = ""
                                    '    tmCbf.sSortField2 = ""
                                    '    tmCbf.sResort = ""
                                    '    tmCbf.lGrimp = -7
                                    'End If
                                    'If (ilCBS) Or ((tmCbf.sResort <> "0   ") And ((tmCbf.lGrimp <> 0) And Not (ilPkgSpot)) And (il1stSpotInWk = True)) Then
                                    mFindPkgReference
                                    If ilPkgSpot Then
                                        If (ilCBS) Or ((tmCbf.sResort <> "0   ") And ((tmCbf.lGrImp <> -7) And Not (ilPkgSpot)) And (il1stSpotInWk = True)) Then
                                        'If (ilCBS) Or ((tmCbf.sResort <> "0   ") And (il1stSpotInWk = True)) Then
                                            If Trim(tmCbf.sLineSurvey) <> "" Then
                                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                            End If
                                        End If
                                    End If
                                    If ilRepSpot Then
                                        If (ilCBS) Or ((tmCbf.sResort <> "0   ") And ((tmCbf.lGrImp <> -7) And Not (ilRepSpot)) And (il1stSpotInWk = True)) Then
                                        'If (ilCBS) Or ((tmCbf.sResort <> "0   ") And (il1stSpotInWk = True)) Then
                                            If Trim(tmCbf.sLineSurvey) <> "" Then
                                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                            End If
                                        End If
                                    End If
                                    'Init/Reset Fields
                                    ilCBS = False
                                    'tmCbf.sLineSurvey = ""
                                    ilContFirstTime = False
                                    tmCbf.sResort = ""
                                    tmCbf.sSurvey = ""
                                    tmCbf.sPrefDT = ""          '2nd half of cbfsurvey field for msg field
                                    tmCbf.iPropOrdDate(0) = 0
                                    tmCbf.iPropOrdDate(1) = 0
                                    tmCbf.iOurShare = 0
                                    tmCbf.iVefCode = 0
                                    '12-1-06 daypart descriptions are not showing, comment out so it shows
                                    'If ilDispOnly Then
                                        'tmCbf.sDysTms = ""
                                        'tmCbf.iLen = 0
                                    'End If
                                    tmCbf.lLineNo = 0
                                    tmCbf.sSortField2 = ""
                                    tmCbf.sSortField1 = ""
                                    tmCbf.sResort = "0"
                                    tmCbf.lGrImp = 0
                                    tmCbf.lGRP = 0
                                    tmCbf.sBuyer = ""

                                    il1stSpotInWk = True
                                Loop While llWkDate <= gDateValue(slEFlightDate)
                                ilCff = tgCffCB(ilCff).iNextCff
                            Loop
                            'Test for spots outside of week- scan to see if any require printing
                            ilShowSpot = False
                            tmCbf.sResort = "0"   '12-20-00  Need to show 0 ordered (not blank) for spots outside ordered week

                            For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1 Step 1 'RptSelCb!lbcSort.ListCount - 1 Step 1
                                slNameCode = tmSdfExtSort(llLbcIndex).sKey  'RptSelCb!lbcSort.List(ilLbcIndex)
                                'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex  'Val(slCode)
                                ilShowSpot = False
                                If Abs(tmSdfExt(llSdfIndex).iLineNo) = tgClfCB(ilClf).ClfRec.iLine Then
                                    'Only show spots within date span
                                    If (tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Then
                                        If (tmSdfExt(llSdfIndex).lMdDate >= llStartDate) And (tmSdfExt(llSdfIndex).lMdDate <= llEndDate) And ((tmSdfExt(llSdfIndex).iStatus And &H1000) <> &H1000) Then
                                            ilShowSpot = True
                                            Exit For
                                        Else
                                            gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                            If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) And ((tmSdfExt(llSdfIndex).iStatus And &H2000) <> &H2000) Then
                                                ilShowSpot = True
                                                Exit For
                                            End If
                                        End If
                                        'Show spots with date error even if not within weeks specified
                                        '8-17-05 ckcSelc3(0): show the spots with date error (that are outside the date range)
                                        If ilShowSpot Then      '1-4-10 if ok to show spot, test whether the invalid dates should be shown
                                            If ((tmSdfExt(llSdfIndex).iStatus And &H1) = &H1) And (RptSelCb!ckcSelC3(0).Value = vbChecked) Then
                                                ilShowSpot = True
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                        If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                            ilShowSpot = True
                                            Exit For
                                        End If
                                        'Show spots with date error even if not within weeks specified
                                          '8-17-05 ckcSelc3(0): show the spots with date error (that are outside the date range)
                                        If ilShowSpot Then      '1-4-10 if ok to show spot, test whether the invalid dates should be shown
                                            If ((tmSdfExt(llSdfIndex).iStatus And &H1) = &H1) And (RptSelCb!ckcSelC3(0).Value = vbChecked) Then
                                                ilShowSpot = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                    'If tmSdfExt(ilSdfIndex).iStatus > 0 Then
                                    '    ilShowSpot = True
                                    'End If
                                End If
                            Next llLbcIndex
                            If ilShowSpot Then
                                il1stSpotInWk = True
                                llNoFlights = llNoFlights + 1
                                If llNoFlights > UBound(smOrdered, 2) Then
                                    'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                    ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                End If
                                smOrdered(5, llNoFlights) = ""  'Invalid
                                smOrdered(6, llNoFlights) = ""
                                smOrdered(7, llNoFlights) = "0"
                                'Build Aired Array
                                ilWkCount = 0
                                For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1 Step 1 'RptSelCb!lbcSort.ListCount - 1 Step 1
                                    slNameCode = tmSdfExtSort(llLbcIndex).sKey  'RptSelCb!lbcSort.List(ilLbcIndex)
                                    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex  'Val(slCode)
                                    ilShowSpot = False
                                    If Abs(tmSdfExt(llSdfIndex).iLineNo) = tgClfCB(ilClf).ClfRec.iLine Then
                                        'Only show spots within date span
                                        tmSmf.iOrigSchVef = tmSdfExt(llSdfIndex).iVefCode
                                        If (tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Then
                                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tmSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                            'Obtain original dates
                                            'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                                            'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                                            'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                                            'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                                            'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                            'DL: 4-20-05 use key2 instead of key0 for speed
                                            tmSmfSrchKey2.lCode = tmSdf.lCode
                                            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                                If tmSmf.lSdfCode = tmSdf.lCode Then
                                                    Exit Do
                                                End If
                                                ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                            Loop
                                            If (tmSdfExt(llSdfIndex).lMdDate >= llStartDate) And (tmSdfExt(llSdfIndex).lMdDate <= llEndDate) And ((tmSdfExt(llSdfIndex).iStatus And &H1000) <> &H1000) Then
                                                ilShowSpot = True
                                            Else
                                                gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                                If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) And ((tmSdfExt(llSdfIndex).iStatus And &H2000) <> &H2000) Then
                                                    ilShowSpot = True
                                                End If
                                            End If
                                            'Show spots with date error even if not within weeks specified
                                            '8-17-05 ckcSelc3(0): show the spots with date error (that are outside the date range)
                                            If ilShowSpot Then      '1-4-10 if ok to show spot, test whether the invalid dates should be shown
                                                If ((tmSdfExt(llSdfIndex).iStatus And &H1) = &H1) And (RptSelCb!ckcSelC3(0).Value = vbChecked) Then
                                                    ilShowSpot = True
                                                End If
                                            End If
                                        Else
                                            gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                            If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                                                ilShowSpot = True
                                            End If
                                            'Show spots with date error even if not within weeks specified
                                            '8-17-05 ckcSelc3(0): show the spots with date error (that are outside the date range)
                                            If ilShowSpot Then      '1-4-10 if ok to show spot, test whether the invalid dates should be shown
                                                If ((tmSdfExt(llSdfIndex).iStatus And &H1) = &H1) And (RptSelCb!ckcSelC3(0).Value = vbChecked) Then
                                                    ilShowSpot = True
                                                End If
                                            End If
                                        End If
                                        'If tmSdfExt(ilSdfIndex).iStatus > 0 Then
                                        '    ilShowSpot = True
                                        'End If
                                    End If

                                    '6-26-06 option to show fill spots on Spot Placement
                                    If Not ilDispOnly Then      'spot placement vs Discrep only
                                        'spot placement, if "Fill" should it be shown based on user option
                                        If tmSdfExt(llSdfIndex).sSpotType = "X" And RptSelCb!ckcSelC3(1).Value = vbUnchecked Then
                                            ilShowSpot = False
                                        End If
                                        
                                        '5-8-14 ignore open/close bb in future
                                        'Test if Open or Close BB, ignore if in the future
                                        If tmSdfExt(llSdfIndex).sSpotType = "O" Or tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                            gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slSpotDate
                                            If gDateValue(slSpotDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
                                                ilShowSpot = False
                                            End If
                                        End If
                                                                            
                                    Else            '11-23-11 if discrep, dont show the obb and cbb
                                        If tmSdfExt(llSdfIndex).sSpotType = "O" Or tmSdfExt(llSdfIndex).sSpotType = "C" Then
                                            ilShowSpot = False
                                        End If
                                    End If
                                    If ilShowSpot Then
                                        tmSdfExt(llSdfIndex).iLineNo = 0
                                        ilWkCount = ilWkCount + 1
                                        If tmVef.sType <> "V" Then
                                            llNoAiredRows = llNoAiredRows + 1
                                        Else
                                            If il1stSpotInWk Then
                                                llNoAiredRows = llNoAiredRows + 2
                                            Else
                                                llNoAiredRows = llNoAiredRows + 1
                                            End If
                                        End If
                                        il1stSpotInWk = False
                                        If llNoAiredRows > UBound(smAired, 2) Then
                                            'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                            ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                        End If
                                        If llNoAiredRows > llNoFlights Then
                                            llNoFlights = llNoAiredRows
                                            If llNoFlights > UBound(smOrdered, 2) Then
                                                'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                            End If
                                        End If
                                        If tmVef.sType = "V" Then
                                            If (tmSdfExt(llSdfIndex).iVefCode <> tmSmf.iOrigSchVef) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O")) Then
                                                tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                                                ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tlVef.sName = "Missing"
                                                End If
                                                slStr = Trim$(tlVef.sName)
                                            Else
                                                ilRet = gParseItem(slNameCode, 2, "|", slStr)
                                            End If
                                            smOrdered(2, llNoAiredRows) = "  " & Trim$(slStr)
                                            tmCbf.sLineSurvey = Trim(slStr)
                                        End If
                                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slDate
                                        tmCbf.iPropOrdDate(0) = tmSdfExt(llSdfIndex).iDate(0)
                                        tmCbf.iPropOrdDate(1) = tmSdfExt(llSdfIndex).iDate(1)
                                        llDate = gDateValue(slDate)
                                        smAired(1, llNoAiredRows) = Format$(llDate, "ddd") & ", " & slDate
                                        'Aired Day, Date
                                        'tmCbf.sLineSurvey = slStr
                                        'ilDummy = LLDefineFieldExt(hdJob, "AirDay", slStr, LL_TEXT, "")
                                        If (tmSdfExt(llSdfIndex).sSchStatus = "S") Or (tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Then
                                            gUnpackTime tmSdfExt(llSdfIndex).iTime(0), tmSdfExt(llSdfIndex).iTime(1), "A", "1", slTime
                                            smAired(2, llNoAiredRows) = slTime
                                        End If
                                        'Aired Time
                                        tmCbf.sBuyer = slTime
                                        'ilDummy = LLDefineFieldExt(hdJob, "AirTime", slTime, LL_TEXT, "")
                                        Select Case tmSdfExt(llSdfIndex).sSchStatus
                                            Case "S"
                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                    slStr = "Scheduled"
                                                Else
                                                   '1-19-04 change way in which fill/extra are shown
                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                    'If tmSdfExt(llSdfIndex).sPriceType = "N" Then       'fill or extra?
                                                    If slShowOnInv = "N" Then
                                                        slStr = "-Schd Fill"
                                                    Else
                                                        'slStr = "Schd Extra"
                                                        slStr = "+Schd Fill"
                                                    End If
                                                End If
                                            Case "G"
                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                    If tmSmf.lMtfCode > 0 Then              '12-20-00
                                                        slStr = "MG M for N"
                                                    Else
                                                        slStr = "MG for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    End If
                                                Else
                                                    '1-19-04 change way in which fill/extra are shown
                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                    If slShowOnInv = "N" Then
                                                    'If tmSdfExt(llSdfIndex).sPriceType = "N" Then
                                                        slStr = "-Fill MG for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    Else
                                                        'slStr = "Extra MG for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                        slStr = "+Fill MG for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    End If
                                                End If
                                                'slStr = "MG for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                            Case "O"
                                                If tmSdfExt(llSdfIndex).sSpotType <> "X" Then
                                                    If tmSmf.lMtfCode > 0 Then              '12-20-00
                                                        slStr = "MG M for N"
                                                    Else
                                                        slStr = "Outside for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    End If
                                                Else
                                                    '1-19-04 change way in which fill/extra are shown
                                                    slShowOnInv = gTestShowFill(tmSdfExt(llSdfIndex).sPriceType, tmSdfExt(llSdfIndex).iAdfCode)
                                                    If slShowOnInv = "N" Then
                                                    'If tmSdfExt(llSdfIndex).sPriceType = "N" Then
                                                        slStr = "-Fill Outside for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    Else
                                                        'slStr = "Extra Outside for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                        slStr = "+Fill Outside for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                                    End If
                                                End If
                                                'slStr = "Outside for " & Format$(tmSdfExt(llSdfIndex).lMdDate, "m/d/yy")
                                            Case "M"
                                                slStr = "Missed"
                                            Case "C"
                                                slStr = "Cancelled"
                                            Case "H"
                                                slStr = "Hidden"
                                            Case "R"
                                                slStr = "Ready MG"
                                            Case "U"
                                                slStr = "Unschedule MG"
                                        End Select
                                        If tmSdfExt(llSdfIndex).sBill = "Y" Then
                                            slStr = slStr & " Billed "
                                        End If
                                        If tmSdfExt(llSdfIndex).sTracer = "M" Then  'Mouse move
                                            slStr = slStr & " (am)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "P" Then  'Post Log
                                            slStr = slStr & " (ap)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "C" Then  'Line Change
                                            slStr = slStr & " (ac)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "U" Then  'Unschedule
                                            slStr = slStr & " (au)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "1" Then  'Mouse move in past
                                            slStr = slStr & " (bm)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "2" Then  'post log
                                            slStr = slStr & " (bp)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "3" Then  'Line change
                                            slStr = slStr & " (bc)"
                                        ElseIf tmSdfExt(llSdfIndex).sTracer = "4" Then  'Unschedule
                                            slStr = slStr & " (bu)"
                                        End If
                                        slInvalid = ""
                                        If (tmSdfExt(llSdfIndex).iStatus And &H10) = &H10 Then
                                            slInvalid = " Invalid Price?"
                                        End If
                                        If (tmSdfExt(llSdfIndex).iStatus And &H1) = &H1 Then
                                            If slInvalid <> "" Then
                                                slInvalid = slInvalid & ", Date"
                                            Else
                                                slInvalid = " Invalid Date"
                                            End If
                                        End If
                                        If (tmSdfExt(llSdfIndex).iStatus And &H2) = &H2 Then
                                            If slInvalid <> "" Then
                                                slInvalid = slInvalid & ", Time"
                                            Else
                                                slInvalid = " Invalid Time"
                                            End If
                                        End If
                                        If (tmSdfExt(llSdfIndex).iStatus And &H4) = &H4 Then
                                            If slInvalid <> "" Then
                                                slInvalid = slInvalid & ", Vehicle"
                                            Else
                                                slInvalid = " Invalid Vehicle"
                                            End If
                                        End If
                                        If (tmSdfExt(llSdfIndex).iStatus And &H8) = &H8 Then
                                            If slInvalid <> "" Then
                                                slInvalid = slInvalid & ", Length"
                                            Else
                                                slInvalid = " Invalid Length"
                                            End If
                                        End If
                                        If (tmSdfExt(llSdfIndex).iStatus And &H20) = &H20 Then
                                            If slInvalid <> "" Then
                                                slInvalid = slInvalid & ", M for N"
                                            Else
                                                slInvalid = " M for N"
                                            End If
                                        End If
    '**************************************************
                                        If slInvalid <> "" Then
                                            slStr = slStr & slInvalid
                                        End If
                                        smAired(3, llNoAiredRows) = slStr
                                        If (tmSdfExt(llSdfIndex).iVefCode <> tmSmf.iOrigSchVef) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O")) Then
                                            llNoAiredRows = llNoAiredRows + 1
                                            If llNoAiredRows > UBound(smAired, 2) Then
                                                'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                            End If
                                            If llNoAiredRows > llNoFlights Then
                                                llNoFlights = llNoAiredRows
                                                If llNoFlights > UBound(smOrdered, 2) Then
                                                    'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                    ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                                End If
                                            End If
                                            tmVefSrchKey.iCode = tmSdfExt(llSdfIndex).iVefCode
                                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                tlVef.sName = "Missing"
                                            End If
                                            smAired(1, llNoAiredRows) = Trim$(tlVef.sName)
                                            tmCbf.sLineSurvey = Trim$(tlVef.sName)
                                        End If
                                        'Show in aired day date or status column the vehicle for mg spots for m for n customers
                                        If (lgMtfNoRecs > 0) And (tmSmf.lMtfCode > 0) And ((ilPkgSpot) Or (ilRepSpot)) And ((tmSdfExt(llSdfIndex).sSchStatus = "G") Or (tmSdfExt(llSdfIndex).sSchStatus = "O") Or (tmSdfExt(llSdfIndex).sSchStatus = "M")) Then
                                            llNoAiredRows = llNoAiredRows + 1
                                            If llNoAiredRows > UBound(smAired, 2) Then
                                                'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                                ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                            End If
                                            If llNoAiredRows > llNoFlights Then
                                                llNoFlights = llNoAiredRows
                                                If llNoFlights > UBound(smOrdered, 2) Then
                                                    'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                                    ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                                End If
                                            End If
                                            tmVefSrchKey.iCode = tmSdfExt(llSdfIndex).iVefCode
                                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                tlVef.sName = "Missing"
                                            End If
                                            smAired(1, llNoAiredRows) = Trim$(tlVef.sName)
                                            If tmSdfExt(llSdfIndex).sSchStatus = "M" Then
                                                'Aired Status column for m for n customers
                                                tmCbf.iVefCode = tlVef.iCode
                                            Else
                                                'Aired Day Date column for m for n customers
                                                tmCbf.iOurShare = tlVef.iCode
                                            End If
                                        End If
                                        'Aired Status
                                         ilLen = Len(slStr)
                                        tmCbf.sSurvey = Mid(slStr, 1, ilLen)
                                        tmCbf.sPrefDT = ""                      '8-31-16 init previous text in field
                                        If ilLen - 30 > 0 Then
                                            tmCbf.sPrefDT = Mid(slStr, 31, ilLen - 30)
                                        End If
                                        'ilDummy = LLDefineFieldExt(hdJob, "Status", slStr, LL_TEXT, "")
                                        'tmCbf.iGenTime(0) = igNowTime(0)
                                        'tmCbf.iGenTime(1) = igNowTime(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tmCbf.lGenTime = lgNowTime
                                        tmCbf.iGenDate(0) = igNowDate(0)
                                        tmCbf.iGenDate(1) = igNowDate(1)
                                        tmCbf.lLineNo = tmClf.iLine
                                        tmCbf.lExtra4Byte = tmClf.iPkLineNo
                                        tmCbf.sProduct = tmChf.sProduct
                                        tmCbf.iAdfCode = tmChf.iAdfCode
                                        tmCbf.lContrNo = tmChf.lCntrNo
                                        tmCbf.sType = tmChf.sType       '1-13-03

                                        If ilPkgSpot Then
                                            tmCbf.sDysTms = "(Package)"
                                            tmCbf.iLen = 0
                                            tmCbf.sResort = ""
                                            tmCbf.lGrImp = 0
                                        End If
                                        If ilRepSpot Then
                                            tmCbf.sDysTms = "(Rep)"
                                            tmCbf.iLen = 0
                                            tmCbf.sResort = ""
                                            tmCbf.lGrImp = 0
                                        End If
                                        If ilRet = ilRet Then ' And tmCbf.iPropOrdDate(0) <> 0 Then
                                            'If Not ilPkgSpot Or (ilPkgSpot And tmCbf.sLineSurvey <> "") Then
                                                mFindPkgReference
                                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                            'End If
                                        End If
                                    End If
                                    tmCbf.iOurShare = 0
                                    tmCbf.iVefCode = 0
                                Next llLbcIndex

                                'Add total week count
                                llNoAiredRows = llNoAiredRows + 1
                                If llNoAiredRows > UBound(smAired, 2) Then
                                    'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                    ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                End If
                                If llNoAiredRows > llNoFlights Then
                                    llNoFlights = llNoAiredRows
                                    If llNoFlights > UBound(smOrdered, 2) Then
                                        'ReDim Preserve smOrdered(1 To 8, 1 To llNoFlights) As String
                                        ReDim Preserve smOrdered(0 To 8, 0 To llNoFlights) As String
                                    End If
                                End If
                                smAired(1, llNoAiredRows) = "Week Total: " & Trim$(str$(ilWkCount))
                                If llNoAiredRows < llNoFlights Then
                                    llNoAiredRows = llNoFlights
                                    If llNoAiredRows > UBound(smAired, 2) Then
                                        'ReDim Preserve smAired(1 To 3, 1 To llNoAiredRows) As String
                                        ReDim Preserve smAired(0 To 3, 0 To llNoAiredRows) As String
                                    End If
                                End If
                            End If
                            llNoOrderedRows = UBound(smOrdered, 2)
                            llNoOrderedRowsPrt = 0
                            llNoAiredRowsPrt = 0
                            If Not ilHeaderInit Then
                                slStr = Trim$(str$(tgChfCB.lCntrNo)) & Chr$(0)
                                'Contract Number
                                tmCbf.lContrNo = tgChfCB.lCntrNo
                                'ilDummy = LLDefineVariableExt(hdJob, "CntrNo", slStr, LL_TEXT, "")
                                If tgChfCB.iAdfCode <> tmAdf.iCode Then
                                    tmAdfSrchKey.iCode = tgChfCB.iAdfCode
                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tmAdf.sName = "Missing"
                                    End If
                                End If
                                slStr = Trim$(tmAdf.sName) & Chr$(0)
                                sAdvName = tmAdf.sName
                                'ilDummy = LLDefineVariableExt(hdJob, "AdvtName", slStr, LL_TEXT, "")
                                slStr = Trim$(tmChf.sProduct) & Chr$(0)
                                tmCbf.sProduct = tmChf.sProduct
                                'ilDummy = LLDefineVariableExt(hdJob, "Product", slStr, LL_TEXT, "")
                                ilHeaderInit = True
                            End If
                            ilLineInit = False
                            ilLineEOF = False
                            'outer loop - one loop per page
                            'Do While (Not ilLineEof) And ilErrorFlag = 0 And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
                            Do While (Not ilLineEOF) And ilErrorFlag = 0
                                Screen.MousePointer = vbDefault
                                If ilNewPage Then
                                    Screen.MousePointer = vbDefault
                                    'ilLLRet = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter & Chr$(10) & "For" & Str$(tmChf.lCntrNo), (100# * llRecNo / llRecsRemaining))
                                    'DoEvents
                                    'If ilLLRet <> 0 Then
                                    '    ilErrorFlag = ilLLRet
                                    '    Exit Do
                                    'End If
                                    'All object must be enabled to be printed
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", True)
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Aired", True)
    'VB6**                                ilLLRet = LLPrint(hdJob)
                                    ilNewPage = False
                                    ilMaxRowPerPage = 37    'number of rows without lines'LLPrintGetRemainingItemsPerTable(hdJob, ":Ordered")
                                    llRowRemainingPerPage = ilMaxRowPerPage
                                    ilNoLinesPerPage = 0
                                    llNoTotalLines = 0
                                End If
                                'Determine number of rows to output
                                llNoRowsToPrt = llRowRemainingPerPage - ilNoLinesPerPage \ 7 '- llNoTotalLines 'LLPrintGetRemainingItemsPerTable(hdJob, ":Ordered")
                                If ((llNoOrderedRows - llNoOrderedRowsPrt) < llNoRowsToPrt) And ((llNoAiredRows - llNoAiredRowsPrt) < llNoRowsToPrt) Then
                                    If (llNoOrderedRows - llNoOrderedRowsPrt) > (llNoAiredRows - llNoAiredRowsPrt) Then
                                        llNoRowsToPrt = llNoOrderedRows - llNoOrderedRowsPrt
                                    Else
                                        llNoRowsToPrt = llNoAiredRows - llNoAiredRowsPrt
                                    End If
                                End If
                                If llNoRowsToPrt > 0 Then
                                    'Output any ordered rows
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Aired", False)
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", True)
                                    'LlDefineFieldStart hdJob

                                    'For ilField = LBound(smOrdered, 1) To UBound(smOrdered, 1) Step 1
                                    For ilField = LBONE To UBound(smOrdered, 1) Step 1
                                        slStr = ""
                                        For llRow = 1 To llNoRowsToPrt Step 1
                                            If slStr <> "" Then
                                                slStr = slStr + Chr$(10)
                                            End If
                                            If llNoOrderedRowsPrt + llRow <= llNoOrderedRows Then
                                                slStr = slStr & smOrdered(ilField, llNoOrderedRowsPrt + llRow)
                                           End If
                                            If slStr = "" Then
                                                slStr = " "
                                            End If
                                        Next llRow
                                        slStr = slStr & Chr$(0)
                                        Select Case ilField
                                            Case 1
                                                'ilDummy = LLDefineFieldExt(hdJob, "LineNo", slStr, LL_TEXT, "")
                                                'tmCbf.lLineNo = gStrDecToLong(slStr, 2)
                                            Case 2
                                                'ilDummy = LLDefineFieldExt(hdJob, "Vehicle", slStr, LL_TEXT, "")
                                                'tmCbf.ivefCode = tmVef.icode
                                                'tmCbf.sLineSurvey = Trim$(slStr)
                                            Case 3
                                                'ilDummy = LLDefineFieldExt(hdJob, "Times", slStr, LL_TEXT, "")
                                                'tmCbf.sDysTms = slStr
                                            Case 4
                                                'ilDummy = LLDefineFieldExt(hdJob, "Length", slStr, LL_TEXT, "")
                                                'tmCbf.iLen = gStrDecToInt(slStr, 0)
                                            Case 5
                                                'ilDummy = LLDefineFieldExt(hdJob, "Dates", slStr, LL_TEXT, "")
                                                'tmCbf.sSortField1 = slStr
                                                'tmCbf.lDrfCode = DateValue(slDate)
                                            Case 6
                                                'ilDummy = LLDefineFieldExt(hdJob, "Days", slStr, LL_TEXT, "")
                                                'tmCbf.sSortField2 = slStr
                                            Case 7
                                                'ilDummy = LLDefineFieldExt(hdJob, "NoPerWk", slStr, LL_TEXT, "")
                                                'tmCbf.sResort = slStr
                                            Case 8
                                                If (Left$(slStr, 1) = " ") Or (Left$(slStr, 1) >= "A") And (Left$(slStr, 1) <= "Z") Then
                                                    'ilDummy = LLDefineFieldExt(hdJob, "Rate", slStr, LL_TEXT, "")
                                                    'tmCbf.lGrp = gStrDecToLong(slStr, 2)
                                                   ' tmCbf.lGrp = 0
                                                Else
                                                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                                                    'ilDummy = LLDefineFieldExt(hdJob, "Rate", slStr, LL_TEXT, "")
                                                    'tmCbf.lGrp = 0
                                                    'tmCbf.lGrp = gStrDecToLong(slStr, 2)
                                                End If
                                        End Select
                                    Next ilField

    'VB6**                                ilLLRet = LlPrintFields(hdJob)
                                    llNoOrderedRowsPrt = llNoOrderedRowsPrt + llNoRowsToPrt
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", False)
    'VB6**                                ilLLRet = LLPrintEnableObject(hdJob, ":Aired", True)
                                    'For ilField = LBound(smAired, 1) To UBound(smAired, 1) Step 1
                                    For ilField = LBONE To UBound(smAired, 1) Step 1
                                        slStr = ""
                                        For llRow = 1 To llNoRowsToPrt Step 1
                                            If slStr <> "" Then
                                                slStr = slStr + Chr$(10)
                                            End If
                                            If llNoAiredRowsPrt + llRow <= llNoAiredRows Then
                                                slStr = slStr & smAired(ilField, llNoAiredRowsPrt + llRow)
                                            End If
                                            If slStr = "" Then
                                                slStr = " "
                                            End If
                                        Next llRow
                                        slStr = slStr & Chr$(0)
                                        Select Case ilField
                                            Case 1
                                                'ilDummy = LLDefineFieldExt(hdJob, "AirDay", slStr, LL_TEXT, "")
                                               ''' tmCbf.sLineSurvey = slStr
                                            Case 2
                                                'ilDummy = LLDefineFieldExt(hdJob, "AirTime", slStr, LL_TEXT, "")
                                                'tmCbf.sBuyer = slStr
                                            Case 3
                                                'ilDummy = LLDefineFieldExt(hdJob, "Status", slStr, LL_TEXT, "")
                                                'tmCbf.sSurvey = slStr
                                        End Select
                                    Next ilField
    'VB6**                                ilLLRet = LlPrintFields(hdJob)
                                    llNoAiredRowsPrt = llNoAiredRowsPrt + llNoRowsToPrt
                                    llRowRemainingPerPage = llRowRemainingPerPage - llNoRowsToPrt
                                    ilNoLinesPerPage = ilNoLinesPerPage + 1
                                End If
                                If (llNoOrderedRowsPrt >= llNoOrderedRows) And (llNoAiredRowsPrt >= llNoAiredRows) Then
                                    ilLineEOF = True
                                    llNoTotalLines = llNoTotalLines + 1
                                Else
                                    If (llRowRemainingPerPage <= 0) Or (llNoRowsToPrt <= 0) Then
                                        ilNewPage = True
                                    End If
                                End If
                            Loop    ' while not EOF
                            If ilErrorFlag <> 0 Then
                                Exit For
                            End If
                        End If
                        If ilErrorFlag <> 0 Then
                            Exit For
                        End If
                    End If                      'tmclf.stype <> "A" and tmclf.stype <> "O"
                Next ilClf
                If ilErrorFlag <> 0 Then
                    Exit Do
                End If
        'Loop While llRecNo < llRecsRemaining
        Loop
        If (Not ilAnyOutput) And ilDispOnly Then
            'ilDummy = LLDefineVariableExt(hdJob, "CntrNo", "NO DISCREPANCIES", LL_TEXT, "")
            'ilDummy = LLDefineVariableExt(hdJob, "AdvtName", "", LL_TEXT, "")
            'ilDummy = LLDefineVariableExt(hdJob, "Product", "", LL_TEXT, "")
            'ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", False)
            'ilLLRet = LLPrintEnableObject(hdJob, ":Aired", False)
            'ilLLRet = LLPrint(hdJob)
        ElseIf (Not ilAnyOutput) Then
            'ilDummy = LLDefineVariableExt(hdJob, "CntrNo", "NO OUTPUT", LL_TEXT, "")
            'ilDummy = LLDefineVariableExt(hdJob, "AdvtName", "", LL_TEXT, "")
            'ilDummy = LLDefineVariableExt(hdJob, "Product", "", LL_TEXT, "")
            'ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", False)
            'ilLLRet = LLPrintEnableObject(hdJob, ":Aired", False)
            'ilLLRet = LLPrint(hdJob)
        Else
            'ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", True)
            'ilLLRet = LLPrintEnableObject(hdJob, ":Aired", True)
        End If
        'ilLLRet = LlPrintEnd(hdJob, 0)

        'end print
        'ilLLRet = LLPrintEnableObject(hdJob, ":Ordered", True)
        'ilLLRet = LLPrintEnableObject(hdJob, ":Aired", True)
        'ilLLRet = LlPrintEnd(hdJob, 0)
        'in case of preview: show the preview
        'If ilPreview <> 0 Then
        '    If ilErrorFlag = 0 Then
        '        ilDummy = LlPreviewDisplay(hdJob, slFileName, sgRptSavePath, RptSelCB.hWnd)
        '    Else
        '        mErrMsg ilErrorFlag
        '    End If
        '    ilDummy = LlPreviewDeleteFiles(hdJob, slFileName, sgRptSavePath)
        'Else
        '    If ilErrorFlag <> 0 Then
        '        mErrMsg ilErrorFlag
        '    End If
        'End If
    Else  ' LlPrintWithBoxStart
        ilErrorFlag = ilLLRet
'VB6**        ilLLRet = LlPrintEnd(hdJob, 0)
        mErrMsg ilErrorFlag
    End If  ' LlPrintWithBoxStart

    If imUpdateCntrNo Then
        Do
            ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, lmSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            'tmSRec = tmSpf
            'ilRet = gGetByKeyForUpdate("Spf", hmSpf, tmSRec)
            'tmSpf = tmSRec
            tmSpf.lDiscCurrCntrNo = 0
            ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    Erase tmSdfExt
    Erase tmSdfExtSort
    Erase tgClfCB
    Erase tgCffCB
    Erase smOrdered
    Erase smAired
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hmCbf)
    ilRet = btrClose(hmSpf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmCbf
    btrDestroy hmSpf
    btrDestroy hmSsf
    btrDestroy hmVsf
    btrDestroy hmVef
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmAdf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmRdf
End Sub

'******************************************************************************
'*
'*      Procedure Name:gSpotAdvtRpt
'       SPOTS BY ADVT OR MG REVENUE REPORTS
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:12/3/98       By:D. Hosaka
'*
'*            Comments: Generate Spot by Date report
'
'*      5/28/98 DH :  if Extra or Fill spot, leave
'                     Status column blank (previously
'                     was flagged as Outside)
'                     Show only exceptions such as MG
'                     Missed, Outside (not fill) rather
'                     than Scheduled status
'*      12/3/98 DH : Convert from "Bridge" to Crystal
'*                   Sort & select by either advt, agency or
'*                   slsp
'       6-16-00 Add option for MG Revenue report
'       6-30-00 gather and write btr records one vehicle at a time
'           vs gathering all spots for all vehicles to avoid
'           # spots > 32766
'*      8-23-00 tmplsdf was not initialized for each vehicle due to above change
'*              (32,000 spots)
'       7-20-04 Option to include/exclude local/network spots
'       11-30-04 change to access smf by key2 instead of key0
'       4-22-05 Speed up MG Revenue report by sorting the spot table results
'               by advt, then test for advt already in memory to
'               avoid rereading advt.
'       5-24-06 show x-mid spots on true date it aired
'       10-2-08 when all advertisers selected, the contract header is not retrieved to check
'               whether commissionable or not
'******************************************************************************
Sub gSpotAdvtRpt()
'   ilSpotType(I)- 1=Scheduled only; 2=Missed only; 3=Both
    Dim slName As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim slStartDate As String   'aired start dates for spots by advt, missed dates for mg revenue (decreased by 30 days)
    Dim llStartDate As Long
    Dim slEndDate As String     'aired end dates for spots by advt, missed dates for mg revenue (increased by 30 days)
    Dim llEndDate As Long
    Dim slStartEDate As String
    Dim slEndEDate As String
    'Dim ilIndex As Integer     'chg to long
    Dim llIndex As Long         '4-22-05
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim ilSelType As Integer
    Dim ilShowPrice As Integer
    Dim ilVsf As Integer
    Dim ilCostType As Integer
    Dim tlVef As VEF
    Dim ilListIndex As Integer          'report option (mgRevenue vs Spots by Advt)
    Dim ilFound As Integer
    Dim ilIncludeMG As Integer
    Dim ilIncludeOut As Integer
    Dim ilIncludeMissed As Integer
    Dim ilIncludeCancel As Integer
    Dim ilSwitch As Integer             'show only those spots whose mg/out in different vehicles than ordered
    Dim ilLocal As Integer
    Dim ilFeed As Integer
    Dim llUpper As Long                 '4-22-05
    Dim ilDate(0 To 1) As Integer
    Dim ilCommPct As Integer
    Dim slAmount As String
    Dim slSharePct As String
    Dim ilNet As Integer
    Dim ilIncludeCodes As Integer           '11-25-09
    Dim ilSaveInclMGForRev As Integer       '10-20-10 save check box input for inclusion of mg rates.
                                            'most reports count line rate of MG and mg/out as one; but
                                            'mg revenue needs to keep those types separate
    Dim slMGStartDate As String     '3-26-11
    Dim slMGEndDAte As String
    Dim llMGStartDate As Long
    Dim llMGEndDate As Long
    Dim ilBillStatus As Integer     '3-27-11 0 = both, 1 = billed only, 2 = unbilled only
    Dim slHdr As String
    Dim slMissedStart As String
    Dim slMissedEnd As String
    Dim slMGStart As String
    Dim slMGEnd As String
    Dim ilWhichKey As Integer           '10-23-14  if single advt selection , use key by advt; other use key1 for sdf retrieval
    Dim ilGameSelect As Integer
    Dim ilLineSelect As Integer
    Dim llChfSelect As Long
    Dim ilLoopOnKey As Integer
    Dim ilSelectionIndex As Integer
    Dim blValidSelection As Boolean
    Dim ilAudioIndex As Integer                 '0 = all audio types
    Dim ilUseAcqRate As Integer                 '11-5-15 option if using barters and showing rates
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean
    Dim tlCntTypes As CNTTYPES                  '12-28-17
    Dim ilIncludeISCI As Integer                '12-13/18
    
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String
    Dim llWhichCode As Long
    Dim ilRemainMonths As Integer
    Dim llCPMStartDate As Long
    Dim llCPMEndDate As Long
    Dim ilNoDays As Integer
    Dim slRvfStart As String                    'earliest date to retrieve billed adserver trans
    Dim slRvfEnd As String                      'latest date to retrieve billed adserver trans
    Dim llLastBilledStd As Long
    Dim llLastBilledCal As Long
    Dim tlTranTypes As TRANTYPES            '12-29-06
    Dim llAmount As Long
    Dim llTotalReceivedAmount As Long
    Dim slLastRvf As String 'prevent loading RVF for same selection
    Dim dlDailyAmt As Double
    Dim ilDaysBilled As Integer
    Dim ilDaysUnbilled As Integer
    Dim llRemainingAmount As Long
    Dim dlTotalAmount As Double
    Dim slDate As String
    Dim llPCFIndex As Long
    Dim slFileName As String
    Dim slRepeat As String
    Dim tlExp As EXPWOINVLN 'For CSV Export
    Dim llTime As Long
    Dim slSpotPriceType As String
    Dim slLastBilledDate As String 'JW 6/16/23 for clarification, storing LastBilled Date here (Cal or Std depending on CntrBillCycle)
    Dim llTestCntrNo As Long
    
    'TTP 10961 - Spot and Digital Line combo report: show more information about digital line invoice adjustments on special export version
    Dim slTranComments() As String
    llTestCntrNo = 0
    If RptSelCb.edcSet3.Text <> "" Then
        llTestCntrNo = Val(RptSelCb.edcSet3.Text)
    End If
    Dim ilVpf As Integer
    
    ilListIndex = RptSelCb!lbcRptType.ListIndex
    Screen.MousePointer = vbHourglass

    'TTP 10674 - Get Last Billed Date
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate 'convert last bdcst billing date to string
    llLastBilledStd = gDateValue(slDate)            'convert last month billed to long
    gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate 'convert last bdcst billing date to string
    llLastBilledCal = gDateValue(slDate)            'convert last month billed to long
    
    'TTP 10674 - the Tran typs are to retrieve receivables for past billing. the remaining $ on adserver IDs will be averaged
    'TTP 10892 - Spot and Digital Line Combo report: new option to exclude invoice adjustments
    'tlTranTypes.iAdj = True
    If RptSelCb!chkIncludeAdjustments.Value = vbChecked And RptSelCb!ckcSelDigital.Value = vbChecked Then
        tlTranTypes.iAdj = True
    Else
        tlTranTypes.iAdj = False
    End If
    tlTranTypes.iAirTime = True
    tlTranTypes.iCash = True
    tlTranTypes.iInv = True
    tlTranTypes.iMerch = False
    tlTranTypes.iNTR = False
    tlTranTypes.iPromo = False
    tlTranTypes.iPymt = False
    tlTranTypes.iTrade = False
    tlTranTypes.iWriteOff = False
    
    'TTP 10674 reset last looked up
    smLastCLF = ""
    
    'SPOTS BY ADVT OR MG REVENUE REPORTS
    If ilListIndex = CNT_MGREVENUE Then           '12-28-17
        ilIncludeMG = gSetCheck(RptSelCb!ckcSelC6(0).Value)    'Include MG
        ilIncludeOut = gSetCheck(RptSelCb!ckcSelC6(1).Value)   'include Outsides
        ilIncludeMissed = gSetCheck(RptSelCb!ckcSelC6(2).Value)    'include missed spots
        ilIncludeCancel = gSetCheck(RptSelCb!ckcSelC6(3).Value)    'include cancelled spots
        ilSwitch = gSetCheck(RptSelCb!ckcSelC6(4).Value)       'include only spot whose mg/out in differ vehicles than ordered
        ilLocal = gSetCheck(RptSelCb!ckcSelC12(0).Value)       'include local spots (vs network feed)
        ilFeed = gSetCheck(RptSelCb!ckcSelC12(1).Value)       'include network (feed) vs local
        '12-28-17
        tlCntTypes.iHold = True
        tlCntTypes.iOrder = True
        tlCntTypes.iStandard = True
        tlCntTypes.iReserv = True
        tlCntTypes.iRemnant = True
        tlCntTypes.iDR = True
        tlCntTypes.iPI = True
        tlCntTypes.iPSA = True
        tlCntTypes.iPromo = True
    End If
    ilAudioIndex = 0                        'assume all audio types , unless spots by advt option

    ilUseAcqRate = False                    'default
    If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then            '10-6-15 only Spots by advt can select audio types
        If ilListIndex <> CNT_SPTCOMBO Then
            ilAudioIndex = RptSelCb!cbcSet1.ListIndex
        End If
        ilUseAcqRate = gSetCheck(RptSelCb!ckcSelC8(0).Value)        'use acq rate
        '12-28-17
        tlCntTypes.iHold = gSetCheck(RptSelCb!ckcSelC3(0).Value)
        tlCntTypes.iOrder = gSetCheck(RptSelCb!ckcSelC3(1).Value)
        tlCntTypes.iStandard = gSetCheck(RptSelCb!ckcSelC3(2).Value)
        tlCntTypes.iReserv = gSetCheck(RptSelCb!ckcSelC3(3).Value)
        tlCntTypes.iRemnant = gSetCheck(RptSelCb!ckcSelC3(4).Value)
        tlCntTypes.iDR = gSetCheck(RptSelCb!ckcSelC3(5).Value)
        tlCntTypes.iPI = gSetCheck(RptSelCb!ckcSelC3(6).Value)
        tlCntTypes.iPSA = gSetCheck(RptSelCb!ckcSelC3(7).Value)
        tlCntTypes.iPromo = gSetCheck(RptSelCb!ckcSelC3(8).Value)
        '12/13/18   added "Include ISCII/Create Title" option
        ilIncludeISCI = gSetCheck(RptSelCb!ckcIncludeISCI.Value)
    End If
    
    '9-11-19 use csi calendar control vs editbox
'    slStartDate = RptSelCb!edcSelCFrom.Text   'Start date
'    slEndDate = RptSelCb!edcSelCFrom1.Text   'End date
'    '3-26-11 for MG REvenue, this is the MG dates
'    slStartEDate = RptSelCb!edcSelCTo.Text   'Start date
'    slEndEDate = RptSelCb!edcSelCTo1.Text   'End date
    slStartDate = RptSelCb!CSI_CalFrom.Text     'missed start date
    slEndDate = RptSelCb!CSI_CalTo.Text         'missed end date
    slStartEDate = RptSelCb!CSI_CalFrom2.Text     'mg start date
    slEndEDate = RptSelCb!CSI_CalTo2.Text           'mg end date
    
    If ilListIndex = CNT_MGREVENUE Then             '3-27-11
        slMissedStart = slStartDate             'user entered missed start date
        slMissedEnd = slEndDate                 'user entered missed end date
        slMGStart = slStartEDate                'user entered mg start date
        slMGEnd = slEndEDate                    'user entered mg end date
        If slMissedStart = "" And slMissedEnd = "" Then
            slHdr = "Missed Dates: All:"
        ElseIf slMissedStart = "" Then          'start date blank
            slHdr = "Missed Dates: thru " & Format$(gDateValue(slMissedEnd), "m/d/yy")
        ElseIf slMissedEnd = "" Then            'end date blank
            slHdr = "Missed Dates: from " & Format$(gDateValue(slMissedStart), "m/d/yy")
        Else
            slHdr = "Missed Dates: " & Format$(gDateValue(slMissedStart), "m/d/yy") & "-" & Format$(gDateValue(slMissedEnd), "m/d/yy")
        End If
        If Not gSetFormula("DateRange", "'" & slHdr & "'") Then
            MsgBox "RptGenCb - error in DateRange", vbOKOnly, "Report Error"
            Exit Sub
        End If
        If slMGStart = "" And slMGEnd = "" Then
            slHdr = "MG Dates: All"
        ElseIf slMGStart = "" Then          'start date blank
            slHdr = "MG Dates: thru " & Format$(gDateValue(slMGEnd), "m/d/yy")
        ElseIf slMGEnd = "" Then            'end date blank
            slHdr = "MG Dates: from " & Format$(gDateValue(slMGStart), "m/d/yy")
        Else
            slHdr = "MG Dates: " & Format$(gDateValue(slMGStart), "m/d/yy") & "-" & Format$(gDateValue(slMGEnd), "m/d/yy")
        End If
        If Not gSetFormula("MGDateRange", "'" & slHdr & "'") Then
            MsgBox "RptGenCb - error in DateRange", vbOKOnly, "Report Error"
            Exit Sub
        End If
    Else
        slDateRange = "For " & slStartDate & " To " & slEndDate
        If slEndDate = "TFN" Then
            slEndDate = ""
        End If
        If slEndEDate = "TFN" Then
            slEndEDate = ""
        End If
        If (slStartDate = "") And (slEndDate = "") Then
            slDateRange = "For All Dates"
        End If
        
        '-------------------------------
        'TTP 10674 - Spot and Digital Line combo Export or report?
        If ilListIndex = CNT_SPTCOMBO And RptSelCb.rbcOutput(3) Then
            'Don't Open Crystal, we are Exporing to CSV
        Else
            If Not gSetFormula("DateRange", "'" & slDateRange & "'") Then
                MsgBox "RptGenCb - error in DateRange", vbOKOnly, "Report Error"
                Exit Sub
            End If
        End If
        
    End If
        
    If RptSelCb!rbcSelCInclude(0).Value Then
        ilShowPrice = True
    Else
        ilShowPrice = False
    End If

    ilBillStatus = 0        '3-27-11 option to show billed, unbilled or both in MGRevenue
    'if MG Revenue report, decrease the start date, and increase the end date to get all possible missed spots since
    'the base file is SDF instead of SMF
    If ilListIndex = CNT_MGREVENUE Then     '6-16-00 if MG revenue, adjust the dates
       If RptSelCb!rbcSelC9(0).Value = True Then
            ilBillStatus = 1
        ElseIf RptSelCb!rbcSelC9(1).Value = True Then
            ilBillStatus = 2
        End If
        '3-26-11 determine mg dates entered
        If slStartEDate <> "" Then      'mg startdate
            llMGStartDate = gDateValue(slStartEDate)
            slMGStartDate = Format(llMGStartDate, "m/d/yy")
        Else
            'No date entered, force it to earliest date possible
            llMGStartDate = gDateValue("1/1/1970")
        End If
        If slEndEDate <> "" Then            'mg end date
            llMGEndDate = gDateValue(slEndEDate)
            slMGEndDAte = Format(llMGEndDate, "m/d/yy")
        Else
            llMGEndDate = gDateValue("12/31/2069")
        End If
        
        If slStartDate <> "" Then       'missed date
            llStartDate = gDateValue(slStartDate)
        Else
            'No date entered, force it to earliest date possible
            llStartDate = gDateValue("1/1/1970")
        End If
        If slEndDate <> "" Then
            llEndDate = gDateValue(slEndDate)
        Else
            llEndDate = gDateValue("12/31/2069")
        End If
        
        ilSaveInclMGForRev = gSetCheck(RptSelCb!ckcSelC5(7).Value)
        RptSelCb!ckcSelC5(7).Value = vbChecked              'always include mg rate lines, filter later and not in generalized subroutine
    Else
        If ilListIndex = CNT_SPTCOMBO Then
            If RptSelCb.rbcOutput(3) Then
                'Don't Open Crystal, we are Exporing to CSV
            Else
                'Spots Combo has gross/net option
                If RptSelCb!rbcSelC7(1).Value = True Then       'gross
                    If Not gSetFormula("GrossNet", "'G'") Then
                        MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotAdvtRpt"
                        Exit Sub
                    End If
                    ilNet = False
                Else
                    If Not gSetFormula("GrossNet", "'N'") Then  'net
                        MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotAdvtRpt"
                        Exit Sub
                    End If
                    ilNet = True
                End If
            End If
        Else
            'spots by advt has gross/net option
            If RptSelCb!rbcSelC11(0).Value = True Then      'gross
                If Not gSetFormula("GrossNet", "'G'") Then
                    MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotAdvtRpt"
                    Exit Sub
                End If
                ilNet = False
            Else
                If Not gSetFormula("GrossNet", "'N'") Then  'net
                    MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotAdvtRpt"
                    Exit Sub
                End If
                ilNet = True
            End If
        End If
    End If
    
    'for MG revenue, the checkbox to include MG (spot rates) may have been altered (tested later for filtering)
    mSetCostType ilCostType             'set bit pattern of types of spots to include
    gObtainVirtVehList

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
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)

    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmVsf
        btrDestroy hmSdf
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
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmFsf
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)

    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    If ilListIndex = CNT_SPTCOMBO Then
        hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmVsf)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCif)
            btrDestroy hmAgf
            btrDestroy hmFsf
            btrDestroy hmGrf
            btrDestroy hmSmf
            btrDestroy hmVsf
            btrDestroy hmSdf
            btrDestroy hmVef
            btrDestroy hmAdf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmCif
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imCifRecLen = Len(tmCif)
        
        hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmVsf)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCif)
            ilRet = btrClose(hmCpf)
            btrDestroy hmAgf
            btrDestroy hmFsf
            btrDestroy hmGrf
            btrDestroy hmSmf
            btrDestroy hmVsf
            btrDestroy hmSdf
            btrDestroy hmVef
            btrDestroy hmAdf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmCif
            btrDestroy hmCpf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imCpfRecLen = Len(tmCpf)
    
        hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmVsf)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCif)
            ilRet = btrClose(hmCpf)
            ilRet = btrClose(hmMnf)
            btrDestroy hmAgf
            btrDestroy hmFsf
            btrDestroy hmGrf
            btrDestroy hmSmf
            btrDestroy hmVsf
            btrDestroy hmSdf
            btrDestroy hmVef
            btrDestroy hmAdf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmCif
            btrDestroy hmCpf
            btrDestroy hmMnf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imMnfRecLen = Len(tmMnf)
                
        hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRdf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmGrf)
            btrDestroy hmRdf
            btrDestroy hmFsf
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmAdf
            btrDestroy hmVef
            btrDestroy hmSdf
            btrDestroy hmSmf
            btrDestroy hmCff
            btrDestroy hmGrf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imRdfRecLen = Len(tmRdf)

        hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRdf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCxf)
            btrDestroy hmRdf
            btrDestroy hmFsf
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmAdf
            btrDestroy hmVef
            btrDestroy hmSdf
            btrDestroy hmSmf
            btrDestroy hmCff
            btrDestroy hmGrf
            btrDestroy hmCxf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imRdfRecLen = Len(tmRdf)
    End If
    
    '--------------------------------------------------------
    'TTP 10674 - Spot and Digital Line combo Export or report?
    If ilListIndex = CNT_SPTCOMBO And RptSelCb.rbcOutput(3) Then
        'Generate Export Filename
        slFileName = ""
        slRepeat = "A"
        'Get Client Name
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
        
        Do
            ilRet = 0
            slFileName = "SpotCbo "
            slFileName = slFileName & Format(gNow, "mmddyy")
            slFileName = slFileName & slRepeat

            slFileName = slFileName & " " & gFileNameFilter2(Trim$(smClientName))
            slFileName = slFileName & ".csv"
            'Check if exists, make new character
            ilRet = gFileExist(sgExportPath & slFileName)
            If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
                slRepeat = Chr(Asc(slRepeat) + 1)
            End If
        Loop While ilRet = 0
        RptSelCb.edcFileName.Text = slFileName
        
        'Create File
        ilRet = gFileOpen(sgExportPath & slFileName, "OUTPUT", hmExport)
        If ilRet <> 0 Then
            MsgBox "Error writing file:" & sgExportPath & slFileName & vbCrLf & "Error:" & ilRet & " - " & Error(ilRet)
            Close #hmExport
            Exit Sub
        End If
    
        'Write CSV Header
        mWriteExportHeader ilListIndex, hmExport
    End If
    
    ReDim tmSelChf(0 To 0) As Long
    ReDim tmSelAgf(0 To 0) As Integer
    ReDim tmSelSlf(0 To 0) As Integer
    ReDim tmSelAdf(0 To 0) As Integer    'array of advt to select matching network feeds by adv
    ReDim tmSelVef(0 To 0) As Integer

    ilWhichKey = INDEXKEY1              'sdf default key (vef,date,time)

    If RptSelCb!rbcSelCSelect(0).Value Then 'Advertiser/Contracts (vs slsp or agy)
        If RptSelCb!ckcAll.Value = vbChecked Then       'all advt?
            ilSelType = 3
        Else
            ilSelType = 0                               'selective adv, build array of contr codes selected
            ilLocal = False
            ilFeed = False
            'if more than 1 advt selected, use the normal vehicle/date key; otherwise use Advt/date key
            
            For illoop = 0 To RptSelCb!lbcSelection(0).ListCount - 1 Step 1
                If RptSelCb!lbcSelection(0).Selected(illoop) Then
                    slNameCode = RptSelCb!lbcCntrCode.List(illoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmSelChf(UBound(tmSelChf)) = Val(slCode)
                    If Val(slCode) = 0 Then
                        ilFeed = True
                    Else
                        ilLocal = True
                    End If
                    ReDim Preserve tmSelChf(0 To UBound(tmSelChf) + 1) As Long
                End If
            Next illoop
        End If
        'build array of selected advt
        If RptSelCb!lbcSelection(5).SelCount = 1 Then
            ilWhichKey = INDEXKEY7              'single advt selection, use key7 for faster access
        End If
        For illoop = 0 To RptSelCb!lbcSelection(5).ListCount - 1 Step 1
            If RptSelCb!lbcSelection(5).Selected(illoop) Then
                slNameCode = tgAdvertiser(illoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmSelAdf(UBound(tmSelAdf)) = Val(slCode)
                ReDim Preserve tmSelAdf(0 To UBound(tmSelAdf) + 1) As Integer
            End If
        Next illoop

    ElseIf RptSelCb!rbcSelCSelect(1).Value Then 'Agency
        ilSelType = 1
        gObtainAgyAdvCodes ilIncludeCodes, tmSelAgf(), 1, RptSelCb    'use lbcselection(1) for list box of direct & agencies
    ElseIf RptSelCb!rbcSelCSelect(2).Value Then 'Salesperson
        ilSelType = 2
        For illoop = 0 To RptSelCb!lbcSelection(2).ListCount - 1 Step 1
            If RptSelCb!lbcSelection(2).Selected(illoop) Then
                slNameCode = tgSalesperson(illoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                tmSelSlf(UBound(tmSelSlf)) = Val(slCode)
                ReDim Preserve tmSelSlf(0 To UBound(tmSelSlf) + 1) As Integer
            End If
        Next illoop
    End If
    tmAgf.iCode = 0
    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0
    ReDim tmPLSdf(0 To 0) As SPOTTYPESORT

    For ilVehicle = 0 To RptSelCb!lbcSelection(6).ListCount - 1 Step 1      'loop thru vehicles
        If RptSelCb!lbcSelection(6).Selected(ilVehicle) Then                'vehicle selected?
            slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelCb!lbcCSVNameCode.List(ilVehicle)
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
            tmSelVef(UBound(tmSelVef)) = Val(slCode)
            ReDim Preserve tmSelVef(0 To UBound(tmSelVef) + 1) As Integer
            '5-9-11 Remove all the invalid bb spots that doesnt belong
            ilGameSelect = 0
            ilLineSelect = 0
            llChfSelect = 0
            ilRet = gRemoveBBSpots(hmSdf, ilVefCode, ilGameSelect, slStartDate, slEndDate, llChfSelect, ilLineSelect)
        End If
    Next ilVehicle
    
    '----------------------------------------------------------------------------------------------------------
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    '----------------------------------------------------------------------------------------------------------
    'The Spots Combo report can include digital lines and allows selecting Digital lines on and off
    If RptSelCb.ckcSelDigital.Value = vbChecked And ilListIndex = CNT_SPTCOMBO Then
        '------------------------------------------
        'Determine SelectionIndex: 0=Contracts, 1=Agency, 5=Advertiser, 6=Vehicles
        If RptSelCb.rbcSelCSelect(0).Value Then 'Advertiser
            If UBound(tmSelVef) > 0 And tmSelVef(0) <> 0 Then
                ilSelectionIndex = 6
            End If
            If UBound(tmSelAdf) > 0 And tmSelAdf(0) <> 0 And RptSelCb.ckcAll.Value = vbUnchecked Then
                ilSelectionIndex = 5
            End If
            If UBound(tmSelChf) > 0 And tmSelChf(0) <> 0 And RptSelCb.ckcAll.Value = vbUnchecked Then
                ilSelectionIndex = 0
            End If
        ElseIf RptSelCb.rbcSelCSelect(1).Value Then  'Agency
            If UBound(tmSelVef) > 0 And tmSelVef(0) <> 0 Then
                ilSelectionIndex = 6
            End If
        End If
        
        '------------------------------------------
        'Loop through user's Selection of Vehicles, Contracts, or Advertisers
        For ilLoopOnKey = 0 To RptSelCb!lbcSelection(ilSelectionIndex).ListCount - 1
            'SelectionIndex: 0=Contracts, 1=Agency, 5=Advertiser, 6=Vehicles
            If RptSelCb!lbcSelection(ilSelectionIndex).Selected(ilLoopOnKey) Then
                If ilSelectionIndex = 6 Then
                    slNameCode = tgCSVNameCode(ilLoopOnKey).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ElseIf ilSelectionIndex = 5 Then
                    slNameCode = tgAdvertiser(ilLoopOnKey).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ElseIf ilSelectionIndex = 0 Then
                    slNameCode = RptSelCb!lbcSelection(ilSelectionIndex).List(ilLoopOnKey)
                    slCode = Val(RptSelCb!lbcSelection(ilSelectionIndex).List(ilLoopOnKey))
                End If
                llWhichCode = Val(slCode)
                
                'Saves time by not looking for PCF records on non-digital vehicles
                If ilSelectionIndex = 6 Then
                    'By Vehicle - check that this is a digital vehicle
                    ilVpf = gBinarySearchVpf(Val(slCode))
                    If ilVpf <> -1 Then
                        If tgVpf(ilVpf).sGMedium <> "P" Then
                            GoTo skipVehicle
                        End If
                    End If
                End If
                
                '------------------------------------------
                'Get PCF records for selection
                ReDim tmPLPcf(0 To 0) As PCFTYPESORT
                mObtainSelPcf ilSelectionIndex, llWhichCode, slStartDate, slEndDate, tmSelAgf(), tmSelAdf(), tmSelChf(), tlCntTypes, llTestCntrNo
                '------------------------------------------
                'loop through all the PCF records that we got in mObtainSelPcf
                For llPCFIndex = LBound(tmPLPcf) To UBound(tmPLPcf) - 1
                    tmPcf = tmPLPcf(llPCFIndex).tPcf
                    smFormulaComment = ""
                    'Get Digital Line Date range
                    gUnpackDateLong tmPcf.iStartDate(0), tmPcf.iStartDate(1), llCPMStartDate
                    gUnpackDateLong tmPcf.iEndDate(0), tmPcf.iEndDate(1), llCPMEndDate
                    If llCPMStartDate <= llCPMEndDate Then '1. For a digital line that has at least one day in the report period
                        '------------------------------------------
                        '1. For a digital line that has at least one day in the report period, determine the number of broadcast months that it's running for
                        ilRemainMonths = gObtainMonthsOfCPMID(llCPMStartDate, llCPMEndDate, tmPLPcf(llPCFIndex).sCalType)
                        Debug.Print "dSpotAdvRpt(DigitalLine) Cntr:" & tmPLPcf(llPCFIndex).lCntrNo & ", " & tmPLPcf(llPCFIndex).tPcf.iPodCPMID & " :" & Format(llCPMStartDate, "ddddd") & " - " & Format(llCPMEndDate, "ddddd") & "; Months:" & ilRemainMonths & ", Days:" & DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(llCPMEndDate, "ddddd")) + 1
                        
                        '------------------------------------------
                        'Get Monthly Start Dates lmStartDates(), so we can know how many days per Month this Line is active for
                        slDate = Format(llCPMStartDate, "ddddd")
                        If tmPLPcf(llPCFIndex).sCalType = "S" Then
                            mBuildMonthlyDates slDate, 4, ilRemainMonths + 3
                        Else
                            mBuildMonthlyDates slDate, 1, ilRemainMonths + 3
                        End If
                        
                        '------------------------------------------
                        '3.a. Determine the invoice amount for each previously billed broadcast month.
                        slRvfStart = tmPLPcf(llPCFIndex).sCntrStartDate
                        If tmPLPcf(llPCFIndex).sCalType = "S" Then
                            slLastBilledDate = Format(llLastBilledStd, "ddddd")
                        Else
                            slLastBilledDate = Format(llLastBilledCal, "ddddd")
                        End If
                        slRvfEnd = slLastBilledDate
                        If slRvfStart <> "" Then
                            'Fix per Jason Email - v81 TTP 10804 testing Fri 8/11/23 9:53 AM
                            'If Contract Ends before Last billed date, just pull rvf for the length of the contract
                            'If DateValue(tmPLPcf(llPCFIndex).sCntrEndDate) < DateValue(slLastBilledDate) Then
                            '    slRvfEnd = tmPLPcf(llPCFIndex).sCntrEndDate
                            'End If
                            If slLastRvf <> tmPLPcf(llPCFIndex).lCntrNo & "," & slRvfStart & "," & slRvfEnd Then
                                ReDim tlRvf(0 To 0) As RVF
                                If slRvfStart <> "" And slRvfEnd <> "" Then
                                    'Get the Receivables
                                    ilRet = gObtainPhfRvfbyCntr(RptSelCb, tmPLPcf(llPCFIndex).lCntrNo, slRvfStart, slRvfEnd, tlTranTypes, tlRvf())
                                    'Debug.Print " - dSpotAdvRpt Rvf:" & tmPLPcf(llPCFIndex).lCntrNo & " for:" & slRvfStart & " - " & slRvfEnd & "; " & UBound(tlRvf())
                                    slLastRvf = tmPLPcf(llPCFIndex).lCntrNo & "," & slRvfStart & "," & slRvfEnd
                                Else
                                    slLastRvf = ""
                                End If
                            End If
                        End If
                        
                        '------------------------------------------
                        'put the Receivables into monthly buckets
                        llTotalReceivedAmount = 0
                        'TTP 10961 - Spot and Digital Line combo report: show more information about digital line invoice adjustments on special export version
                        ReDim slTranComments(0)
                        For illoop = 0 To UBound(tlRvf()) - 1
                            If tlRvf(illoop).lCntrNo = tmPLPcf(llPCFIndex).lCntrNo Then
                                If gObtainPcfCPMID(tlRvf(illoop).lPcfCode) = tmPLPcf(llPCFIndex).tPcf.iPodCPMID Then
                                    gPDNToLong tlRvf(illoop).sGross, llAmount
                                    If llAmount <> 0 Then
                                        'Which Month was this Received for?  Store RVF $ into different month buckets
                                        'Debug.Print " - PCFCode:" & tlRvf(illoop).lPcfCode & ", matches for CntrNo:"; tmPLPcf(llPCFIndex).lCntrNo & ", Line:" & tmPLPcf(llPCFIndex).tPcf.iPodCPMID & " for amount:" & llAmount / 100
                                        For ilNoDays = 1 To ilRemainMonths + 1 'RE: RAB cal spots discrepancy Thu 6/15/23 9:59 AM (Issue 12)
                                            If tgSpf.sSEnterAgeDate = "E" Then
                                                'use entered date or ageing date
                                                gUnpackDate tlRvf(illoop).iTranDate(0), tlRvf(illoop).iTranDate(1), slDate
                                                If tmPLPcf(llPCFIndex).sCalType = "S" Then
                                                    slDate = gObtainEndStd(slDate)
                                                End If
                                            Else
                                                'use Age Period
                                                slDate = Trim$(str$(tlRvf(illoop).iAgePeriod) & "/15/" & Trim$(str$(tlRvf(illoop).iAgingYear)))
                                                slDate = gObtainEndStd(slDate)
                                            End If
                                            If DateValue(slDate) >= DateValue(Format(lmStartDates(ilNoDays), "ddddd")) And DateValue(slDate) < DateValue(Format(lmStartDates(ilNoDays + 1), "ddddd")) Then
                                                tmPLPcf(llPCFIndex).lReceivables(ilNoDays) = tmPLPcf(llPCFIndex).lReceivables(ilNoDays) + llAmount
                                                llTotalReceivedAmount = llTotalReceivedAmount + llAmount
                                                Debug.Print " - Cntr:" & tlRvf(illoop).lCntrNo & ", TranType:" & tlRvf(illoop).sTranType & ", Inv#:" & tlRvf(illoop).lInvNo & ", line:" & tmPLPcf(llPCFIndex).tPcf.iPodCPMID & ", Received " & llAmount / 100 & " in Month:" & ilNoDays
                                                
                                                'TTP 10961 - Spot and Digital Line combo report: show more information about digital line invoice adjustments on special export version
                                                If RptSelCb.rbcOutput(3) And RptSelCb.ckcSelDigitalComments.Value = vbChecked Then
                                                    'Write Transaction Comments to an array so they can follow the detail record in the Export file
                                                    slTranComments(UBound(slTranComments)) = " -> Date:" & slDate & "; TranType:" & tlRvf(illoop).sTranType & "; Inv#:" & tlRvf(illoop).lInvNo & "; Amount " & Format(llAmount / 100, "0.00")
                                                    ReDim Preserve slTranComments(UBound(slTranComments) + 1)
                                                End If
                                            End If
                                        Next ilNoDays
                                    End If
                                End If
                            End If
                        Next illoop
                        
                        '------------------------------------------
                        'For billed months
                        '4.a. Subtract the total amount billed so far from the total line cost to get the remaining unbilled amount
                        llRemainingAmount = tmPcf.lTotalCost - llTotalReceivedAmount
                        '4.b. Using the line dates, determine how many remaining unbilled line days there are
                        '4.d. Using the report dates, determine how many days total are remaining in the unbilled report period and multiply that by the remaining unbilled daily average to get the unbilled report total
                        ilDaysBilled = 0
                        'ilDaysUnbilled = DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(llCPMEndDate, "ddddd")) + 1
                        'If llTotalReceivedAmount <> 0 Then 'RE: RAB cal spots discrepancy Thu 6/15/23 9:59 AM (Issue 11)
                        If tmPLPcf(llPCFIndex).sCalType = "S" Then
                            'TTP 10823 - Spot and Digital Line Combo report - daily average off by one day for flat rate line
                            'ilDaysUnbilled = DateDiff("d", Format(IIF(llCPMStartDate > llLastBilledStd, llCPMStartDate, llLastBilledStd), "ddddd"), Format(llCPMEndDate, "ddddd")) + 1 'Combo report update 6-29-23 (Issue 13)
                            ilDaysUnbilled = DateDiff("d", Format(IIF(llCPMStartDate > llLastBilledStd, llCPMStartDate, llLastBilledStd + 1), "ddddd"), Format(llCPMEndDate, "ddddd")) + 1 'Combo report update 6-29-23 (Issue 13)
                            ilDaysBilled = DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(IIF(llCPMEndDate < llLastBilledStd, llCPMEndDate, llLastBilledStd), "ddddd")) + 1
                        Else
                            'TTP 10823 - Spot and Digital Line Combo report - daily average off by one day for flat rate line
                            'ilDaysUnbilled = DateDiff("d", Format(IIF(llCPMStartDate > llLastBilledCal, llCPMStartDate, llLastBilledCal), "ddddd"), Format(llCPMEndDate, "ddddd")) + 1 'Combo report update 6-29-23 (Issue 13)
                            ilDaysUnbilled = DateDiff("d", Format(IIF(llCPMStartDate > llLastBilledCal, llCPMStartDate, llLastBilledCal + 1), "ddddd"), Format(llCPMEndDate, "ddddd")) + 1 'Combo report update 6-29-23 (Issue 13)
                            ilDaysBilled = DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(IIF(llCPMEndDate < llLastBilledCal, llCPMEndDate, llLastBilledCal), "ddddd")) + 1
                        End If
                        'End If
                        If ilDaysUnbilled < 0 Then ilDaysUnbilled = 0
                        If ilDaysBilled < 0 Then ilDaysBilled = 0
                        'Debug.Print " - BilledDays:" & ilDaysBilled & ", UnbilledDays:" & ilDaysUnbilled & " = " & ilDaysBilled + ilDaysUnbilled & " total days"
                        'Debug.Print " - LineTotal:" & tmPLPcf(llPCFIndex).tPcf.lTotalCost / 100 & ", RemainingAmount:" & llRemainingAmount / 100
                        smFormulaComment = "BilledAmt:" & llTotalReceivedAmount / 100 & "; Remaining:" & llRemainingAmount / 100
                        
                        '------------------------------------------
                        'Total each Month's Digital line values
                        dlTotalAmount = 0
                        dlDailyAmt = 0
                        For illoop = 1 To ilRemainMonths + 1
                            '2. Determine the number of days it's running in each broadcast month that the line runs for
                            ilNoDays = mNumberOfDaysRunningInMonth(lmStartDates(illoop), lmStartDates(illoop + 1) - 1, llCPMStartDate, llCPMEndDate)
                            If ilNoDays <> 0 Then
                                'Debug.Print " - has " & ilNoDays & " days in month:" & illoop & ", received:" & Format(tmPLPcf(llPCFIndex).lReceivables(illoop) / 100, "#.00")
                                'If tmPLPcf(llPCFIndex).lReceivables(illoop) <> 0 Then 'Might have billed $0.00
                                If lmStartDates(illoop) <= IIF(tmPLPcf(llPCFIndex).sCalType = "S", llLastBilledStd, llLastBilledCal) Then 'per Jason $0.00 bonus items can be billed
                                    'this month was Billed, the billed amount might be $0.00
                                    dlDailyAmt = (tmPLPcf(llPCFIndex).lReceivables(illoop) / 100) / ilNoDays
                                Else
                                    '4.c. Divide the remaining unbilled amount by the remaining unbilled line days to get the remaining unbilled daily average
                                    If llRemainingAmount = 0 Or ilDaysUnbilled = 0 Then  'RE: RAB cal spots discrepancy Thu 6/15/23 9:59 AM (Issue 11)
                                        dlDailyAmt = 0
                                    Else
                                        dlDailyAmt = (llRemainingAmount / 100) / ilDaysUnbilled
                                    End If
                                End If
                                '3.c. For each billed broadcast month, determine how many days the report period is covered by the line, and multiply that by the billed daily average for that month to get the billed report total
                                ilNoDays = mNumberOfDaysRunningInMonth(lmStartDates(illoop), lmStartDates(illoop + 1) - 1, llCPMStartDate, llCPMEndDate, slStartDate, slEndDate)
                                'Debug.Print " - runs for " & ilNoDays & " days in month:" & illoop & " @ " & Format(dlDailyAmt, "#.00")
                                dlTotalAmount = dlTotalAmount + (dlDailyAmt * ilNoDays)
                                If ilNoDays <> 0 Then
                                    If smFormulaComment <> "" Then smFormulaComment = smFormulaComment & "; "
                                    smFormulaComment = smFormulaComment & ilNoDays & " days in " & MonthName(Month(DateAdd("d", 15, Format(lmStartDates(illoop), "ddddd"))), True) & " @" & dlDailyAmt
                                End If
                            End If
                        Next illoop

                        '-------------------------------
                        'TTP 10674 - Spot and Digital Line combo Export or report?
                        If RptSelCb.rbcOutput(3) Then
                            'Write to Export file
                            '--------------------
                            '1. ContractNo,
                            tlExp.lChfCode = tmPcf.lChfCode
                            tlExp.lContract_Number = tmPLPcf(llPCFIndex).lCntrNo
                            '2. ExtContractNo,
                            tlExp.lExternal_Version_Number = tmPLPcf(llPCFIndex).lExtCntrNo 'ExtCntrNo
                            '3. LineType
                            tlExp.sLine_Type = "Digital"
                            '4. ContractType
                            tlExp.sContract_Type = tmPLPcf(llPCFIndex).sContractType
                            '5. Agency,
                            tlExp.iAgency_ID = tmPLPcf(llPCFIndex).iAgfCode
                            '6. Advertiser
                            tlExp.iAdvertiser_ID = tmPLPcf(llPCFIndex).iAdfCode
                            '7. Product
                            tlExp.sProduct_Name = tmPLPcf(llPCFIndex).sProduct
                            '8. Line
                            tlExp.iPkLineNo = tmPLPcf(llPCFIndex).tPcf.iPodCPMID
                            '9. Vehicle
                            tlExp.iVehicle_ID = tmPLPcf(llPCFIndex).tPcf.iVefCode
                            '10. Day (N/A)
                            tlExp.sDay = ""
                            '11. SpotAirDate (N/A)
                            tlExp.sSpotAirDate = ""
                            '12. SpotAirTime (N/A)
                            tlExp.sSpotAirTime = ""
                            '13. SpotAudioType (N/A)
                            tlExp.sSpotAudioType = ""
                            '14. ISCIcode (N/A)
                            tlExp.sISCIcode = ""
                            '15. Len (N/A)
                            tlExp.iSpot_Length = tmPLPcf(llPCFIndex).tPcf.iLen 'Boostr Phase 2: Spot and Digital Line Combo report: show length for digital lines
                            '16. DigitalLineStartDate
                            tlExp.sLine_Start_Date = Format(llCPMStartDate, "ddddd")
                            '17. DigitalLineEndDate
                            tlExp.sLine_End_Date = Format(llCPMEndDate, "ddddd")
                            '18. Net - will be calculated in mWriteExportFile
                            '19. Gross
                            tlExp.dTotal_Gross = Format(dlTotalAmount, "0.00") * 100 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 9)
                            '20. Spot status: the status of the spot (scheduled, missed, hidden, cancelled)
                            tlExp.sStatus = ""
                            '21. Spot price type: the spot price type from the contract line (charge, N/C, MG, ADU, Recap, Spinoff, Bonus, Package.)
                            tlExp.sPrice_Type = ""
                            '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
                            tlExp.sDaypart = tmPLPcf(llPCFIndex).tPcf.iRdfCode
                            '23. Line comment: the line comment from the spot line or digital line
                            tlExp.lLineCxfComment = tmPLPcf(llPCFIndex).tPcf.lCxfCode
                            '24. Digital Line calculation note
                            tlExp.sFormulaComment = smFormulaComment
                            mWriteExportFile ilListIndex, hmExport, hmCHF, hmAgf, hmAdf, hmClf, hmRdf, tlExp
                            
                            'TTP 10961 - Spot and Digital Line combo report: show more information about digital line invoice adjustments on special export version
                            If RptSelCb.rbcOutput(3) And RptSelCb.ckcSelDigitalComments.Value = vbChecked Then
                                For illoop = 0 To UBound(slTranComments) - 1
                                    If slTranComments(illoop) <> "" Then mWriteExportFileComment ilListIndex, hmExport, slTranComments(illoop)
                                Next illoop
                            End If
                        Else
                            'Write line to GRF
                            tmGrf.lGenTime = lgNowTime                          'Gen Date/Time
                            tmGrf.iGenDate(0) = igNowDate(0)
                            tmGrf.iGenDate(1) = igNowDate(1)
                            tmGrf.lChfCode = tmPcf.lChfCode                     'Contract code
                            tmGrf.iCode2 = tmPcf.iPodCPMID                      'Pod Line #
                            tmGrf.iVefCode = tmPcf.iVefCode                     'Vehicle
                            tmGrf.iPerGenl(0) = tmPcf.iLen                      'Boostr Phase 2: Spot and Digital Line Combo report: show length for digital lines
                            tmGrf.iPerGenl(2) = 0                               'default to not a mg or outside (vehicle for mg or out)
                            tmGrf.iDateGenl(0, 0) = tmPcf.iStartDate(0)         'Digital Start Date
                            tmGrf.iDateGenl(1, 0) = tmPcf.iStartDate(1)
                            tmGrf.iDateGenl(0, 1) = tmPcf.iEndDate(0)           'Digital End Daet
                            tmGrf.iDateGenl(1, 1) = tmPcf.iEndDate(1)
                            tmGrf.sGenDesc = ""                                 '
                            tmGrf.sBktType = "D"                                'Indicate Digital - this is used to determine if it's a spot or digital in Crystal
                            tmGrf.lDollars(0) = dlTotalAmount * 100             'Reported Line Amount
                            tmGrf.sDateType = tmPcf.sPriceType                  'C=CPM, F=Flat Rate
                            tmGrf.iPerGenl(3) = 0                               'if cancelled or missed, Origin sch vef has been set with ordered vehicle
                            tmGrf.iPerGenl(5) = IIF(ilDaysBilled <> 0, 1, 0)    'tell crystal its been invoiced
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                    End If
                Next llPCFIndex 'next Digital Line
            End If 'is Vehicle, Contract, or Advertiser Selected?
skipVehicle:
        Next ilLoopOnKey 'End Loop of user's Selected of Vehicles, Contracts, or Advertisers
    End If
    
    '------------------------------------------------------------
    'Spots
    'The Spots Combo report allows selecting Spots on and off, but all other reports need to select spots
    If ilWhichKey = INDEXKEY1 Then
        ilSelectionIndex = 6
    Else
        ilSelectionIndex = 5
    End If
    If RptSelCb.ckcSelSpots.Value = vbChecked Or ilListIndex <> CNT_SPTCOMBO Then
        For ilLoopOnKey = 0 To RptSelCb!lbcSelection(ilSelectionIndex).ListCount - 1
            If RptSelCb!lbcSelection(ilSelectionIndex).Selected(ilLoopOnKey) Then
                If ilSelectionIndex = 6 Then
                    slNameCode = tgCSVNameCode(ilLoopOnKey).sKey
                Else
                    slNameCode = tgAdvertiser(ilLoopOnKey).sKey
                End If
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                
                ReDim tmPLSdf(0 To 0) As SPOTTYPESORT   '8-23-00
                
                '3-26-11   mg date selectivity added
                If ilListIndex = CNT_MGREVENUE Then
                    mObtainSelSdf ilWhichKey, ilVefCode, slMGStartDate, slMGEndDAte, "", "", ilSelType, ilCostType, ilLocal, ilFeed, ilIncludeCodes, tmSelAgf(), ilUseAcqRate, tlCntTypes
                Else
                    '11-11-15 if using acq and its a barter and using vehicle search, get the sdf by vehicle
                    'OR, if using Acq and its single advt, get the sdf by advt
                    'OR, if not using acq get the sdf by vehicle
                    'ilVefcode that is sent may be the advt code if single advt
                    If ((ilUseAcqRate = True) And (gIsOnInsertions(ilVefCode) = True) And ilWhichKey = INDEXKEY1) Or ((ilUseAcqRate = True) And ilWhichKey = INDEXKEY7) Or (ilUseAcqRate = False) Then
                        mObtainSelSdf ilWhichKey, ilVefCode, slStartDate, slEndDate, slStartEDate, slEndEDate, ilSelType, ilCostType, ilLocal, ilFeed, ilIncludeCodes, tmSelAgf(), ilUseAcqRate, tlCntTypes, llTestCntrNo
                    End If
                End If
                
            'End If         '6-30-00    gather and write btr records for one vehicle at a time to avoid > 32000 spots per vehicle
        'Next ilVehicle     '6-30-00
    
        '4-22-05 reinstate the sorting of the spots array
                llUpper = UBound(tmPLSdf)
                If llUpper > 0 Then
                    ArraySortTyp fnAV(tmPLSdf(), 0), llUpper, 0, LenB(tmPLSdf(0)), 0, Len(tmPLSdf(0).sKey), 0
                End If
    
                For llIndex = LBound(tmPLSdf) To UBound(tmPLSdf) - 1
                    tmSdf = tmPLSdf(llIndex).tSdf
                    blValidSelection = True
                    If RptSelCb!rbcSelCSelect(0).Value = True And RptSelCb!ckcAll.Value = vbUnchecked Then      'advt selection with selective advertisers:  need to check the contracts selected
                        'test for contracts selected
                        blValidSelection = False
                        For illoop = LBound(tmSelChf) To UBound(tmSelChf) - 1
                            If tmSdf.lChfCode = tmSelChf(illoop) Then
                                blValidSelection = True
                                Exit For
                            End If
                        Next illoop
                    End If
                    
                    If (ilUseAcqRate = True) And (gIsOnInsertions(tmSdf.iVefCode) = False) Then      'if using acq cost and its not a barter, ignore it.  Single selection advt didnt have this test in above test; only was tested for the vehicle
                        blValidSelection = False
                    End If
                    
                    If blValidSelection Then
                        If tmAdf.iCode <> tmSdf.iAdfCode Then
                            tmAdfSrchKey.iCode = tmSdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        
                        'TTP 10674 - Spot and Digital Line combo report: get ISCI
                        If ilListIndex = CNT_SPTCOMBO Then
                            mObtainCopy slProduct, slZone, slCart, slISCI
                        End If
                        tmGrf.lDollars(1) = tmSdf.lFsfCode          'network feed code
                        tmGrf.lChfCode = tmSdf.lChfCode             'contr code
                        tmGrf.iCode2 = tmSdf.iLineNo
                        tmGrf.iVefCode = tmSdf.iVefCode
                        tmGrf.lDollars(8) = tmSdf.lCopyCode              '12-26-18 like to CIF for ISCI/Creative Title
    
                        If tmSdf.sXCrossMidnight = "Y" Then         '5-24-06
                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                            gPackDateLong llDate + 1, ilDate(0), ilDate(1)
                            tmGrf.iStartDate(0) = ilDate(0)
                            tmGrf.iStartDate(1) = ilDate(1)
                        Else
                            tmGrf.iStartDate(0) = tmSdf.iDate(0)
                            tmGrf.iStartDate(1) = tmSdf.iDate(1)
                        End If
                        gUnpackDateLong tmGrf.iStartDate(0), tmGrf.iStartDate(1), tmGrf.lCode4
                        tmGrf.iTime(0) = tmSdf.iTime(0)
                        tmGrf.iTime(1) = tmSdf.iTime(1)
                        tmGrf.iPerGenl(0) = tmSdf.iLen
                        tmGrf.iPerGenl(2) = 0                   'default to not a mg or outside (vehicle for mg or out)
                        tmSmf.iOrigSchVef = tmSdf.iVefCode
                        'TTP 10674 - Spot and Digital Line combo report: get ISCI
                        If ilListIndex = CNT_SPTCOMBO Then
                            tmGrf.lDollars(2) = 0
                        End If
                        If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                            '11-30-04 change to access smf by key2 instead of key0 for speed
                            tmSmfSrchKey2.lCode = tmSdf.lCode
                            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        End If
                        If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                            tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        If tmSdf.sSchStatus = "S" Then
                            slStr = ""
                        ElseIf tmSdf.sSchStatus = "M" Then
                            slStr = "Missed"
                            'TTP 10674 - Spot and Digital Line combo report: get ISCI
                            If ilListIndex = CNT_SPTCOMBO Then
                               tmGrf.lDollars(2) = 1 'Missed Indicator
                            End If
                        ElseIf tmSdf.sSchStatus = "R" Then
                            slStr = "Ready"
                        ElseIf tmSdf.sSchStatus = "U" Then
                            slStr = "UnSched"
                        ElseIf tmSdf.sSchStatus = "G" Then
                            If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                                slStr = "Makegood"
                                tmGrf.iPerGenl(2) = tlVef.iCode
                            Else
                                slStr = "Makegood"
                            End If
                        ElseIf tmSdf.sSchStatus = "O" Then
                            If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                                If tmSdf.sSpotType = "X" Then
                                    slStr = ""
                                Else
                                    slStr = "Outside"
                                    tmGrf.iPerGenl(2) = tlVef.iCode
                                End If
                            Else
                                If tmSdf.sSpotType = "X" Then
                                    slStr = ""
                                Else
                                    slStr = "Outside"
                                End If
                            End If
                        ElseIf tmSdf.sSchStatus = "C" Then
                            slStr = "Cancelled"
                            'TTP 10674 - Spot and Digital Line combo report: get ISCI
                            If ilListIndex = CNT_SPTCOMBO Then
                               tmGrf.lDollars(2) = 1 'Missed Indicator
                            End If
                        ElseIf tmSdf.sSchStatus = "H" Then
                            slStr = "Hidden"
                            'TTP 10674 - Spot and Digital Line combo report: get ISCI
                            If ilListIndex = CNT_SPTCOMBO Then
                               tmGrf.lDollars(2) = 1 'Missed Indicator
                            End If
                        ElseIf tmSdf.sSchStatus = "A" Then
                            slStr = "On Alt"
                        ElseIf tmSdf.sSchStatus = "B" Then
                            slStr = "On Alt & MG"
                        End If
                        tmGrf.sGenDesc = Trim$(slStr)
            
                        'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
                        If ilListIndex = CNT_SPTCOMBO Then
                            tmGrf.sGenDesc = Trim$(slISCI)
                        End If
                        
                        tmGrf.sBktType = tmSdf.sSpotType
                        tmGrf.sDateType = tmSdf.sPriceType
                        slStr = Trim$(tmPLSdf(llIndex).sCostType)
                        slSpotPriceType = Trim(slStr) 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 6)
                        tmGrf.lDollars(0) = 0
            
                        If InStr(slStr, ".") <> 0 Then          'its an amount, not N/c or any other text
                            slSpotPriceType = "Charge" 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 6)
                            tmGrf.lDollars(0) = gStrDecToLong(slStr, 2) 'convert string decimal to long value
                            'if requested Gross or its a feed spot, dont calc net
                            'If (Not ilNet) Or (tmSdf.lCode = 0) Then          'show gross
                            If (Not ilNet) Or (tmSdf.lCode = 0) Or (ilListIndex = CNT_SPTCOMBO) Then          'show gross, TTP 10674 in CNT_SPTCOMBO crystal already has connection to AGF table and a Formula for {@GrossNet} no reason to access this twice
                                ilCommPct = 10000                'no commission
                            Else
                                ilCommPct = 10000               '9-7-10 assume no comm, check for direct if net since no comm taken for directs
                                'need to access contract to see if this contract is commissionable
                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                If tmChf.iAgfCode > 0 Then      'agency exists; otherwise no comm
                                    ilCommPct = 8500         'default to commissionable if no agency found
                                    'see what the agency comm is defined as
                                    tmAgfSrchKey.iCode = tmChf.iAgfCode
                                    ilRet = btrGetEqual(hmAgf, tmAgf, Len(tmAgf), tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                                    If ilRet = BTRV_ERR_NONE Then
                                        ilCommPct = (10000 - tmAgf.iComm)
                                    End If          'ilret = btrv_err_none
                                End If
                            End If
            
                            If ilNet Then       'first get the net value if applicable
                                If ilUseAcqRate = True Then
                                    ilAcqCommPct = 0
                                    blAcqOK = gGetAcqCommInfoByVehicle(tmSdf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                    ilAcqCommPct = gGetEffectiveAcqComm(tmGrf.lCode4, ilAcqLoInx, ilAcqHiInx)
                                    gCalcAcqComm ilAcqCommPct, tmGrf.lDollars(0), llAcqNet, llAcqComm
                                    tmGrf.lDollars(0) = llAcqNet
                                Else
                                    slAmount = gLongToStrDec(tmGrf.lDollars(0), 2)
                                    slSharePct = gIntToStrDec(ilCommPct, 4)
                                    slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                                    slStr = gRoundStr(slStr, ".01", 2)
                                    tmGrf.lDollars(0) = gStrDecToLong(slStr, 2) 'adjusted net
                                End If
                            End If
                            
                        ElseIf slStr = Trim$("ADU") Then
                            tmGrf.sDateType = "A"                'flag ADU spot
                            slSpotPriceType = "ADU"
                        
                        ElseIf slStr = Trim$("Bonus") Then
                            tmGrf.sDateType = "B"
                            slSpotPriceType = "Bonus"
                        
                        ElseIf slStr = Trim$("Feed") Then    'network feed spot
                            tmGrf.sDateType = "F"
                            slSpotPriceType = "Feed"
                        
                        ElseIf slStr = Trim$("MG") Then    '10-20-10 MG rate line
                            tmGrf.sDateType = "M"
                            slSpotPriceType = "MG"
                        End If
    
                        tmGrf.iPerGenl(6) = -1           'set for valid selection and filtering of audio types
                        '6-24-20 trim live copy test when testing for the blank
                        If (Trim$(tmPLSdf(llIndex).sLiveCopy) = "" Or tmPLSdf(llIndex).sLiveCopy = "R") And (ilAudioIndex = 0 Or ilAudioIndex = 5) Then  'recorded coml and requested all audio types or just the recorded comls
                            tmGrf.iPerGenl(6) = 0            'recorded
                        ElseIf (tmPLSdf(llIndex).sLiveCopy = "L") And (ilAudioIndex = 0 Or ilAudioIndex = 1) Then   'live coml and requested all audio types or just live coml
                            tmGrf.iPerGenl(6) = 1               'live coml
                        ElseIf (tmPLSdf(llIndex).sLiveCopy = "M") And (ilAudioIndex = 0 Or ilAudioIndex = 2) Then    'live Promo and requested all audio types or just live promo
                            tmGrf.iPerGenl(6) = 2               'Live Promo
                        ElseIf (tmPLSdf(llIndex).sLiveCopy = "S") And (ilAudioIndex = 0 Or ilAudioIndex = 6) Then    'recd Promo and requested all audio types or just recd promo
                            tmGrf.iPerGenl(6) = 3               'Recorded Promo
                        ElseIf (tmPLSdf(llIndex).sLiveCopy = "P") And (ilAudioIndex = 0 Or ilAudioIndex = 3) Then    'pre-recd coml and requested all audio types or just pre-recd coml
                            tmGrf.iPerGenl(6) = 4               'Pre-recorded comml
                        ElseIf (tmPLSdf(llIndex).sLiveCopy = "Q") And (ilAudioIndex = 0 Or ilAudioIndex = 4) Then    'pre-recd promo and requested all audio types or just pre-recd promo
                            tmGrf.iPerGenl(6) = 5               'pre-recorded promo
                        End If
    
                        If tmGrf.iPerGenl(6) >= 0 Then  'bypass if valid code not set
                            'Determine if Virtual vehicle, place * next to spot
                            tmGrf.iPerGenl(1) = 0
                            If tmSdf.iVefCode <> tmPLSdf(llIndex).iVefCode Then
                                For ilVsf = LBound(igVirtVefCode) To UBound(igVirtVefCode) - 1 Step 1
                                    If igVirtVefCode(ilVsf) = tmPLSdf(llIndex).iVefCode Then
                                        If tmSdf.sSpotType <> "X" Then
                                            tmGrf.iPerGenl(1) = 1       'flag for Crystal
                                        End If
                                    End If
                                Next ilVsf
                            End If
                            tmGrf.lGenTime = lgNowTime
                            tmGrf.iGenDate(0) = igNowDate(0)
                            tmGrf.iGenDate(1) = igNowDate(1)
                            If ilListIndex = CNT_MGREVENUE Then
                                ilFound = False
                                'check which types of spots to be included (mg/outside/missed/cancel)
                                If tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O" Then
                                    tmGrf.iDateGenl(0, 0) = tmSmf.iMissedDate(0)
                                    tmGrf.iDateGenl(1, 0) = tmSmf.iMissedDate(1)
                                    tmGrf.iMissedTime(0) = tmSmf.iMissedTime(0)
                                    tmGrf.iMissedTime(1) = tmSmf.iMissedTime(1)
                                    tmGrf.iPerGenl(4) = 0       'assume mg/out sorted before missed/cancelled
                                    gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llDate
                                    If llDate >= llStartDate And llDate <= llEndDate Then
                                        If tmSdf.sSchStatus = "G" And ilIncludeMG And tmSdf.sSpotType <> "X" Then
                                            'show only those spots across different vehicles?
                                            ilFound = True
                                            If ilSwitch And tmSmf.iOrigSchVef = tmSdf.iVefCode Then         'only show mg&out that are across vehicles
                                                ilFound = False
                                            End If
                                        ElseIf tmSdf.sSchStatus = "O" And ilIncludeOut And tmSdf.sSpotType <> "X" Then
                                            ilFound = True
                                            If ilSwitch And tmSmf.iOrigSchVef = tmSdf.iVefCode Then     'only show mg&out that are across vehicles
                                                ilFound = False
                                            End If
                                        End If
                                    End If
                                    'test missed date of mg (must be within range)
                                ElseIf tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "C" Then
                                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                    If llDate >= llStartDate And llDate <= llEndDate Then
                                        If tmSdf.sSchStatus = "M" And ilIncludeMissed Then
                                            'tmGrf.iPerGenl(5) = 1   'sort missed/cancelled after mg/out
                                            tmGrf.iPerGenl(4) = 1   'sort missed/cancelled after mg/out
                                            ilFound = True
                                        ElseIf tmSdf.sSchStatus = "C" And ilIncludeCancel Then
                                            'tmGrf.iPerGenl(5) = 1   'sort missed/cancelled after mg/out
                                            tmGrf.iPerGenl(4) = 1   'sort missed/cancelled after mg/out
                                            ilFound = True
                                        End If
                                    End If
                                    'missed date must be within range
                                End If
                                '3-27-11 new option to include billed,unbilled, or both
                                If (ilBillStatus = 1 And tmSdf.sBill <> "Y") Or (ilBillStatus = 2 And tmSdf.sBill <> "N") Then
                                    ilFound = False
                                End If
                                If ilFound Then
                                    '10-20-10 found a mg or outside spot,
                                    'if MG OK to include mg, let it pass; if excluding mg rates, test to see if its a mg rate line
                                    If (ilSaveInclMGForRev = True) Or ((Not ilSaveInclMGForRev) And (Trim$(tmPLSdf(llIndex).sCostType) <> "MG")) Then     'mg spot rates to be included
                                        'Original vehicle
                                        'tmGrf.iPerGenl(4) = tmSmf.iOrigSchVef          'if cancelled or missed, Origin sch vef has been set with ordered vehicle
                                        ''3-27-11 show billed flag on mg rev report
                                        'tmGrf.iPerGenl(6) = 0                          'assume unbilled
                                        'If tmSdf.sBill = "Y" Then
                                        '    tmGrf.iPerGenl(6) = 1                       'tell crystal its been invoiced
                                        'End If
                                        tmGrf.iPerGenl(3) = tmSmf.iOrigSchVef          'if cancelled or missed, Origin sch vef has been set with ordered vehicle
                                        '3-27-11 show billed flag on mg rev report
                                        tmGrf.iPerGenl(5) = 0                          'assume unbilled
                                        If tmSdf.sBill = "Y" Then
                                            tmGrf.iPerGenl(5) = 1                       'tell crystal its been invoiced
                                        End If
                                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                    End If
                                End If
                            Else
                                '2-18-15 spots by advt / Spots Combo:  save the date as a number for sorting due to crystal needing to sort a combination of date and text.
                                'i.e. Date & vehicle name needs to convert date to text in the formula. ie:  12/24/12 & 1/3/13 doesnt sort properly
                                gUnpackDateLong tmGrf.iStartDate(0), tmGrf.iStartDate(1), tmGrf.lLong
                                
                                '-------------------------------
                                'TTP 10674 - Spot and Digital Line combo Export or report?
                                If RptSelCb.rbcOutput(3) Then
                                    'Don't write to GRF, we are Exporing to CSV
                                    '--------------------
                                    '1. ContractNo,
                                    tlExp.lChfCode = tmGrf.lChfCode 'This will be used to Lookup Contract
                                    tlExp.lContract_Number = 0 'will Lookup Contract
                                    '2. ExtContractNo,
                                    tlExp.lExternal_Version_Number = 0 'will Lookup Contract
                                    '3. LineType
                                    If tlExp.sLine_Type = "Digital" Then  'will lookup Line
                                        tlExp.sLine_Type = ""
                                    End If
                                    '4. ContractType
                                    'tlExp.sContract_Type = "" 'will Lookup Contract
                                    '5. Agency,
                                    'tlExp.iAgency_ID = 0 'will Lookup Contract
                                    '6. Advertiser
                                    'tlExp.iAdvertiser_ID = tmSdf.iAdfCode 'will Lookup Contract
                                    '7. Product
                                    'tlExp.sProduct_Name = "" 'will Lookup Contract
                                    '8. Line
                                    tlExp.iPkLineNo = tmSdf.iLineNo
                                    '9. Vehicle
                                    tlExp.iVehicle_ID = tmSdf.iVefCode
                                    'tlExp.sVehicle_Name = "" 'Will Lookup Vehicle
                                    '10. Day
                                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                                    tlExp.sDay = Mid(WeekdayName(Weekday(Format(llDate, "ddddd"))), 1, 3)
                                    '11. SpotAirDate
                                    tlExp.sSpotAirDate = Format(llDate, "ddddd")
                                    '12. SpotAirTime
                                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                                    tlExp.sSpotAirTime = gFormatTimeLong(llTime, "A", "1")
                                    '13. SpotAudioType
                                    'tlExp.sSpotAudioType = "" 'Will lookup Line
                                    '14. ISCIcode
                                    tlExp.sISCIcode = slISCI
                                    '15. Len
                                    tlExp.iSpot_Length = tmSdf.iLen
                                    '16. DigitalLineStartDate (N/A)
                                    tlExp.sLine_Start_Date = ""
                                    '17. DigitalLineEndDate (N/A)
                                    tlExp.sLine_End_Date = ""
                                    '18. Net (Will be computed in mWriteExportFile)
                                    '19. Gross
                                    tlExp.dTotal_Gross = tmGrf.lDollars(0)
                                    '20. Spot status: the status of the spot (scheduled, missed, hidden, cancelled)
                                    If tmSdf.sSpotType = "X" And tmSdf.sSchStatus = "O" Then 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 8)
                                        tlExp.sStatus = ""
                                    Else
                                        Select Case tmSdf.sSchStatus
                                            Case "S": tlExp.sStatus = "Scheduled"
                                            Case "M": tlExp.sStatus = "Missed"
                                            Case "R": tlExp.sStatus = "Ready"
                                            Case "U": tlExp.sStatus = "UnSched"
                                            Case "G": tlExp.sStatus = "Makegood"
                                            Case "A": tlExp.sStatus = "On Alt"
                                            Case "B": tlExp.sStatus = "On Alt & MG"
                                            Case "C": tlExp.sStatus = "Cancelled"
                                            Case "O": tlExp.sStatus = "Outside"
                                            Case "H": tlExp.sStatus = "Hidden"
                                            Case Else: tlExp.sStatus = tmSdf.sSchStatus
                                        End Select
                                    End If
                                    '21. Spot price type: the spot price type from the contract line (charge, N/C, MG, ADU, Recap, Spinoff, Bonus, Package.)
                                    'L=Use line Price; P=Post Log N/C old values => T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff;
                                    'R=Recapturable; A=Audience Deficiency Unit (adu),
                                    'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 6)
                                    tlExp.sPrice_Type = slSpotPriceType
                                    
                                    '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
                                    'tlExp.sDaypart = "" 'Will Lookup Line & Daypart
                                    '23. Line comment: Will Lookup CLF
                                    '24. Digital Line calculation note (N/A)
                                    tlExp.sFormulaComment = ""
                                    
                                    mWriteExportFile ilListIndex, hmExport, hmCHF, hmAgf, hmAdf, hmClf, hmRdf, tlExp
                                Else
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            End If
                        End If                  'tmgrf.ipergenl(7) >= 0
                    End If
                Next llIndex                        'Next spot record gthered
            End If         '6-30-00    gather and write btr records for one vehicle at a time to avoid > 32000 spots per vehicle
        Next ilLoopOnKey        '10-23-14
    End If
    
    '----------------------------------------------------------
    'End If
    If ilListIndex = CNT_MGREVENUE Then         'restore the inclusion/excl of mg spot rate input
        If ilSaveInclMGForRev Then
            RptSelCb!ckcSelC5(7).Value = vbChecked
        Else
            RptSelCb!ckcSelC5(7).Value = vbUnchecked
        End If
    End If
    
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    If RptSelCb.rbcOutput(3) Then
        Close #hmExport
    End If

    Screen.MousePointer = vbDefault
    Erase tmSelChf
    Erase tmSelAgf
    Erase tmSelSlf
    Erase tmPLSdf
    Erase tmSelAdf
    Erase tmSelVef
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAgf)
    btrDestroy hmGrf
    btrDestroy hmSmf
    btrDestroy hmVsf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmAdf
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmFsf
    btrDestroy hmAgf
    Exit Sub
End Sub

'**********************************************************************
'*
'*      Procedure Name:gSpotDateRpt
'*      SPOTS BY DATE AND TIME OR MISSED SPOT REPORTS
'*             Created:4/21/94       By:D. LeVine
'*             Modified:             By:D. Smith
'*
'*            Comments: Generate Spot by Date report
'*
'*      5/28/98 DH :  if Extra or Fill spot, leave
'*                    Status column blank (previously
'*                    was flagged as Outside)
'*
'*      4/20/00 DS :  Converted Spots by Date and Time from
'*                    Bridge to Crystal. Added Start and End
'*                    Time to Missed Spots and Spots by Day and Time
'       8-24-01 dh : option to show full price or cash portion if trade
'       7-21-04 Option to include/exclude local/network spots
'               for both Missed spot report & Spots by Date & Time
'       10-23-04 handle 32000 spots (common routine mobtainsdf) changed to
'               handle 32000+ spots
'       3-30-05 use key2 instead of key0 for speed
'       2-16-06 Add option by slsp for missed spot
'               add gross/net option to missed report
'               add summary/detail option
'       4-12-06 Show game info on spots by Date & Time, not on Missed
'       5-24-06 show x-mid spots on true date it aired, but combine cross-midnight
'               in the same game
'       5-25-06 add gross/net option to spots by date & time
'       6-24-06 default all days to be selected (common mobtainsdf code)
'**********************************************************************
Sub gSpotDateRpt(ilSpotSelType As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilUpper                       ilIndex                                                 *
'******************************************************************************************
'   ilSpotSelType(I)- 2 = missed, 3 = spots by date & time
    Dim slName As String
    Dim llRecNo As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llUpper As Long
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim ilMissedType As Integer
    Dim ilLbcIndex As Integer
    Dim ilVsf As Integer
    Dim ilCostType As Integer
    Dim ilByOrderOrAir As Integer   '0=Order; 1=Aired
    Dim tlVef As VEF
    Dim llContrCode As Long             'selective contr code
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer

    Dim llStartTime As Long             'start time filter entered
    Dim llEndTime As Long               'end time filter entered
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slPctTrade As String            '8-24-01
    Dim slCash As String                '8-24-01
    Dim slDollar As String              '8-24-01
    Dim ilShowFullPrice  As Integer    '8-24-01 true if show full price, else false for cash portion of cash/trade split
    Dim ilLocal As Integer              'true to include contracts spots
    Dim ilFeed As Integer               'true to include network feed spots
    Dim ilPctTrade As Integer           '% trade for contract spot, 0 if network feed spot
    Dim ilIncludeCodes As Integer
    ReDim iluseslfcodes(0 To 0) As Integer
    Dim ilFoundSpot As Integer          'for option by slsp, include /exclude
    Dim ilSlspOption As Integer         'true orfalse
    Dim ilNet As Integer                'true if net
    Dim ilCommPct As Integer            '% of agency comm
    Dim slSharePct As String
    Dim slAmount As String
    Dim ilPropPrice As Integer
    Dim llSpotSeq As Long
    Dim ilListIndex As Integer
    Dim ilVefInx As Integer
    Dim tlAiringSDF() As SPOTTYPESORT
    Dim tlPlSDF As SPOTTYPESORT
    Dim llLLDate As Long
    Dim slLLDate As String
    Dim ilOk As Integer
    Dim slDate As String
    Dim ilSpotSeq As Integer
    Dim llTime As Long
    Dim ilAirVefInx As Integer
    Dim ilAudioIndex As Integer
    Dim ilIncludeType As Integer        '12-29-17
    Dim tlCntTypes As CNTTYPES
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String
    
    'SPOTS BY DATE AND TIME OR MISSED SPOT REPORTS
    ilListIndex = RptSelCb!lbcRptType.ListIndex     'report option
    ilPropPrice = False
    ilIncludeType = False               '12-29-17 assume no contract type testing
    'if entering as Guide, user allowed to see question to show Actual spot price or proposal spot price on Spots by DAte and time report
    If RptSelCb!rbcSelC11(1).Value Then
        ilPropPrice = True
    End If

    slStartTime = RptSelCb!edcSet1.Text        'start time
    llStartTime = gTimeToLong(slStartTime, False)
    slEndTime = RptSelCb!edcSet2.Text            'end time
    llEndTime = gTimeToLong(slEndTime, True)
    ilShowFullPrice = True            '8-24-01
    If RptSelCb!rbcSelC4(1).Value Then  'show cash portion only of cash/trade split
        ilShowFullPrice = False
    End If
    

    'default all days selected, common code (mobtainsdf) used
    For illoop = 0 To 6
        RptSelCb!ckcSelC8(illoop) = vbChecked
    Next illoop

    Screen.MousePointer = vbHourglass
    If ilSpotSelType = 2 Then                   'ilSpotSelType(I)- 2 = missed, 3 = spots by date & time
        If Not gSetFormula("ReportType", "'M'") Then 'Missed Spots
            MsgBox "RptGenCb - error in DateRange", vbOKOnly, "Report Error"
            Exit Sub
        End If
        If RptSelCb!rbcSelC7(0).Value = True Then          'vehicle option
            ilSlspOption = False
        Else
            ilSlspOption = True
        End If
        
        
        '4-7-20 change to use same control as Spots by date and time for summary only, then used combined code below
'        If RptSelCb!ckcSelC10(0).Value = vbChecked Then       'Summary only
'        If RptSelCb!ckcSelC15.Value = vbChecked Then       '4-7-20 change control used for Summary only
'            If Not gSetFormula("TotalsBy", "'S'") Then      'summary
'                MsgBox "RptGenCb - error in Formula TotalsBy", vbOKOnly, "gSpotDateRpt"
'                Exit Sub
'            End If
'        Else                                            'detail
'            If Not gSetFormula("TotalsBy", "'D'") Then
'                MsgBox "RptGenCb - error in Formula TotalsBy", vbOKOnly, "gSpotDateRpt"
'                Exit Sub
'            End If
'        End If
        ilAudioIndex = 0            'force to filter all audio types for missed
    End If                  '4-7-20
'    ElseIf ilSpotSelType = 1 Or ilSpotSelType = 3 Then 'ilSpotSelType(I)- 2 = missed, 3 = spots by date & time
        '4-7-20 Missed spots now come thru here to select contract types
        ilIncludeType = True                                '12-29-17 do contract type testing
        tlCntTypes.iHold = gSetCheck(RptSelCb!ckcSelC10(0).Value)
        tlCntTypes.iOrder = gSetCheck(RptSelCb!ckcSelC10(1).Value)
        tlCntTypes.iStandard = gSetCheck(RptSelCb!ckcSelC6(0).Value)
        tlCntTypes.iReserv = gSetCheck(RptSelCb!ckcSelC6(1).Value)
        tlCntTypes.iRemnant = gSetCheck(RptSelCb!ckcSelC6(2).Value)
        tlCntTypes.iDR = gSetCheck(RptSelCb!ckcSelC6(3).Value)
        tlCntTypes.iPI = gSetCheck(RptSelCb!ckcSelC6(4).Value)
        tlCntTypes.iPSA = gSetCheck(RptSelCb!ckcSelC6(5).Value)
        tlCntTypes.iPromo = gSetCheck(RptSelCb!ckcSelC6(6).Value)
    If ilSpotSelType = 3 Then                   '4-7-20 Spots by Date & time
        ilAudioIndex = RptSelCb!cbcSet1.ListIndex           'audio type selection, 0 = all audio types
        If Not gSetFormula("ReportType", "'D'") Then
            MsgBox "RptGenCb - error in DateRange", vbOKOnly, "gSpotDateRpt"
            Exit Sub
        End If
        ilSlspOption = False        'this option applies to Missed report only
    End If
        '4-7-20 Both Missed and Spots by DAte and Time have summary only option
        If RptSelCb!ckcSelC15.Value = vbChecked Then       'Summary only
            If Not gSetFormula("TotalsBy", "'S'") Then      'summary
                MsgBox "RptGenCb - error in Formula TotalsBy", vbOKOnly, "gSpotDateRpt"
                Exit Sub
            End If
        Else                                            'detail
            If Not gSetFormula("TotalsBy", "'D'") Then
                MsgBox "RptGenCb - error in Formula TotalsBy", vbOKOnly, "gSpotDateRpt"
                Exit Sub
            End If
        End If

 '   End If
    If RptSelCb!rbcSelC9(0).Value = True Then       'gross
        If Not gSetFormula("GrossNet", "'G'") Then
            MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotDateRpt"
            Exit Sub
        End If
        ilNet = False
    Else                                            'net
        If Not gSetFormula("GrossNet", "'N'") Then      'net
            MsgBox "RptGenCb - error in Formula GrossNet", vbOKOnly, "gSpotDateRpt"
            Exit Sub
        End If
        ilNet = True
    End If

'    slStartDate = RptSelCb!edcSelCFrom.Text   'Start date
    slStartDate = RptSelCb!CSI_CalFrom.Text   'Start date   9-12-19 use csi calendar control

'    slEndDate = RptSelCb!edcSelCTo.Text   'End date
    slEndDate = RptSelCb!CSI_CalTo.Text   'End date
    slDateRange = "From " & slStartDate & " To " & slEndDate
    If slEndDate = "TFN" Then
        slEndDate = ""
    End If
    If (slStartDate = "") And (slEndDate = "") Then
        slDateRange = "All Dates"
    End If
    
    If slStartDate = "" Then
        slStartDate = "1/1/1970"
        slDateRange = "Thru " & slEndDate
    End If
    If slEndDate = "" Then
        slEndDate = "12/31/2069"
        slDateRange = "From " & slStartDate
    End If
    
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    
    If Not gSetFormula("DateRange", "'" & slDateRange & "'") Then
        MsgBox "RptGenCb - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    
    ilMissedType = 0
    If ilListIndex = CNT_MISSED Then            '12-28-17
        If RptSelCb!ckcSelC3(0).Value = vbChecked Then
            ilMissedType = 1
        End If
        If RptSelCb!ckcSelC3(1).Value = vbChecked Then
            ilMissedType = ilMissedType + 2
        End If
        If RptSelCb!ckcSelC3(2).Value = vbChecked Then
            ilMissedType = ilMissedType + 4
        End If
    End If
    
    mSetCostType ilCostType             'set bit pattern of types of spots to include
    gObtainVirtVehList
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
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        btrDestroy hmCHF
        btrDestroy hmClf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)

    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmFsf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmAgf
        btrDestroy hmFsf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSsf
        btrDestroy hmAgf
        btrDestroy hmFsf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)
    If RptSelCb!rbcSelC14(1).Value = True Then          'use airing vehicles
        hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmVLF)
            ilRet = btrClose(hmSsf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmVsf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmGrf)
            btrDestroy hmVLF
            btrDestroy hmSsf
            btrDestroy hmAgf
            btrDestroy hmFsf
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmAdf
            btrDestroy hmVef
            btrDestroy hmSdf
            btrDestroy hmVsf
            btrDestroy hmSmf
            btrDestroy hmCff
            btrDestroy hmGrf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        imVlfRecLen = Len(tmVlf)
        
        'need 2 handles for the SSF, one for selling and one for airing
        hmAirSSF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAirSSF, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAirSSF)
            ilRet = btrClose(hmVLF)
            ilRet = btrClose(hmSsf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmFsf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmSdf)
            ilRet = btrClose(hmVsf)
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmGrf)
            btrDestroy hmAirSSF
            btrDestroy hmVLF
            btrDestroy hmSsf
            btrDestroy hmAgf
            btrDestroy hmFsf
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmAdf
            btrDestroy hmVef
            btrDestroy hmSdf
            btrDestroy hmVsf
            btrDestroy hmSmf
            btrDestroy hmCff
            btrDestroy hmGrf
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If ilSlspOption = True Then           'slsp option , open file
        hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmSlf)
            btrDestroy hmSlf
        End If
        imSlfRecLen = Len(tmSlf)

        gObtainCodesForMultipleLists 2, tgSalesperson(), ilIncludeCodes, iluseslfcodes(), RptSelCb
    End If


    ilLocal = gSetCheck(RptSelCb!ckcSelC12(0).Value)       'include local spots (vs network feed)
    ilFeed = gSetCheck(RptSelCb!ckcSelC12(1).Value)       'include network (feed) vs local

    llContrCode = 0                     'assume all contracts to be output (else contr code #)
    slStr = RptSelCb!edcSet3.Text  'see if selective contract entred
    If slStr <> "" Then
        llRecNo = Val(slStr)
        tmChfSrchKey1.lCntrNo = llRecNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (llRecNo = tmChf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")            'set the selective contr code only if no errors
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If llRecNo = tmChf.lCntrNo Then
            llContrCode = tmChf.lCode
        Else
            llContrCode = -1                    'get nothing, invalid contr #
        End If
    End If

    tmAdf.iCode = 0
    tmSlf.iCode = 0
    tmVef.iCode = 0

    If ilSpotSelType <> 2 Then      'spots by date and time option ( 1=Scheduled only; 2=Missed only; 3=Both)
        If RptSelCb!rbcSelC14(0).Value = True Then          'selling vehicles
            ilLbcIndex = 6
        Else
            ilLbcIndex = 8          'airing option
        End If
    Else                    'missed option
        ilLbcIndex = 6
    End If

   
    DoEvents
        llRecNo = 0
'VB6**        slPort = LlVBPrintGetPort(hdJob)
        RptSelCb!lbcLnCode.Clear        'Init list box, sorted by internal code.   build in a
                                        'sequence # to that the order of spots can be maintained
        
        ReDim tmPLSdf(0 To 0) As SPOTTYPESORT
        ReDim tmSeqSortType(0 To 0) As SEQSORTTYPE
        For ilVehicle = 0 To RptSelCb!lbcSelection(ilLbcIndex).ListCount - 1 Step 1
            If RptSelCb!lbcSelection(ilLbcIndex).Selected(ilVehicle) Then
                If ilLbcIndex = 8 Then                  'airing vehicle list
                    slNameCode = tgVehicle(ilVehicle).sKey
                Else                                    'selling vehicle list
                    slNameCode = tgCSVNameCode(ilVehicle).sKey
                End If
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVefInx = gBinarySearchVef(ilVefCode)
                If ilVefInx >= 0 Then
                    If tgMVef(ilVefInx).sType = "A" Then            'airing vehicle, get the associated links to retrieve the spot time
                        gGetAiringSpots hmVLF, hmSsf, hmAirSSF, hmSdf, ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, tlAiringSDF()
                        'if any airing spots, merge them with the tmPlSDF array to sort them
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
                        For illoop = LBound(tlAiringSDF) To UBound(tlAiringSDF) - 1
                            
                            ilSpotSeq = 0
                            'ilSpotSelType(I)- 1=Scheduled only; 2=Missed only; 3=Both
'                            ilOk = mSpotDateFilter(tlAiringSDF(ilLoop), llStartDate, llEndDate, llStartTime, llEndTime, ilSpotSelType, 3, True, ilMissedType, False, ilCostType, ilByOrderOrAir, False, llContrCode, ilPropPrice, llLLDate)
                            ilOk = mSpotDateFilter(tlAiringSDF(illoop), llStartDate, llEndDate, llStartTime, llEndTime, ilSpotSelType, 3, True, ilMissedType, False, ilCostType, ilByOrderOrAir, ilIncludeType, llContrCode, ilPropPrice, llLLDate, tlCntTypes)
                            
                            If ilOk Then
                                tmPLSdf(llUpper) = tlAiringSDF(illoop)
                                'ilSpotSeq = mGetSeqNo(tmPLSdf(llUpper).tSdf)
                                ilAirVefInx = gBinarySearchVef(tmPLSdf(llUpper).iVefCode)       'get the AIRING vehiclename; in the SDF entry is the selling vehicle code
                                'If ilAirVefInx > 0 Then
                                If ilAirVefInx >= 0 Then
                                    tmPLSdf(llUpper).sKey = tgMVef(ilAirVefInx).sName
                                    gUnpackDateForSort tmPLSdf(llUpper).tSdf.iDate(0), tmPLSdf(llUpper).tSdf.iDate(1), slDate
                                    tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slDate
                                    If (tmPLSdf(llUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "O") Then
                                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|A"
                                    Else
                                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|Z"
                                    End If
                                    gUnpackTimeLong tmPLSdf(llUpper).tSdf.iTime(0), tmPLSdf(llUpper).tSdf.iTime(1), False, llTime
                                    slStr = Trim$(str$(llTime))
                                    Do While Len(slStr) < 6
                                        slStr = "0" & slStr
                                    Loop
                                    If ilSpotSeq < 10 Then
                                        slStr = slStr & "0" & Trim$(str$(ilSpotSeq))
                                    Else
                                        slStr = slStr & Trim$(str$(ilSpotSeq))
                                    End If
                                
                                    tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slStr
                                    tmPLSdf(llUpper).tSdf.iVefCode = tmPLSdf(llUpper).iVefCode          'set the spots selling vehicle code to reference the airing vehicle for the common code to write out to grf for crystal
                                    ReDim Preserve tmPLSdf(LBound(tmPLSdf) To llUpper + 1) As SPOTTYPESORT
                                    llUpper = llUpper + 1
                                End If
                            End If
                            
                        Next illoop
                    Else                'selling, conventional, game
'                        mObtainSdf ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilSpotSelType, 3, True, ilMissedType, False, ilCostType, ilByOrderOrAir, False, llContrCode, ilLocal, ilFeed, ilPropPrice, ilListIndex
                        'ilSpotSelType(I)- 1=Scheduled only; 2=Missed only; 3=Both
                        mObtainSdf ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilSpotSelType, 3, True, ilMissedType, False, ilCostType, ilByOrderOrAir, ilIncludeType, llContrCode, ilLocal, ilFeed, ilPropPrice, ilListIndex, tlCntTypes
                        llUpper = UBound(tmPLSdf)
                    End If
                End If
            End If
        Next ilVehicle


        llUpper = UBound(tmSeqSortType)
        If llUpper > 0 Then
            ArraySortTyp fnAV(tmSeqSortType(), 0), llUpper, 0, LenB(tmSeqSortType(0)), 0, LenB(tmSeqSortType(0).sKey), 0
        End If
        
        For llUpper = LBound(tmPLSdf) To UBound(tmPLSdf) - 1
            tmSdf = tmPLSdf(llUpper).tSdf
            If tmSdf.lChfCode = 0 Then                 'network feed spot
                ilPctTrade = 0
            Else
                tmChfSrchKey.lCode = tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                ilPctTrade = tmChf.iPctTrade
                ilFoundSpot = True      'assume not by slsp and to print spot
                If ilSlspOption = True Then           'slsp option , test for selectivity
                    ilFoundSpot = False
                    If ilIncludeCodes Then
                        For illoop = LBound(iluseslfcodes) To UBound(iluseslfcodes) - 1 Step 1
                            If iluseslfcodes(illoop) = tmChf.iSlfCode(0) Then
                                ilFoundSpot = True
                                Exit For
                            End If
                        Next illoop
                    Else
                        ilFoundSpot = True
                        For illoop = LBound(iluseslfcodes) To UBound(iluseslfcodes) - 1 Step 1
                            If iluseslfcodes(illoop) = tmChf.iSlfCode(0) Then
                                ilFoundSpot = False
                                Exit For
                            End If
                        Next illoop
                    End If
                End If

            End If

            If ilFoundSpot Then                 'OK to include spot?
                If tmAdf.iCode <> tmChf.iAdfCode Then
                    tmAdfSrchKey.iCode = tmSdf.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                'TTP 10674 - Spot and Digital Line combo report: get ISCI
                mObtainCopy slProduct, slZone, slCart, slISCI
                
                tmGrf.lChfCode = tmSdf.lChfCode            'contr code
                'tmGrf.lDollars(2) = tmSdf.lFsfCode      'network feed code
                tmGrf.lDollars(1) = tmSdf.lFsfCode      'network feed code
                tmGrf.iCode2 = tmSdf.iLineNo
                tmGrf.iVefCode = tmSdf.iVefCode
                If tmSdf.sXCrossMidnight = "Y" Then     'cross midnight, show spot on actual date it aired, also so it sorts properly
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    gPackDateLong llDate + 1, ilDate(0), ilDate(1)
                    tmGrf.iStartDate(0) = ilDate(0)
                    tmGrf.iStartDate(1) = ilDate(1)
                Else
                    tmGrf.iStartDate(0) = tmSdf.iDate(0)
                    tmGrf.iStartDate(1) = tmSdf.iDate(1)
                End If
                tmGrf.iTime(0) = tmSdf.iTime(0)
                tmGrf.iTime(1) = tmSdf.iTime(1)
                'tmGrf.iPerGenl(1) = tmSdf.iLen
                'tmGrf.iPerGenl(3) = 0                   'default to not a mg or outside (vehicle for mg or out)
                tmGrf.iPerGenl(0) = tmSdf.iLen
                tmGrf.iPerGenl(2) = 0                   'default to not a mg or outside (vehicle for mg or out)
                tmGrf.iRdfCode = 0                        '2-24-03 init the missed reason code
                tmSmf.iOrigSchVef = tmSdf.iVefCode
                'If ilSpotSelType = 2 Then               '4-12-06 missed option, no game info
                '    tmGrf.iPerGenl(4) = 0
                'Else
                    '1-23-07 Missed spots and spots by date and time show spots by game if applicable
                    'tmGrf.iPerGenl(4) = tmSdf.iGameNo       '4-12-06 show game info for vehicle
                    tmGrf.iPerGenl(3) = tmSdf.iGameNo       '4-12-06 show game info for vehicle
                'End If
                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    'tmSmfSrchKey.lFsfCode = tmSdf.lFsfCode
                    'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                    'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                    'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                    'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo) And (tmSmf.lFsfCode = tmSdf.lFsfCode)
                    '    If (tmSmf.lSdfCode = tmSdf.lCode) And (tmSmf.lFsfCode = tmSdf.lFsfCode) Then
                    '        Exit Do
                    '    End If
                    '    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    'Loop

                    '3-30-05 use key2 instead of key0 for speed
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation

                End If
                If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                    tmVefSrchKey.iCode = tmSmf.iOrigSchVef
                    ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If

                If InStr(tmPLSdf(llUpper).sCostType, "Fill") > 0 Then '+/- fill spot, dont give it a sch status
                    slStr = ""
                ElseIf tmSdf.sSchStatus = "S" Then
                    slStr = "Scheduled"
                    'slStr = ""
                ElseIf tmSdf.sSchStatus = "M" Then
                    slStr = "Missed"
                    tmGrf.iRdfCode = tmSdf.iMnfMissed         '2-24-03
                ElseIf tmSdf.sSchStatus = "R" Then
                    slStr = "Ready"
                ElseIf tmSdf.sSchStatus = "U" Then
                    slStr = "UnSched"
                ElseIf tmSdf.sSchStatus = "G" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        'slStr = "Makegood" & Chr$(10) & Trim$(tlVef.sname)
                        slStr = "Makegood"
                        'tmGrf.iPerGenl(3) = tlVef.iCode
                        tmGrf.iPerGenl(2) = tlVef.iCode
                    Else
                        slStr = "Makegood"
                    End If
                ElseIf tmSdf.sSchStatus = "O" Then
                    If tmSmf.iOrigSchVef <> tmSdf.iVefCode Then
                        If tmSdf.sSpotType = "X" Then
                            slStr = ""
                        Else
                            'slStr = "Outside" & Chr$(10) & Trim$(tlVef.sname)
                            slStr = "Outside"
                            'tmGrf.iPerGenl(3) = tlVef.iCode
                            tmGrf.iPerGenl(2) = tlVef.iCode
                        End If
                    Else
                        If tmSdf.sSpotType = "X" Then
                            slStr = ""
                        Else
                            slStr = "Outside"
                        End If
                    End If
                ElseIf tmSdf.sSchStatus = "C" Then
                    slStr = "Cancelled"
                    tmGrf.iRdfCode = tmSdf.iMnfMissed         '2-24-03
                ElseIf tmSdf.sSchStatus = "H" Then
                    slStr = "Hidden"
                    tmGrf.iRdfCode = tmSdf.iMnfMissed         '2-24-03
                ElseIf tmSdf.sSchStatus = "A" Then
                    slStr = "On Alt"
                ElseIf tmSdf.sSchStatus = "B" Then
                    slStr = "On Alt & MG"
                End If
                tmGrf.sGenDesc = Trim$(slStr)

                tmGrf.sBktType = tmSdf.sSpotType
                tmGrf.sDateType = tmSdf.sPriceType
                slStr = Trim$(tmPLSdf(llUpper).sCostType)
                'tmGrf.lDollars(1) = 0
                tmGrf.lDollars(0) = 0
                'slPctTrade = gIntToStrDec(ilPctTrade, 0)
                If InStr(slStr, ".") <> 0 Then          'its an amount, not N/c or any other text
                    'tmGrf.lDollars(1) = gStrDecToLong(slStr, 2) 'convert string decimal to long value
                    tmGrf.lDollars(0) = gStrDecToLong(slStr, 2) 'convert string decimal to long value
                    'determine gross or net
                    If tmChf.iAgfCode = 0 Then          'direct
                        ilCommPct = 10000                'no commission
                    Else
                        ilCommPct = 8500         'default to commissionable if no agency found
                        'see what the agency comm is defined as
                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, Len(tmAgf), tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                        If ilRet = BTRV_ERR_NONE Then
                            ilCommPct = (10000 - tmAgf.iComm)
                        End If          'ilret = btrv_err_none
                    End If

                    If ilNet Then       'first get the net value if applicable
                        'slAmount = gLongToStrDec(tmGrf.lDollars(1), 2)
                        slAmount = gLongToStrDec(tmGrf.lDollars(0), 2)
                        slSharePct = gIntToStrDec(ilCommPct, 4)
                        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                        slStr = gRoundStr(slStr, ".01", 2)
                        'tmGrf.lDollars(1) = gStrDecToLong(slStr, 2) 'adjusted net
                        tmGrf.lDollars(0) = gStrDecToLong(slStr, 2) 'adjusted net
                    End If

                    '8-24-01 if trade, see if option to get cash portion only
                    If Not ilShowFullPrice And ilPctTrade > 0 Then
                        slPctTrade = gIntToStrDec(ilPctTrade, 0)
                        slCash = gSubStr("100.", slPctTrade)         'get the cash portion this order
                        slDollar = gDivStr(gMulStr(slStr, slCash), "100")
                        'tmGrf.lDollars(1) = gStrDecToLong(slDollar, 2)
                        tmGrf.lDollars(0) = gStrDecToLong(slDollar, 2)
                    End If
                ElseIf slStr = Trim$("ADU") Then
                    tmGrf.sDateType = "A"                'flag ADU spot
                ElseIf slStr = Trim$("Bonus") Then
                    tmGrf.sDateType = "B"
                ElseIf slStr = Trim$("Feed") Then    'network feed spot
                    tmGrf.sDateType = "F"
                ElseIf slStr = Trim$("MG") Then     '10-20-10 mg rate line
                    tmGrf.sDateType = "M"
                End If

'                '5-31-12 get the live copy flag, convert to digit as no string fields remain to place code into
'                If tmPLSdf(llUpper).sLiveCopy = "" Or tmPLSdf(llUpper).sLiveCopy = "R" Then
'                    tmGrf.iPerGenl(7) = 0           'recorded
'                ElseIf tmPLSdf(llUpper).sLiveCopy = "L" Then
'                    tmGrf.iPerGenl(7) = 1               'live coml
'                ElseIf tmPLSdf(llUpper).sLiveCopy = "M" Then
'                    tmGrf.iPerGenl(7) = 2               'Live Promo
'                ElseIf tmPLSdf(llUpper).sLiveCopy = "S" Then
'                    tmGrf.iPerGenl(7) = 3               'Recorded Promo
'                ElseIf tmPLSdf(llUpper).sLiveCopy = "P" Then
'                    tmGrf.iPerGenl(7) = 4               'Pre-recorded comml
'                ElseIf tmPLSdf(llUpper).sLiveCopy = "Q" Then
'                    tmGrf.iPerGenl(7) = 5               'pre-recorded promo
'                End If
'
                '10-26-15 audio type selection
'                tmGrf.iPerGenl(7) = -1
'                If (tmPLSdf(llUpper).sLiveCopy = "" Or tmPLSdf(llUpper).sLiveCopy = "R") And (ilAudioIndex = 0 Or ilAudioIndex = 5) Then  'recorded coml and requested all audio types or just the recorded comls
'                    tmGrf.iPerGenl(7) = 0           'recorded
'                ElseIf (tmPLSdf(llUpper).sLiveCopy = "L") And (ilAudioIndex = 0 Or ilAudioIndex = 1) Then   'live coml and requested all audio types or just live coml
'                    tmGrf.iPerGenl(7) = 1               'live coml
'                ElseIf (tmPLSdf(llUpper).sLiveCopy = "M") And (ilAudioIndex = 0 Or ilAudioIndex = 2) Then    'live Promo and requested all audio types or just live promo
'                    tmGrf.iPerGenl(7) = 2               'Live Promo
'                ElseIf (tmPLSdf(llUpper).sLiveCopy = "S") And (ilAudioIndex = 0 Or ilAudioIndex = 6) Then    'recd Promo and requested all audio types or just recd promo
'                    tmGrf.iPerGenl(7) = 3               'Recorded Promo
'                ElseIf (tmPLSdf(llUpper).sLiveCopy = "P") And (ilAudioIndex = 0 Or ilAudioIndex = 3) Then    'pre-recd coml and requested all audio types or just pre-recd coml
'                    tmGrf.iPerGenl(7) = 4               'Pre-recorded comml
'                ElseIf (tmPLSdf(llUpper).sLiveCopy = "Q") And (ilAudioIndex = 0 Or ilAudioIndex = 4) Then    'pre-recd promo and requested all audio types or just pre-recd promo
'                    tmGrf.iPerGenl(7) = 5               'pre-recorded promo
'                End If
                tmGrf.iPerGenl(6) = -1
                '6-24-20 trim live copy test when testing for the blank
                If (Trim$(tmPLSdf(llUpper).sLiveCopy) = "" Or tmPLSdf(llUpper).sLiveCopy = "R") And (ilAudioIndex = 0 Or ilAudioIndex = 5) Then  'recorded coml and requested all audio types or just the recorded comls
                    tmGrf.iPerGenl(6) = 0           'recorded
'Case "R": slString = "RC"
                ElseIf (tmPLSdf(llUpper).sLiveCopy = "L") And (ilAudioIndex = 0 Or ilAudioIndex = 1) Then   'live coml and requested all audio types or just live coml
                    tmGrf.iPerGenl(6) = 1               'live coml
'Case "L": slString = "LC"
                ElseIf (tmPLSdf(llUpper).sLiveCopy = "M") And (ilAudioIndex = 0 Or ilAudioIndex = 2) Then    'live Promo and requested all audio types or just live promo
                    tmGrf.iPerGenl(6) = 2               'Live Promo
'Case "M": slString = "LP"
                ElseIf (tmPLSdf(llUpper).sLiveCopy = "S") And (ilAudioIndex = 0 Or ilAudioIndex = 6) Then    'recd Promo and requested all audio types or just recd promo
                    tmGrf.iPerGenl(6) = 3               'Recorded Promo
'Case "S": slString = "RP"
                ElseIf (tmPLSdf(llUpper).sLiveCopy = "P") And (ilAudioIndex = 0 Or ilAudioIndex = 3) Then    'pre-recd coml and requested all audio types or just pre-recd coml
                    tmGrf.iPerGenl(6) = 4               'Pre-recorded comml
'Case "P": slString = "PC"
                ElseIf (tmPLSdf(llUpper).sLiveCopy = "Q") And (ilAudioIndex = 0 Or ilAudioIndex = 4) Then    'pre-recd promo and requested all audio types or just pre-recd promo
                    tmGrf.iPerGenl(6) = 5               'pre-recorded promo
'Case "Q": slString = "PP"
                End If

                'If tmGrf.iPerGenl(7) >= 0 Then  'bypass if valid code not set
                If tmGrf.iPerGenl(6) >= 0 Then  'bypass if valid code not set

                    'Determine if Virtual vehicle, place * next to spot
                    'tmGrf.iPerGenl(2) = 0
                    tmGrf.iPerGenl(1) = 0
                    If tmSdf.iVefCode <> tmPLSdf(llUpper).iVefCode Then
                        For ilVsf = LBound(igVirtVefCode) To UBound(igVirtVefCode) - 1 Step 1
                            'If igVirtVefCode(ilVsf) = tmClf.iVefCode Then
                            If igVirtVefCode(ilVsf) = tmPLSdf(llUpper).iVefCode Then
                                If tmSdf.sSpotType <> "X" Then
                                    'tmGrf.iPerGenl(2) = 1       'flag for Crystal
                                    tmGrf.iPerGenl(1) = 1       'flag for Crystal
                                End If
                            End If
                        Next ilVsf
                    End If
                    
    '                'this method seems to be slower than binary search
    '                slStr = Trim$(str(tmPLSdf(llUpper).tSdf.lCode))
    '                Do While Len(slStr) < 10
    '                    slStr = "0" & slStr
    '                Loop
    '                tmGrf.lLong = 0
    '                If RptSelCb!lbcLnCode.ListCount > 0 Then
    '                    gFindMatch slStr, 0, RptSelCb!lbcLnCode
    '                    ilRet = gLastFound(RptSelCb!lbcLnCode)
    '                    If ilRet >= 0 Then
    '                        tmGrf.lLong = RptSelCb!lbcLnCode.ItemData(ilRet)
    '                    End If
    '                End If
                    
                    tmGrf.lLong = 0
                    'llSpotSeq = mBinarySearchSDF(tmSdf.lCode, tmSpotSeq())
                    llSpotSeq = mBinarySearchSDF(tmSdf.lCode, tmSeqSortType())
                    If llSpotSeq >= 0 Then
                        'tmGrf.lLong = tmSpotSeq(llSpotSeq).iSeq
                        tmGrf.lLong = tmSeqSortType(llSpotSeq).iSeqNo
                    End If
    
                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                    tmGrf.lGenTime = lgNowTime
                    tmGrf.iGenDate(0) = igNowDate(0)
                    tmGrf.iGenDate(1) = igNowDate(1)
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If
            End If                      'if ilFoundSpot
        Next llUpper                        'Next spot record gthered
    Screen.MousePointer = vbDefault
    Erase tmSelChf
    Erase tmSelSlf
    Erase tmPLSdf
    Erase tlAiringSDF
    Erase iluseslfcodes
    Erase tmSpotSeq, tmSeqSortType
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmAgf)
        
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        btrDestroy hmFsf
        btrDestroy hmAgf
        If ilSlspOption = True Then
            ilRet = btrClose(hmSlf)
            btrDestroy hmSlf
        End If
        
        If RptSelCb!rbcSelC14(1).Value = True Then      'use airing vehicles
            ilRet = btrClose(hmVLF)
            ilRet = btrClose(hmAirSSF)
            btrDestroy hmVLF
            btrDestroy hmAirSSF
        End If
    Exit Sub
End Sub

'**********************************************************************
'*
'*      Procedure Name:gSpotSalesAdvtRpt
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Generate Spot Sales report by
'*                      Advertiser
'*
'*      4/6/99 Add option in addition to As Ordered &
'*              As Aired, a third option:  As Aired &
'*              pkg by ordered.
'*
'*       9/11/00 D.S. Convert to Crystal from Bridge
'*       2/01/01 D.S. Corrected Adv Not Showing
'*       6/3/01 dh Out of string space due to slField
'                  not intialized
'*      10-6-03 Prevent subscript out of range for 32000+ records for
'*              Sales Source option
'       10-23-04 handle 32000+ spots
'       11-30-04 change to access smf by key2 instead of key0 for speed
'       6-25-06 add day selectivity
'***********************************************************************
Sub gSpotSalesAdvtRpt()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       ilUpper                                                 *
'******************************************************************************************
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim llDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llIndex As Long             '10-23-04
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim slGross As String
    Dim slNet As String
    Dim slAgyComm As String
    Dim slCommission As String
    Dim llCSofNoSpots As Long
    Dim slCSofGross As String
    Dim slCSofCommission As String
    Dim slCSofNet As String
    Dim llCVehNoSpots As Long
    Dim slCVehGross As String
    Dim slCVehCommission As String
    Dim slCVehNet As String
    Dim llCAllNoSpots As Long
    Dim slCAllGross As String
    Dim slCAllCommission As String
    Dim slCAllNet As String
    Dim llTSofNoSpots As Long
    Dim slTSofGross As String
    Dim slTSofCommission As String
    Dim slTSofNet As String
    Dim llTVehNoSpots As Long
    Dim slTVehGross As String
    Dim slTVehCommission As String
    Dim slTVehNet As String
    Dim llTAllNoSpots As Long
    Dim slTAllGross As String
    Dim slTAllCommission As String
    Dim slTAllNet As String
    Dim ilFound As Integer
    Dim ilAdvtOnly As Integer
    Dim ilNewPage As Integer
    Dim ilNoLinesPerPage As Integer
    'Dim ilNoRowsPrt As Integer
    Dim llNoRowsPrt As Long             '12-5-17 prevent overflow with # of spots

    Dim ilMaxRowPerPage As Integer
    Dim llCntrNo As Long
    Dim slSOFName As String
    Dim slAdvtName As String
    Dim slPctTrade As String
    Dim ilMissedType As Integer
    Dim slIncludeTitle As String
    Dim slTitleNetC As String
    Dim slTitleNetT As String
    ReDim slField(0 To 11) As String
    Dim ilAnyOutput As Integer
    Dim ilPrevVeh As Integer
    Dim ilPass As Integer       'Pass: 1=Cash; 2=Trade
    Dim ilStartPass As Integer
    Dim ilEndPass As Integer
    Dim tlVef As VEF
    Dim ilVsf As Integer
    Dim ilConvUpper As Integer
    Dim ilBucket As Integer
    Dim ilSBucket As Integer
    Dim ilEBucket As Integer
    Dim llBucketDate As Long
    Dim llUpper As Long             '10-23-04
    'Dim ilVsfCode As Integer
    Dim llVsfCode As Long
    Dim ilByOrderOrAir As Integer   '0=Order; 1=Aired, 2 = as aired, pkg ordered
    Dim ilUorC As Integer           'for each spot, either 1 for spotcount, or # 30" units
                                    'units are rounded up (i.e. 1-30 = 1 unit, 60 = 2 units, etc)
    Dim ilCostType As Integer       'inclusion of different spot cost types (n/c, adu, bonus, fill, etc)
    'ReDim ilProcVefCode(1 To 1) As Integer
    ReDim ilProcVefCode(0 To 0) As Integer
    Dim slStartTime As String                   'start time filter entered
    Dim slEndTime As String                     'end time filter entered
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilTtlTest As Integer
    Dim ilTemp As Integer               '8-15-02
    'ReDim tlSlf(1 To 1) As SLF              '8-15-02
    ReDim tlSlf(0 To 0) As SLF              '8-15-02
    Dim ilMatchSS As Integer                'sales source from primary slsp of contract
    Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
    Dim llContrCode As Long     'selective contr code
    Dim blOwnerOnly As Boolean      '5-9-17
    Dim llStartDate As Long         '5-9-17 start date requested
    Dim llLoop As Long              '12-4-17 prevent overflow in spot array
    Dim llFound As Long             '12-4-17 prevent overflow in spot array
    Dim slWhichPrice As String * 1      '12-5-18 option to use ACtual or proposal price
    Dim slISCI As String
    
    ilPrevVeh = 0               'init first time thru
    Screen.MousePointer = vbHourglass

    slStartTime = RptSelCb!edcSelCTo.Text        'start time
    llStartTime = gTimeToLong(slStartTime, False)
    slEndTime = RptSelCb!edcSelCTo1.Text            'end time
    llEndTime = gTimeToLong(slEndTime, True)

'    slStartDate = RptSelCb!edcSelCFrom.Text   'Start date
    slStartDate = RptSelCb!CSI_CalFrom.Text   'Start date, 9-11-19 use csi calendar control vs edit box

    If slStartDate = "" Then
        slStartDate = "1/5/1970" 'Monday
    End If
'    slEndDate = RptSelCb!edcSelCFrom1.Text   'End date
    slEndDate = RptSelCb!CSI_CalTo.Text   'End date     9-11-19 csi calendar

    If (StrComp(slEndDate, "TFN", 1) = 0) Or (Len(slEndDate) = 0) Then
        slEndDate = "12/29/2069"    'Sunday
    End If
    slDateRange = "From " & slStartDate & " To " & slEndDate & ", " & slStartTime & "-" & slEndTime
    If Not gSetFormula("Show Date Range", "'" & slDateRange & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    'If RptSelCb!ckcSelC3(0).Value Then
        ilAdvtOnly = False
    'Else
    '    ilAdvtOnly = True
    'End If

    ilMissedType = 0
    slIncludeTitle = ""
    slStr = ""
    ilCostType = 0
    mSetCostType ilCostType                         'set bit map of spot types to include (n/c bonus, adu, etc)
    mSpotSalesTitle ilMissedType, slIncludeTitle, slStr    'slstr won't be used
    If Not gSetFormula("Show Included Cont Types", "'" & slIncludeTitle & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    If RptSelCb!rbcSelCInclude(0).Value Then
        slTitleNetC = "Net*"
        slTitleNetT = "Net*"
    Else
        slTitleNetC = "Net-Net*"
        slTitleNetT = "Net*"
    End If
    If Not gSetFormula("Show Net Terms", "'" & slTitleNetT & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    If RptSelCb!rbcSelC7(0).Value Then    'Order
        ilByOrderOrAir = 0
    ElseIf RptSelCb!rbcSelC7(1).Value Then  'as aired
        ilByOrderOrAir = 1
    Else
        ilByOrderOrAir = 2      'as aired, pkg ordered
    End If
    '12-5-18    Use Actual spot price or the proposal price
    slWhichPrice = "A"          'default to Actual Price
    If RptSelCb!ckcSelC15.Value = vbChecked Then
        slWhichPrice = "P"
    End If
    If Not gSetFormula("WhichPrice", "'" & slWhichPrice & "'") Then
        MsgBox "RptGenCb - error in WhichPrice formula", vbOKOnly, "Report Error"
        Exit Sub
    End If

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
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmAgf
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
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
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
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmAdf
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmAdf
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmCbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmCbf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmCbf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)

    tmVef.iCode = 0
    DoEvents
    
    llContrCode = 0
    slStr = RptSelCb!edcSet1.Text  'see if selective contract entred
    If slStr <> "" Then
        llContrCode = Val(slStr)
        tmChfSrchKey1.lCntrNo = llContrCode
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (llContrCode = tmChf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")            'set the selective contr code only if no errors
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If llContrCode = tmChf.lCntrNo Then
            llContrCode = tmChf.lCode
        Else
            llContrCode = -1                    'get nothing, invalid contr #
        End If
    End If

     ilRet = gObtainSlf(RptSelCb, hmSlf, tlSlf())            '8-15-02 build table of all slsp
    'build table of all selling ofices
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tmSof.iCode
        tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    
    ReDim imProdPct(0 To 1) As Integer            '5-9-07  Participant share
    ReDim imMnfGroup(0 To 1) As Integer           '5-9-07  Participants
    ReDim imMnfSSCode(0 To 1) As Integer          '5-9-07  Particpant sales source


    '5-9-17 if net net, need to get the owners share for splitting
    If RptSelCb!rbcSelCInclude(1).Value Then
        llStartDate = gDateValue(slStartDate)
        blOwnerOnly = False              'only need owners share
        gCreatePIFForRpts llStartDate, tmPifKey(), tmPifPct(), RptSelCb, blOwnerOnly
    End If

'VB6**    slPrinter = LlVBPrintGetPrinter(hdJob)
'VB6**    slPort = LlVBPrintGetPort(hdJob)
    llCAllNoSpots = 0
    slCAllGross = ".00"
    slCAllCommission = ".00"
    slCAllNet = ".00"
    llTAllNoSpots = 0
    slTAllGross = ".00"
    slTAllCommission = ".00"
    slTAllNet = ".00"
    'ReDim imSpotSaleVefCode(1 To 1) As Integer
    ReDim imSpotSaleVefCode(0 To 0) As Integer
    For ilVehicle = 0 To RptSelCb!lbcSelection(3).ListCount - 1 Step 1
        If RptSelCb!lbcSelection(3).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            imSpotSaleVefCode(UBound(imSpotSaleVefCode)) = ilVefCode
            'ReDim Preserve imSpotSaleVefCode(1 To UBound(imSpotSaleVefCode) + 1) As Integer
            ReDim Preserve imSpotSaleVefCode(0 To UBound(imSpotSaleVefCode) + 1) As Integer
        End If
    Next ilVehicle
    ilConvUpper = UBound(imSpotSaleVefCode) - 1
    'Add virtual vehicles if one of its conventional vehicle was selected
    ilRet = btrGetFirst(hmVef, tlVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tlVef.sType = "V" Then
            tmVsfSrchKey.lCode = tlVef.lVsfCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilFound = False
                For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilVsf) > 0 Then
                        For illoop = LBound(imSpotSaleVefCode) To ilConvUpper Step 1
                            If tmVsf.iFSCode(ilVsf) = imSpotSaleVefCode(illoop) Then
                                imSpotSaleVefCode(UBound(imSpotSaleVefCode)) = tlVef.iCode  'tmVsf.iFSCode(ilVsf)
                                'ReDim Preserve imSpotSaleVefCode(1 To UBound(imSpotSaleVefCode) + 1) As Integer
                                ReDim Preserve imSpotSaleVefCode(0 To UBound(imSpotSaleVefCode) + 1) As Integer
                                ilFound = True
                                Exit For
                            End If
                        Next illoop
                        If ilFound Then
                            Exit For
                        End If
                    End If
                Next ilVsf
            End If
        End If
        ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Loop
    ReDim tlspotsale(0 To 0) As SPOTSALE
    'For ilVehicle = 0 To RptSelCb!lbcSelection(3).ListCount - 1 Step 1
    '    If RptSelCb!lbcSelection(3).Selected(ilVehicle) Then
    '        slNameCode = Traffic!lbcVehicle.List(ilVehicle)
    '        ilRet = gParseItem(slNameCode, 1, "\", slName)
    '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
    For ilVehicle = LBound(imSpotSaleVefCode) To UBound(imSpotSaleVefCode) - 1 Step 1
        tmVefSrchKey.iCode = imSpotSaleVefCode(ilVehicle)
        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        llVsfCode = tlVef.lVsfCode
        If ilRet = BTRV_ERR_NONE Then
            'slName = Trim$(tlVef.sname)
            ilVefCode = tlVef.iCode
            'ilVefCode = Val(slCode)
            mObtainSdfBySOF ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilMissedType, ilCostType, ilByOrderOrAir, llContrCode
            mObtainMissedForMG 0, ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilByOrderOrAir, ilCostType, llContrCode
            'Sort key
            llUpper = UBound(tmSpotSOF)
            If llUpper > 0 Then
                ArraySortTyp fnAV(tmSpotSOF(), 0), llUpper, 0, LenB(tmSpotSOF(0)), 0, LenB(tmSpotSOF(0).sKey), 0
            End If
            llCSofNoSpots = 0
            slCSofGross = ".00"
            slCSofCommission = ".00"
            slCSofNet = ".00"
            llTSofNoSpots = 0
            slTSofGross = ".00"
            slTSofCommission = ".00"
            slTSofNet = ".00"
            llCVehNoSpots = 0
            slCVehGross = ".00"
            slCVehCommission = ".00"
            slCVehNet = ".00"
            llTVehNoSpots = 0
            slTVehGross = ".00"
            slTVehCommission = ".00"
            slTVehNet = ".00"
            For llIndex = LBound(tmSpotSOF) To UBound(tmSpotSOF) - 1 Step 1
                tmSdf = tmSpotSOF(llIndex).tSdf
'                    If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then

                    If RptSelCb!rbcSelC4(1).Value Then        'unit count , calc # 30"units
                        ilUorC = 0
                        illoop = tmSdf.iLen
                        Do While illoop >= 30
                            illoop = illoop - 30
                            ilUorC = ilUorC + 1
                        Loop
                        If illoop > 0 And illoop < 30 Then         'round up
                            ilUorC = ilUorC + 1
                        End If
                    Else
                        ilUorC = 1                  'assume real spot count, add 1 for each spot
                    End If
                    ilRet = gParseItem(tmSpotSOF(llIndex).sKey, 4, "|", slSOFName)
                    ilRet = gParseItem(tmSpotSOF(llIndex).sKey, 2, "|", slAdvtName)
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                    llDate = gDateValue(slDate)
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
                        llCntrNo = tmChf.lCntrNo

                        '8-15-02 Obtain the Sales Source from the primary salesperson, first testing the sales office, then getting the sales source
                         ilMatchSS = 0
                         For illoop = LBound(tlSlf) To UBound(tlSlf) - 1
                            If tlSlf(illoop).iCode = tmChf.iSlfCode(0) Then
                                'find the matching sales office
                                For ilTemp = LBound(tlSofList) To UBound(tlSofList)
                                    If tlSofList(ilTemp).iSofCode = tlSlf(illoop).iSofCode Then
                                        ilMatchSS = tlSofList(ilTemp).iMnfSSCode
                                        Exit For
                                    End If
                                Next ilTemp
                            End If
                         Next illoop
                        'gPDNToStr tmChf.sPctTrade, 0, slPctTrade
                        slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                        tmClfSrchKey.lChfCode = tmChf.lCode
                        tmClfSrchKey.iLine = tmSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo   ' 0 show latest version
                        tmClfSrchKey.iPropVer = tmChf.iPropVer  ' 0 show latest version
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        If ilVehicle <= ilConvUpper Then
                            If tmClf.iVefCode = ilVefCode Then
                                'ReDim ilProcVefCode(1 To 2) As Integer
                                'ilProcVefCode(1) = ilVefCode
                                ReDim ilProcVefCode(0 To 1) As Integer
                                ilProcVefCode(0) = ilVefCode
                            Else
                                'Check original vehicle
                                'ReDim ilProcVefCode(1 To 1) As Integer
                                ReDim ilProcVefCode(0 To 0) As Integer
                                If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                    '11-30-04 change to access smf by key2 instead of key0 for speed
                                    'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                                    'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                                    'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                                    'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                                    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    tmSmfSrchKey2.lCode = tmSdf.lCode
                                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                        If (tmSmf.lSdfCode = tmSdf.lCode) Then
                                            If tmClf.iVefCode = tmSmf.iOrigSchVef Then
                                                'ReDim ilProcVefCode(1 To 2) As Integer
                                                'ilProcVefCode(1) = ilVefCode
                                                ReDim ilProcVefCode(0 To 1) As Integer
                                                ilProcVefCode(0) = ilVefCode
                                            End If
                                            Exit Do
                                        End If
                                        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    Loop
                                End If
                            End If
                        Else
                            'Build an array of vehicles to be processed from the virtual vehicle
                            'ReDim ilProcVefCode(1 To 1) As Integer
                            ReDim ilProcVefCode(0 To 0) As Integer
                            tmVsfSrchKey.lCode = llVsfCode
                            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet = BTRV_ERR_NONE Then
                                For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                    If tmVsf.iFSCode(ilVsf) > 0 Then
                                        For illoop = LBound(imSpotSaleVefCode) To ilConvUpper Step 1
                                            If tmVsf.iFSCode(ilVsf) = imSpotSaleVefCode(illoop) Then
                                                For ilPass = 1 To tmVsf.iNoSpots(ilVsf) Step 1
                                                    ilProcVefCode(UBound(ilProcVefCode)) = tmVsf.iFSCode(ilVsf)
                                                    'ReDim Preserve ilProcVefCode(1 To UBound(ilProcVefCode) + 1) As Integer
                                                    ReDim Preserve ilProcVefCode(0 To UBound(ilProcVefCode) + 1) As Integer
                                                Next ilPass
                                                Exit For
                                            End If
                                        Next illoop
                                    End If
                                Next ilVsf
                            End If
                        End If
                        For ilVsf = LBound(ilProcVefCode) To UBound(ilProcVefCode) - 1 Step 1
                            If tlVef.iCode <> ilProcVefCode(ilVsf) Then
                                tmVefSrchKey.iCode = ilProcVefCode(ilVsf)
                                ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            End If
                            'slName = Trim$(tlVef.sname)
                            ilVefCode = tlVef.iCode
                            If Val(slPctTrade) = 100 Then
                                ilStartPass = 2
                                ilEndPass = 2
                            ElseIf Val(slPctTrade) = 0 Then
                                ilStartPass = 1
                                ilEndPass = 1
                            Else
                                ilStartPass = 1
                                ilEndPass = 2
                            End If
                            tmSmf.iOrigSchVef = ilProcVefCode(ilVsf)    'tmSdf.iVefCode
                            'If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                            If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (ilVehicle <= ilConvUpper) Then
                                '11-30-04 access smf by key2 insted of key0 for sped
                                'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                                'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                                'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                                'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                                'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                tmSmfSrchKey2.lCode = tmSdf.lCode
                                ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                    If (tmSmf.lSdfCode = tmSdf.lCode) Then
                                        Exit Do
                                    End If
                                    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                Loop
                            End If
                            For ilPass = ilStartPass To ilEndPass Step 1
                                slGross = ".00"
                                slNet = ".00"
                                slCommission = ".00"
                                If tmSdf.sSpotType <> "X" Then
                                    'Select Case tmClf.sPriceType
                                        'Case "T"    'True

                                            If (tmSdf.sPriceType <> "N") And (tmSdf.sPriceType <> "P") Then
                                                'gPDNToStr tmClf.sActPrice, 2, slGross
'                                                ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slGross)
                                                ilRet = gGetFlightWhichPrice(tmSdf, tmClf, hmCff, hmSmf, slGross, slWhichPrice)
                                                If InStr(slGross, ".") <> 0 Then
                                                    If (ilPass = 1) And (Val(slPctTrade) <> 0) Then
                                                        slGross = gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")
                                                    ElseIf (ilPass = 2) And (Val(slPctTrade) <> 100) Then
                                                        slGross = gSubStr(slGross, gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")) 'gDivStr(gMulStr(slGross, slPctTrade), "100")
                                                    End If
                                                    slGross = gVirtVefSpotPrice(hmVef, hmVsf, tmClf.iVefCode, tmSmf.iOrigSchVef, slGross)
                                                    If (tmChf.iAgfCode > 0) And ((ilPass = 1) Or ((ilPass = 2) And (tmChf.sAgyCTrade = "Y"))) Then   '(Val(slPctTrade) <> 100) Then
                                                        If tmClf.iVefCode <> ilPrevVeh Then
                                                            ilPrevVeh = tmClf.iVefCode
                                                            tmVefSrchKey.iCode = tmClf.iVefCode
                                                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                        End If
                                                        If tmChf.iAgfCode <> tmAgf.iCode Then
                                                            tmAgfSrchKey.iCode = tmChf.iAgfCode
                                                            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                            If ilRet = BTRV_ERR_NONE Then
                                                                'gPDNToStr tmAgf.sComm, 2, slAgyComm
                                                                slAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                                                                slNet = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyComm)), "100.00")
                                                                slCommission = gSubStr(slGross, slNet)
                                                                'adjust net for producers fee if applicable
                                                                mSpotSalesNetNet slNet, ilMatchSS
                                                                'If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
                                                                '    slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                                '    slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                                'End If
                                                            End If
                                                        Else
                                                            'gPDNToStr tmAgf.sComm, 2, slAgyComm
                                                            slAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                                                            slNet = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyComm)), "100.00")
                                                            slCommission = gSubStr(slGross, slNet)
                                                            'adjust for producers commission
                                                             mSpotSalesNetNet slNet, ilMatchSS
                                                            'If RptSelCb!rbcSelCInclude(1).Value Then          'net net option
                                                            '    slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                            '    slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                            'End If
                                                        End If
                                                    Else                             'Not commissionable, handle net-net
                                                        slNet = slGross
                                                        'adjust net for producers fee if applicable
                                                         mSpotSalesNetNet slNet, ilMatchSS
                                                        'If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
                                                        '    slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                        '    slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                        'End If

                                                    End If
                                                End If
                                            End If
                                    'End Select
                                End If
                                If Not ilAdvtOnly Then  'Contract totals
                                    ilSBucket = 1
                                Else
                                    ilSBucket = 1
                                End If
                                '10-6-03 chg from 2 to 1 pass.  Fix overflow/subscript out of range (by Sales Source) for 32000+ records.  Pass 2 doesnt do anything different and not required
                                ilEBucket = 1       '8-26-03 chg from 4 to 2, passes 3 & 4 only create empty arrys and exceeds 32000
                                For ilBucket = ilSBucket To ilEBucket Step 1
                                    llFound = -1
                                    If ilBucket = 1 Then
                                        If Not ilAdvtOnly Then  'Contract totals
                                            llBucketDate = llDate
                                            For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                                If (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) And (tlspotsale(llLoop).lCntrNo = llCntrNo) Then
                                                    llFound = llLoop
                                                    Exit For
                                                End If
                                            Next llLoop
                                        Else
                                            For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                                If (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) And (Trim$(tlspotsale(llLoop).sAdvtName) = Trim$(slAdvtName)) Then
                                                    llFound = llLoop
                                                    Exit For
                                                End If
                                            Next llLoop
                                        End If
                                    ElseIf ilBucket = 2 Then    'Selling Office total within vehicle
                                        llBucketDate = 99890
                                        For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                            If (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) And (Trim$(tlspotsale(llLoop).sSOFName) = Trim$(slSOFName)) And (Val(tlspotsale(llLoop).sDate) = llBucketDate) Then
                                                llFound = llLoop
                                                Exit For
                                            End If
                                        Next llLoop
                                    ElseIf ilBucket = 3 Then    'Vehicle total
                                        llBucketDate = 99900
                                        For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                            If (Val(tlspotsale(llLoop).sDate) = llBucketDate) And (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) Then
                                                llFound = llLoop
                                                Exit For
                                            End If
                                        Next llLoop
                                    Else    'Grand Totals
                                        llBucketDate = 99910
                                        For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                            If (Val(tlspotsale(llLoop).sDate) = llBucketDate) Then
                                                llFound = llLoop
                                                Exit For
                                            End If
                                        Next llLoop
                                    End If
                                    If llFound = -1 Then
                                        ReDim Preserve tlspotsale(0 To UBound(tlspotsale) + 1) As SPOTSALE
                                        llFound = UBound(tlspotsale) - 1

                                        If ilBucket = 1 Then
                                            gUnpackDateForSort tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                        Else
                                            slDate = Trim$(str$(llBucketDate))
                                        End If
                                        If tlVef.iCode <> ilProcVefCode(ilVsf) Then
                                            tmVefSrchKey.iCode = ilProcVefCode(ilVsf)
                                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                        End If
                                        If ilBucket = 1 Then
                                            tlspotsale(llFound).sKey = tlVef.sName & "|" & slSOFName & "|" & slAdvtName & "|" & slDate
                                            tlspotsale(llFound).iVefCode = ilProcVefCode(ilVsf) 'tmSdf.iVefCode
                                            'tlSpotSale(ilFound).sVehName = slName
                                            tlspotsale(llFound).sSOFName = Trim$(slSOFName)
                                            tlspotsale(llFound).sAdvtName = Trim$(slAdvtName)
                                            If RptSelCb!rbcSelCSelect(3).Value = True Then
                                                tlspotsale(llFound).iSofCode = ilMatchSS
                                            Else
                                                tlspotsale(llFound).iSofCode = 0
                                            End If
                                        ElseIf ilBucket = 2 Then    'SOF within Vehicle total
                                            'tlSpotSale(ilFound).sKey = tlVef.sname & "|" & slSOFName & "|" & "|" & Trim$(Str$(llBucketDate))
                                            'tlSpotSale(ilFound).ivefCode = ilProcVefCode(ilVsf) 'tmSdf.iVefCode
                                            'tlSpotSale(ilFound).sVehName = slName
                                            'tlSpotSale(ilFound).sSOFName = Trim$(slSOFName)
                                            'tlSpotSale(ilFound).sAdvtName = "Totals"
                                            ilTtlTest = 1
                                        ElseIf ilBucket = 3 Then    'Vehicle total
                                            'tlSpotSale(ilFound).sKey = tlVef.sname & "|" & "|" & "|" & Trim$(Str$(llBucketDate))
                                            'tlSpotSale(ilFound).ivefCode = ilProcVefCode(ilVsf) 'tmSdf.iVefCode
                                            'tlSpotSale(ilFound).sVehName = slName
                                            'tlSpotSale(ilFound).sSOFName = ""
                                            'tlSpotSale(ilFound).sAdvtName = ""
                                        Else    'Grand total
                                            'tlSpotSale(ilFound).sKey = "~~~~~~~~~~~~~~~~~~~~" & "|" & "|" & "|" & Trim$(Str$(llBucketDate))
                                            'tlSpotSale(ilFound).ivefCode = 0 'tmSdf.iVefCode
                                            'tlSpotSale(ilFound).sVehName = "Grand Totals"
                                            'tlSpotSale(ilFound).sSOFName = ""
                                            'tlSpotSale(ilFound).sAdvtName = ""
                                        End If
                                        If Not ilAdvtOnly Then  'Contract totals
                                            If ilBucket = 1 Then
                                                tlspotsale(llFound).lCntrNo = llCntrNo
                                            Else
                                                tlspotsale(llFound).lCntrNo = 0
                                            End If
                                        Else
                                            tlspotsale(llFound).lCntrNo = 0
                                        End If
                                        If ilBucket = 1 Then
                                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                            llDate = gDateValue(slDate)
                                            tlspotsale(llFound).lDate = llDate
                                            tlspotsale(llFound).sDate = Left$(slDate, Len(slDate) - 3)
                                        Else
                                            tlspotsale(llFound).lDate = 0
                                            tlspotsale(llFound).sDate = Trim$(str$(llBucketDate))
                                        End If
                                        tlspotsale(llFound).lCNoSpots = 0           '5-5-17
                                        tlspotsale(llFound).sCGross = ".00"
                                        tlspotsale(llFound).sCCommission = ".00"
                                        tlspotsale(llFound).sCNet = ".00"
                                        tlspotsale(llFound).lTNoSpots = 0       '5-5-17
                                        tlspotsale(llFound).sTGross = ".00"
                                        tlspotsale(llFound).sTCommission = ".00"
                                        tlspotsale(llFound).sTNet = ".00"
                                    End If
                                    If ilPass = 2 Then  'Val(slPctTrade) = 100 Then
                                        If ilStartPass = 2 Then 'Trade only
                                            tlspotsale(llFound).lTNoSpots = tlspotsale(llFound).lTNoSpots + ilUorC      '5-5-17
                                        End If
                                        tlspotsale(llFound).sTGross = gAddStr(Trim$(tlspotsale(llFound).sTGross), slGross)
                                        tlspotsale(llFound).sTCommission = gAddStr(Trim$(tlspotsale(llFound).sTCommission), slCommission)
                                        tlspotsale(llFound).sTNet = gAddStr(Trim$(tlspotsale(llFound).sTNet), slNet)
                                    Else
                                        tlspotsale(llFound).lCNoSpots = tlspotsale(llFound).lCNoSpots + ilUorC      '5-5-17
                                        tlspotsale(llFound).sCGross = gAddStr(Trim$(tlspotsale(llFound).sCGross), slGross)
                                        tlspotsale(llFound).sCCommission = gAddStr(Trim$(tlspotsale(llFound).sCCommission), slCommission)
                                        tlspotsale(llFound).sCNet = gAddStr(Trim$(tlspotsale(llFound).sCNet), slNet)
                                    End If
                                Next ilBucket

                                If ilPass = 2 Then  'Val(slPctTrade) = 100 Then
                                    If ilStartPass = 2 Then 'Trade only
                                        llTAllNoSpots = llTAllNoSpots + ilUorC
                                    End If
                                    slTAllGross = gAddStr(slTAllGross, slGross)
                                    slTAllCommission = gAddStr(slTAllCommission, slCommission)
                                    slTAllNet = gAddStr(slTAllNet, slNet)
                                Else
                                    llCAllNoSpots = llCAllNoSpots + ilUorC
                                    slCAllGross = gAddStr(slCAllGross, slGross)
                                    slCAllCommission = gAddStr(slCAllCommission, slCommission)
                                    slCAllNet = gAddStr(slCAllNet, slNet)
                                End If
                            Next ilPass
                        Next ilVsf
                    End If
'                    End If
            Next llIndex
            'Sof totals
            If (ilVehicle <= ilConvUpper) And (UBound(tmSpotSOF) - 1 < LBound(tmSpotSOF)) Then
                'ReDim Preserve tlSpotSale(0 To UBound(tlSpotSale) + 1) As SPOTSALE
                'ilFound = UBound(tlSpotSale) - 1
                'tlSpotSale(ilFound).sKey = tlVef.sname & "|" & "|" & "|" & "99900"
                'tlSpotSale(ilFound).ivefCode = ilVefCode    'tmSdf.iVefCode
                ''tlSpotSale(ilFound).sVehName = slName
                'tlSpotSale(ilFound).sSOFName = ""   'Trim$(slSOFName)
                'tlSpotSale(ilFound).sAdvtName = ""  '"Totals"
                'tlSpotSale(ilFound).lCntrNo = 0 'llCntrNo
                'tlSpotSale(ilFound).lDate = 0
                'tlSpotSale(ilFound).sDate = ""
                'tlSpotSale(ilFound).iCNoSpots = llCSofNoSpots
                'tlSpotSale(ilFound).sCGross = slCSofGross
                'tlSpotSale(ilFound).sCCommission = slCSofCommission
                'tlSpotSale(ilFound).sCNet = slCSofNet
                'tlSpotSale(ilFound).iTNoSpots = llTSofNoSpots
                'tlSpotSale(ilFound).sTGross = slTSofGross
                'tlSpotSale(ilFound).sTCommission = slTSofCommission
                'tlSpotSale(ilFound).sTNet = slTSofNet
            End If
        End If
    Next ilVehicle
    'ReDim Preserve tlSpotSale(0 To UBound(tlSpotSale) + 1) As SPOTSALE
    'ilFound = UBound(tlSpotSale) - 1
    'tlSpotSale(ilFound).sKey = "~~~~~~~~~~~~~~~~~~~~" & "|" & "|" & "|" & "99920"
    'tlSpotSale(ilFound).ivefCode = -1
    'tlSpotSale(ilFound).sVehName = "Cash + Trade Totals"
    'tlSpotSale(ilFound).sSOFName = ""
    'tlSpotSale(ilFound).sAdvtName = ""
    'tlSpotSale(ilFound).lCntrNo = 0 'Not used llCntrNo
    'tlSpotSale(ilFound).lDate = 0
    'tlSpotSale(ilFound).sDate = ""
    'tlSpotSale(ilFound).iCNoSpots = llCAllNoSpots + llTAllNoSpots
    'tlSpotSale(ilFound).sCGross = gAddStr(slCAllGross, slTAllGross)
    'tlSpotSale(ilFound).sCCommission = gAddStr(slCAllCommission, slTAllCommission)
    'tlSpotSale(ilFound).sCNet = gAddStr(slCAllNet, slTAllNet)
    'tlSpotSale(ilFound).iTNoSpots = 0
    'tlSpotSale(ilFound).sTGross = ""
    'tlSpotSale(ilFound).sTCommission = ""
    'tlSpotSale(ilFound).sTNet = ""
    'llNoRecsToProc = UBound(tlSpotSale)' - 1
    llUpper = UBound(tlspotsale)
    If llUpper > 0 Then
        ArraySortTyp fnAV(tlspotsale(), 0), llUpper, 0, LenB(tlspotsale(0)), 0, LenB(tlspotsale(0).sKey), 0
    End If
    'outer loop - one loop per page
    llIndex = LBound(tlspotsale)
    If llIndex >= UBound(tlspotsale) Then
        ilDBRet = 1
    Else
        ilDBRet = BTRV_ERR_NONE
        'ilDummy = LLDefineVariableExt(hdJob, "Logo", sgLogoPath & "RptLogo.Bmp", LL_DRAWING, "")
        'ilDummy = LLDefineVariableExtHandle(hdJob, "CSILogo", Traffic!imcCSILogo, LL_DRAWING_HBITMAP, "")
        'ilDummy = LLDefineVariableExt(hdJob, "ReportDates", slDateRange, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "ReportInclude", slIncludeTitle, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "TitleNetC", slTitleNetC, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "TitleNetT", slTitleNetT, LL_TEXT, "")
    End If
    While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0
        ilAnyOutput = True
        ilMaxRowPerPage = 40    '41    'number of rows without lines'LLPrintGetRemainingItemsPerTable(hdJob, ":Ordered")
        ilNoLinesPerPage = 0
        llNoRowsPrt = 0
        ilNewPage = True
'VB6**        ilret = LLPrintEnableObject(hdJob, ":Spots", True)
'VB6**        ilret = LLPrint(hdJob)
'VB6**        ilret = LLPrintEnableObject(hdJob, ":Spots", True)
        For illoop = 0 To 11 Step 1
            slField(illoop) = ""
        Next illoop
        While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0
            If Not ilNewPage Then
                'If (tlSpotSale(ilIndex).ivefCode <> tlSpotSale(ilIndex - 1).ivefCode) Or ((ilNoRowsPrt + 1) > (ilMaxRowPerPage - ilNoLinesPerPage \ 7)) Then
                '    ilDummy = LLDefineFieldExt(hdJob, "Vehicle", slField(0), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "SalesOffice", slField(1), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "AdvtName", slField(2), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CntrNo", slField(3), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CNoSpots", slField(4), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CGross", slField(5), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CCommission", slField(6), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CNet", slField(7), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TNoSpots", slField(8), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TGross", slField(9), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TCommission", slField(10), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TNet", slField(11), LL_TEXT, "")
                '    ilret = LlPrintFields(hdJob)
                '    If ((ilNoRowsPrt + 1) > (ilMaxRowPerPage - ilNoLinesPerPage \ 7)) Then
                '        ilret = LL_WRN_REPEAT_DATA
                '    Else
                '        ilNoLinesPerPage = ilNoLinesPerPage + 1
                '        For ilLoop = 0 To 11 Step 1
                '            slField(ilLoop) = ""
                '        Next ilLoop
                '    End If
                'End If
            Else
                ilNewPage = False
            End If
'VB6**            If ilRet <> LL_WRN_REPEAT_DATA Then
                If slField(0) = "" Then
                    slField(0) = tlspotsale(llIndex).sVehName
                Else
                    If Trim$(tlspotsale(llIndex).sSOFName) <> "" Then
                        slField(0) = slField(0) & Chr$(10)
                    Else
                        ilTtlTest = 1
                        slField(0) = slField(0) & Chr$(10) & "Totals"
                    End If
                End If
                If Trim$(tlspotsale(llIndex).sSOFName) <> "" Then
                    If slField(1) = "" Then
                        slField(1) = tlspotsale(llIndex).sSOFName
                        tmCbf.sSortField2 = tlspotsale(llIndex).sSOFName
                        tmCbf.iMnfGroup = tlspotsale(llIndex).iSofCode
                    Else
                        If tlspotsale(llIndex).sSOFName <> tlspotsale(llIndex - 1).sSOFName Then
                            slField(1) = slField(1) & Chr$(10) & tlspotsale(llIndex).sSOFName
                            tmCbf.sSortField2 = tlspotsale(llIndex).sSOFName
                            tmCbf.iMnfGroup = tlspotsale(llIndex).iSofCode
                        Else
                            slField(1) = slField(1) & Chr$(10)
                        End If
                    End If
                Else
                    If slField(1) = "" Then
                        slField(1) = ""
                    Else
                        slField(1) = slField(1) & Chr$(10)
                    End If
                End If

                If Trim$(tlspotsale(llIndex).sAdvtName) <> "" Then
                    If slField(2) = "" Then
                        slField(2) = tlspotsale(llIndex).sAdvtName
                        'Adv Name
                        tmCbf.sSortField1 = tlspotsale(llIndex).sAdvtName
                    Else
                        If tlspotsale(llIndex).sAdvtName <> tlspotsale(llIndex - 1).sAdvtName Then
                            slField(2) = slField(2) & Chr$(10) & tlspotsale(llIndex).sAdvtName
                            'Adv Name
                            tmCbf.sSortField1 = tlspotsale(llIndex).sAdvtName
                        Else
                            slField(2) = slField(2) & Chr$(10)
                        End If
                    End If
                Else
                    If slField(2) = "" Then
                        slField(2) = ""
                    Else
                        slField(2) = slField(2) & Chr$(10)
                    End If
                End If
                If (Not ilAdvtOnly) And (tlspotsale(llIndex).lCntrNo > 0) Then
                    If slField(3) = "" And InStr(slField(2), "Totals") = 1 And llNoRowsPrt = 1 Then
                        slField(3) = Chr$(10) & Trim$(str$(tlspotsale(llIndex).lCntrNo))
                        'tmCbf.lWeek(10) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                        tmCbf.lWeek(9) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                    ElseIf slField(3) = "" Then
                        slField(3) = Trim$(str$(tlspotsale(llIndex).lCntrNo))
                        'tmCbf.lWeek(10) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                        tmCbf.lWeek(9) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                    Else
                        slField(3) = slField(3) & Chr$(10) & Trim$(str$(tlspotsale(llIndex).lCntrNo))
                        'tmCbf.lWeek(10) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                        tmCbf.lWeek(9) = gStrDecToLong(Trim$(str$(tlspotsale(llIndex).lCntrNo)), 0)
                    End If
                    'If slField(3) = "" Then
                    '    slField(3) = Trim$(Str$(tlSpotSale(llIndex).lCntrno))
                    'Else
                    '    slField(3) = slField(3) & Chr$(10) & Trim$(Str$(tlSpotSale(llIndex).lCntrno))
                    'End If
                Else
                    If slField(3) = "" Then
                        slField(3) = ""
                    Else
                        slField(3) = slField(3) & Chr$(10)
                    End If
                End If
                If tlspotsale(llIndex).lCNoSpots > 0 Then       '5-5-17
                    If slField(4) = "" Then
                        slField(4) = Trim$(str$(tlspotsale(llIndex).lCNoSpots))     '5-5-17
                    Else
                        slField(4) = slField(4) & Chr$(10) & Trim$(str$(tlspotsale(llIndex).lCNoSpots))     '5-5-17
                    End If
                    slStr = Trim$(tlspotsale(llIndex).sCGross)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(5) = "" Then
                        slField(5) = slStr
                    Else
                        slField(5) = slField(5) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llIndex).sCCommission)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(6) = "" Then
                        slField(6) = slStr
                    Else
                        slField(6) = slField(6) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llIndex).sCNet)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(7) = "" Then
                        slField(7) = slStr
                    Else
                        slField(7) = slField(7) & Chr$(10) & slStr
                    End If
                Else
                    If slField(4) = "" Then
                        slField(4) = " "
                    Else
                        slField(4) = slField(4) & Chr$(10) & " "
                    End If
                    If slField(5) = "" Then
                        slField(5) = " "
                    Else
                        slField(5) = slField(5) & Chr$(10) & " "
                    End If
                    If slField(6) = "" Then
                        slField(6) = " "
                    Else
                        slField(6) = slField(6) & Chr$(10) & " "
                    End If
                    If slField(7) = "" Then
                        slField(7) = " "
                    Else
                        slField(7) = slField(7) & Chr$(10) & " "
                    End If
                End If
                If (tlspotsale(llIndex).lTNoSpots > 0) Or (gCompNumberStr(Trim$(tlspotsale(llIndex).sTGross), ".00") <> 0) Then     '5-5-17
                    If (tlspotsale(llIndex).lTNoSpots > 0) Then     '5-5-17
                        If slField(8) = "" Then
                            slField(8) = Trim$(str$(tlspotsale(llIndex).lTNoSpots))     '5-5-17
                        Else
                            slField(8) = slField(8) & Chr$(10) & Trim$(str$(tlspotsale(llIndex).lTNoSpots))     '5-5-17
                        End If
                    Else
                        If slField(8) = "" Then
                            slField(8) = " "
                        Else
                            slField(8) = slField(8) & Chr$(10) & " "
                        End If
                    End If
                    'tmCbf.lWeek(6) = tlspotsale(llIndex).iTNoSpots
                    tmCbf.lWeek(5) = tlspotsale(llIndex).lTNoSpots      '5-5-17
                    slStr = Trim$(tlspotsale(llIndex).sTGross)
                    'tmCbf.lWeek(7) = gStrDecToLong(slStr, 2)
                    tmCbf.lWeek(6) = gStrDecToLong(slStr, 2)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(9) = "" Then
                        slField(9) = slStr
                    Else
                        slField(9) = slField(9) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llIndex).sTCommission)
                    'tmCbf.lWeek(9) = gStrDecToLong(slStr, 2)
                    tmCbf.lWeek(8) = gStrDecToLong(slStr, 2)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(10) = "" Then
                        slField(10) = slStr
                    Else
                        slField(10) = slField(10) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llIndex).sTNet)
                    'tmCbf.lWeek(8) = gStrDecToLong(slStr, 2)
                    tmCbf.lWeek(7) = gStrDecToLong(slStr, 2)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(11) = "" Then
                        slField(11) = slStr
                    Else
                        slField(11) = slField(11) & Chr$(10) & slStr
                    End If
                Else
                    If slField(8) = "" Then
                        slField(8) = " "
                    Else
                        slField(8) = slField(8) & Chr$(10) & " "
                    End If
                    If slField(9) = "" Then
                        slField(9) = " "
                    Else
                        slField(9) = slField(9) & Chr$(10) & " "
                    End If
                    If slField(10) = "" Then
                        slField(10) = " "
                    Else
                        slField(10) = slField(10) & Chr$(10) & " "
                    End If
                    If slField(11) = "" Then
                        slField(11) = " "
                    Else
                        slField(11) = slField(11) & Chr$(10) & " "
                    End If
                End If
                llNoRowsPrt = llNoRowsPrt + 1
                llIndex = llIndex + 1
                llRecNo = llRecNo + 1
                'notify the user (how far have we come?)
                'ilret = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * llRecNo / llNoRecsToProc))
                DoEvents
                'tell L&L to print the table line
                'next data set if no error or warning
                If llIndex >= UBound(tlspotsale) Then
                    ilDBRet = 1
                End If
'VB6**            End If
            '********************
            'tmCbf.iGenTime(0) = igNowTime(0)
            'tmCbf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmCbf.lGenTime = lgNowTime
            tmCbf.iGenDate(0) = igNowDate(0)
            tmCbf.iGenDate(1) = igNowDate(1)
            tmCbf.iVefCode = tlspotsale(llIndex - 1).iVefCode
            'tmCbf.lWeek(1) = tlspotsale(llIndex - 1).lDate
            'tmCbf.lWeek(2) = tlspotsale(llIndex - 1).iCNoSpots
            'tmCbf.lWeek(3) = gStrDecToLong(tlspotsale(llIndex - 1).sCGross, 2)
            'tmCbf.lWeek(4) = gStrDecToLong(tlspotsale(llIndex - 1).sCCommission, 2)
            'tmCbf.lWeek(5) = gStrDecToLong(tlspotsale(llIndex - 1).sCNet, 2)
            tmCbf.lWeek(0) = tlspotsale(llIndex - 1).lDate
            tmCbf.lWeek(1) = tlspotsale(llIndex - 1).lCNoSpots      '5-5-17
            tmCbf.lWeek(2) = gStrDecToLong(tlspotsale(llIndex - 1).sCGross, 2)
            tmCbf.lWeek(3) = gStrDecToLong(tlspotsale(llIndex - 1).sCCommission, 2)
            tmCbf.lWeek(4) = gStrDecToLong(tlspotsale(llIndex - 1).sCNet, 2)
            tmCbf.sDysTms = slTitleNetC         'net or net-net  for report header
            If (ilTtlTest = 0) And tmCbf.iVefCode <> 0 Then
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            End If

            'tmCbf.lWeek(6) = 0
            'tmCbf.lWeek(7) = 0
            'tmCbf.lWeek(8) = 0
            'tmCbf.lWeek(9) = 0
            tmCbf.lWeek(5) = 0
            tmCbf.lWeek(6) = 0
            tmCbf.lWeek(7) = 0
            tmCbf.lWeek(8) = 0
            ilTtlTest = 0
            For illoop = 0 To 11 Step 1
                slField(illoop) = ""
            Next illoop
        Wend  ' inner loop

        'if error or warning: different reactions:
        If ilRet < 0 Then
 'VB6**           If ilRet <> LL_WRN_REPEAT_DATA Then
 'VB6**               ilErrorFlag = ilRet
 'VB6**           End If
        End If
    Wend    ' while not EOF
    'ilret = LLPrintEnableObject(hdJob, ":Spots", True)
    'ilDummy = LLDefineFieldExt(hdJob, "Vehicle", slField(0), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "SalesOffice", slField(1), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "AdvtName", slField(2), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CntrNo", slField(3), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CNoSpots", slField(4), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CGross", slField(5), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CCommission", slField(6), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CNet", slField(7), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TNoSpots", slField(8), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TGross", slField(9), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TCommission", slField(10), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TNet", slField(11), LL_TEXT, "")
    'ilret = LlPrintFields(hdJob)


    Screen.MousePointer = vbDefault
    Erase slField
    Erase ilProcVefCode
    Erase imSpotSaleVefCode
    Erase tlspotsale
    Erase tmSpotSOF
    If RptSelCb!rbcSelCInclude(1).Value Then        '5-9-17
        Erase tmPifPct, tmPifKey
    End If
    
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmCbf)
    btrDestroy hmSmf
    btrDestroy hmAdf
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmSdf
    btrDestroy hmVsf
    btrDestroy hmVef
    btrDestroy hmAgf
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmCbf
    Exit Sub
End Sub

'******************************************************************
'*
'*      Procedure Name:gSpotSalesVehRpt
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Generate Post Log report
'*
'       4/6/99 In addition to report by Ordered or by
'               aired, add option to report as aired
'               for conventionals, but as ordered for
'               packages
'       9/11/00 D.S. Convert to Crystal from Bridge
'       10-22-04 handle more than 32000 spots as well as
'                more than $21million .  To handle more
'               than $21million, create multiple records
'               for each vehicle in groups of $210000
'       11-30-04 access smf by key2 instead of key0 for speed
'       6-25-06 add day selectivity
'*********************************************************************
Sub gSpotSalesVehRpt()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                       ilUpper                       llFoundSS                 *
'*                                                                                        *
'******************************************************************************************
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim llDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim illoop As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim slDateRange As String
    Dim slGross As String
    Dim slNet As String
    Dim slAgyComm As String
    Dim slCommission As String
    Dim llCVehNoSpots As Long
    Dim slCVehGross As String
    Dim slCVehCommission As String
    Dim slCVehNet As String
    Dim llCAllNoSpots As Long
    Dim slCAllGross As String
    Dim slCAllCommission As String
    Dim slCAllNet As String
    Dim llTVehNoSpots As Long
    Dim slTVehGross As String
    Dim slTVehCommission As String
    Dim slTVehNet As String
    Dim llTAllNoSpots As Long
    Dim slTAllGross As String
    Dim slTAllCommission As String
    Dim slTAllNet As String
    Dim ilFound As Integer
    Dim ilVehOnly As Integer
    Dim ilNewPage As Integer
    Dim ilNoLinesPerPage As Integer
    Dim ilNoRowsPrt As Integer
    Dim ilMaxRowPerPage As Integer
    Dim slPctTrade As String
    Dim ilMissedType As Integer
    Dim ilSpotType As Integer
    Dim slIncludeTitle As String
    Dim slTitleNetC As String
    Dim slTitleNetT As String
    ReDim slField(0 To 9) As String
    Dim ilAnyOutput As Integer
    Dim ilPrevVeh As Integer
    Dim ilPass As Integer       'Pass: 1=Cash; 2=Trade
    Dim ilStartPass As Integer
    Dim ilEndPass As Integer
    Dim tlVef As VEF
    Dim ilVsf As Integer
    Dim ilConvUpper As Integer
    Dim ilBucket As Integer
    Dim ilSBucket As Integer
    Dim ilEBucket As Integer
    Dim llBucketDate As Long
    'Dim ilVsfCode As Integer
    Dim llVsfCode As Long
    Dim ilByOrderOrAir As Integer   '0=Order; 1=Aired , 2 =as aired/ pkg ordered
    Dim ilCostType As Integer               'set to -1 to ignore testing of spot types in routine mObtainSdf
    Dim ilUorC As Integer           '# to add for spot count or unit count, calc each new spot
    'ReDim ilProcVefCode(1 To 1) As Integer
    ReDim ilProcVefCode(0 To 0) As Integer
    Dim llContrCode As Long             'selective contr # code
    Dim slStartTime As String           'start time filter entered by user
    Dim slEndTime As String             'end time filter entered by user
    Dim llStartTime As Long             'start time filter entered by user
    Dim llEndTime As Long              'end time filter entered by user
    Dim ilTtlTest As Integer
    Dim ilCntrSpots As Integer
    Dim ilFeedSpots As Integer

    Dim ilTemp As Integer               '8-15-02
    'ReDim tlSlf(1 To 1) As SLF              '8-15-02
    ReDim tlSlf(0 To 0) As SLF              '8-15-02
    Dim ilMatchSS As Integer                'sales source from primary slsp of contract
    Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
    Dim llUpperSDF As Long
    Dim sl21MillionGross As String
    Dim sl21MillionNet As String
    Dim il21MillionTest As Integer
    Dim ilListIndex As Integer      'report index
    Dim blOwnerOnly As Boolean      '5-9-17
    Dim llStartDate As Long         '5-9-17 start date requested
    Dim llLoop As Long              '12-6-17 prevent overflow
    Dim llFound As Long             '12-6-17 prevent overflow
    Dim ilIncludeType As Integer        '1-3-18
    Dim tlCntTypes As CNTTYPES
    Dim slWhichPrice As String * 1      '12-5-18 use Actual price or Prop Price
    'ilCostType = -1                     'ignore spot type testing in mObtain Sdf
    ilPrevVeh = 0                   'init first time thru
    Screen.MousePointer = vbHourglass
    ilListIndex = RptSelCb!lbcRptType.ListIndex
    
    ilCntrSpots = True                  'always include contract spots
    ilFeedSpots = False                 'always exclude feed spots, no $ on them

    'pass start and end time to report header
    slStartTime = RptSelCb!edcSelCTo.Text        'start time
    llStartTime = gTimeToLong(slStartTime, False)
    slEndTime = RptSelCb!edcSelCTo1.Text            'end time
    llEndTime = gTimeToLong(slEndTime, True)

'    slStartDate = RptSelCb!edcSelCFrom.Text   'Start date
    slStartDate = RptSelCb!CSI_CalFrom.Text   'Start date, 9-11-19 use csi calendar control

    If slStartDate = "" Then
        slStartDate = "1/5/1970" 'Monday
    End If
'    slEndDate = RptSelCb!edcSelCFrom1.Text   'End date
    slEndDate = RptSelCb!CSI_CalTo.Text   'End date
    If (StrComp(slEndDate, "TFN", 1) = 0) Or (Len(slEndDate) = 0) Then
        slEndDate = "12/29/2069"    'Sunday
    End If
    slDateRange = "From " & slStartDate & " To " & slEndDate & ", " & slStartTime & "-" & slEndTime
    If Not gSetFormula("Show Date Range", "'" & slDateRange & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If

    If RptSelCb!rbcSelCSelect(1).Value Then
        ilVehOnly = False
    Else
        ilVehOnly = True
    End If

    ilMissedType = 0
    slIncludeTitle = ""
    slStr = ""
    mSetCostType ilCostType                         'set bit map of spot types to include (n/c bonus, adu, etc)
    mSpotSalesTitle ilMissedType, slIncludeTitle, slStr    'slstr won't be used
    If Not gSetFormula("Show Included Cont Types", "'" & slIncludeTitle & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    If RptSelCb!rbcSelCInclude(0).Value Then      'net
        slTitleNetC = "Net*"
        slTitleNetT = "Net*"
    Else
        slTitleNetC = "Net-Net*"
        slTitleNetT = "Net-Net*"
    End If
    If Not gSetFormula("Show Net Terms", "'" & slTitleNetT & "'") Then
        MsgBox "RptGenCt - error in DateRange", vbOKOnly, "Report Error"
        Exit Sub
    End If
    If RptSelCb!rbcSelC7(0).Value Then    'Order
        ilByOrderOrAir = 0
    ElseIf RptSelCb!rbcSelC7(1).Value Then      'aired
        ilByOrderOrAir = 1
    Else
        ilByOrderOrAir = 2      'for conventionals show as aired, for packages show as ordered
    End If
    
    '12-5-18    Use Actual spot price or the proposal price
    slWhichPrice = "A"          'default to Actual Price
    If RptSelCb!ckcSelC15.Value = vbChecked Then
        slWhichPrice = "P"
    End If
    If Not gSetFormula("WhichPrice", "'" & slWhichPrice & "'") Then
        MsgBox "RptGenCb - error in WhichPrice formula", vbOKOnly, "Report Error"
        Exit Sub
    End If
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
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmAgf
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
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
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
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmAdf
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSsf
        btrDestroy hmSmf
        btrDestroy hmAdf
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSsf
        btrDestroy hmSmf
        btrDestroy hmAdf
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

     hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSlf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)

    tmVef.iCode = 0
    DoEvents
    
    llContrCode = 0
    slStr = RptSelCb!edcSet1.Text  'see if selective contract entred
    If slStr <> "" Then
        llContrCode = Val(slStr)
        tmChfSrchKey1.lCntrNo = llContrCode
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (llContrCode = tmChf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")            'set the selective contr code only if no errors
            ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If llContrCode = tmChf.lCntrNo Then
            llContrCode = tmChf.lCode
        Else
            llContrCode = -1                    'get nothing, invalid contr #
        End If
    End If

    ilRet = gObtainSlf(RptSelCb, hmSlf, tlSlf())            '8-15-02 build table of all slsp
    If Not ilRet Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSof
        btrDestroy hmSlf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmAdf
        btrDestroy hmVef
        btrDestroy hmSdf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'build table of all selling ofices
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tmSof.iCode
        tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop

    ReDim imProdPct(0 To 1) As Integer            '5-9-07  Participant share
    ReDim imMnfGroup(0 To 1) As Integer           '5-9-07  Participants
    ReDim imMnfSSCode(0 To 1) As Integer          '5-9-07  Particpant sales source


    '5-9-17 if net net, need to get the owners share for splitting
    If RptSelCb!rbcSelCInclude(1).Value Then
        llStartDate = gDateValue(slStartDate)
        blOwnerOnly = False              'only need owners share
        gCreatePIFForRpts llStartDate, tmPifKey(), tmPifPct(), RptSelCb, blOwnerOnly
    End If
    
    ilIncludeType = True                    '12-29-17 do contract type testing
    tlCntTypes.iHold = True                 'default holds & orders always included
    tlCntTypes.iOrder = True
    tlCntTypes.iStandard = gSetCheck(RptSelCb!ckcSelC6(0).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelCb!ckcSelC6(1).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelCb!ckcSelC6(2).Value)
    tlCntTypes.iDR = gSetCheck(RptSelCb!ckcSelC6(3).Value)
    tlCntTypes.iPI = gSetCheck(RptSelCb!ckcSelC6(4).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelCb!ckcSelC6(5).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelCb!ckcSelC6(6).Value)

    llCAllNoSpots = 0
    slCAllGross = ".00"
    slCAllCommission = ".00"
    slCAllNet = ".00"
    llTAllNoSpots = 0
    slTAllGross = ".00"
    slTAllCommission = ".00"
    slTAllNet = ".00"
    'ReDim imSpotSaleVefCode(1 To 1) As Integer
    ReDim imSpotSaleVefCode(0 To 0) As Integer
    For ilVehicle = 0 To RptSelCb!lbcSelection(3).ListCount - 1 Step 1
        If RptSelCb!lbcSelection(3).Selected(ilVehicle) Then
            slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            imSpotSaleVefCode(UBound(imSpotSaleVefCode)) = ilVefCode
            'ReDim Preserve imSpotSaleVefCode(1 To UBound(imSpotSaleVefCode) + 1) As Integer
            ReDim Preserve imSpotSaleVefCode(0 To UBound(imSpotSaleVefCode) + 1) As Integer
        End If
    Next ilVehicle
    ilConvUpper = UBound(imSpotSaleVefCode) - 1
    'Add virtual vehicles if one of its conventional vehicle was selected
    ilRet = btrGetFirst(hmVef, tlVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tlVef.sType = "V" Then
            tmVsfSrchKey.lCode = tlVef.lVsfCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilFound = False
                For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilVsf) > 0 Then
                        For illoop = LBound(imSpotSaleVefCode) To ilConvUpper Step 1
                            If tmVsf.iFSCode(ilVsf) = imSpotSaleVefCode(illoop) Then
                                imSpotSaleVefCode(UBound(imSpotSaleVefCode)) = tlVef.iCode  'tmVsf.iFSCode(ilVsf)
                                'ReDim Preserve imSpotSaleVefCode(1 To UBound(imSpotSaleVefCode) + 1) As Integer
                                ReDim Preserve imSpotSaleVefCode(0 To UBound(imSpotSaleVefCode) + 1) As Integer
                                ilFound = True
                                Exit For
                            End If
                        Next illoop
                        If ilFound Then
                            Exit For
                        End If
                    End If
                Next ilVsf
            End If
        End If
        ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Loop
    ReDim tlspotsale(0 To 0) As SPOTSALE
    For ilVehicle = LBound(imSpotSaleVefCode) To UBound(imSpotSaleVefCode) - 1 Step 1
        tmVefSrchKey.iCode = imSpotSaleVefCode(ilVehicle)
        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        llVsfCode = tlVef.lVsfCode

        If ilRet = BTRV_ERR_NONE Then
            'slNameCode = Traffic!lbcVehicle.List(ilVehicle)
            'ilRet = gParseItem(slNameCode, 1, "\", slName)
            'ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'ilVefCode = Val(slCode)
            'slName = Trim$(tlVef.sname)
            ilVefCode = tlVef.iCode
            ReDim tmPLSdf(0 To 0) As SPOTTYPESORT
            mObtainSdf ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilSpotType, 3, True, ilMissedType, False, ilCostType, ilByOrderOrAir, ilIncludeType, llContrCode, ilCntrSpots, ilFeedSpots, False, ilListIndex, tlCntTypes
            'tmVef set in mObtainSDF
            mObtainMissedForMG 1, ilVefCode, slStartDate, slEndDate, llStartTime, llEndTime, ilByOrderOrAir, ilCostType, llContrCode
            'Sort is not required
            llUpperSDF = UBound(tmPLSdf)
            If llUpperSDF > 0 Then
                ArraySortTyp fnAV(tmPLSdf(), 0), llUpperSDF, 0, LenB(tmPLSdf(0)), 0, LenB(tmPLSdf(0).sKey), 0
            End If
            llCVehNoSpots = 0
            slCVehGross = ".00"
            slCVehCommission = ".00"
            slCVehNet = ".00"
            llTVehNoSpots = 0
            slTVehGross = ".00"
            slTVehCommission = ".00"
            slTVehNet = ".00"
            For llUpperSDF = LBound(tmPLSdf) To UBound(tmPLSdf) - 1 Step 1
                tmSdf = tmPLSdf(llUpperSDF).tSdf
                If RptSelCb!rbcSelC4(1).Value Then        'unit count , calc # 30"units
                    ilUorC = 0
                    illoop = tmSdf.iLen
                    Do While illoop >= 30
                        illoop = illoop - 30
                        ilUorC = ilUorC + 1
                    Loop
                    If illoop > 0 And illoop < 30 Then         'round up
                        ilUorC = ilUorC + 1
                    End If
                Else
                    ilUorC = 1                  'assume real spot count, add 1 for each spot
                End If
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                llDate = gDateValue(slDate)
                tmChfSrchKey.lCode = tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If (tmChf.sType <> "S") And (tmChf.sType <> "M") And (tmSdf.sSpotType <> "O" And tmSdf.sSpotType <> "C") Then
                    'gPDNToStr tmChf.sPctTrade, 0, slPctTrade

                     '8-15-02 Obtain the Sales Source from the primary salesperson, first testing the sales office, then getting the sales source
                    ilMatchSS = 0
                    For illoop = LBound(tlSlf) To UBound(tlSlf) - 1
                    If tlSlf(illoop).iCode = tmChf.iSlfCode(0) Then
                        'find the matching sales office
                        For ilTemp = LBound(tlSofList) To UBound(tlSofList)
                            If tlSofList(ilTemp).iSofCode = tlSlf(illoop).iSofCode Then
                                ilMatchSS = tlSofList(ilTemp).iMnfSSCode
                                Exit For
                            End If
                        Next ilTemp
                    End If
                    Next illoop
                    slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                    tmClfSrchKey.lChfCode = tmChf.lCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = tmChf.iCntRevNo '
                    tmClfSrchKey.iPropVer = tmChf.iPropVer
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) 'And (tmClf.sSchStatus = "A")
                    '    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE)
                    'Loop
                    'If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> tmChf.lCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
                    '    tmClf.sPriceType = ""
                    'End If
                    'Determine if spots is to be processed
                    If ilVehicle <= ilConvUpper Then
                        If tmClf.iVefCode = ilVefCode Then
                            'ReDim ilProcVefCode(1 To 2) As Integer
                            'ilProcVefCode(1) = ilVefCode
                            ReDim ilProcVefCode(0 To 1) As Integer
                            ilProcVefCode(0) = ilVefCode
                        Else
                            'Check original vehicle
                            'ReDim ilProcVefCode(1 To 1) As Integer
                            ReDim ilProcVefCode(0 To 0) As Integer
                            If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                                '11-30-04 access smf by key2 instead of key0 for speed
                                'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                                'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                                'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                                'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                                'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                tmSmfSrchKey2.lCode = tmSdf.lCode
                                ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                    If (tmSmf.lSdfCode = tmSdf.lCode) Then
                                        If tmClf.iVefCode = tmSmf.iOrigSchVef Then
                                            'ReDim ilProcVefCode(1 To 2) As Integer
                                            'ilProcVefCode(1) = ilVefCode
                                            ReDim ilProcVefCode(0 To 1) As Integer
                                            ilProcVefCode(0) = ilVefCode
                                        End If
                                        Exit Do
                                    End If
                                    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                Loop
                            End If
                        End If
                    Else
                        'Build an array of vehicles to be processed from the virtual vehicle
                        'ReDim ilProcVefCode(1 To 1) As Integer
                        ReDim ilProcVefCode(0 To 0) As Integer
                        tmVsfSrchKey.lCode = llVsfCode
                        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(ilVsf) > 0 Then
                                    For illoop = LBound(imSpotSaleVefCode) To ilConvUpper Step 1
                                        If tmVsf.iFSCode(ilVsf) = imSpotSaleVefCode(illoop) Then
                                            For ilPass = 1 To tmVsf.iNoSpots(ilVsf) Step 1
                                                ilProcVefCode(UBound(ilProcVefCode)) = tmVsf.iFSCode(ilVsf)
                                                'ReDim Preserve ilProcVefCode(1 To UBound(ilProcVefCode) + 1) As Integer
                                                ReDim Preserve ilProcVefCode(0 To UBound(ilProcVefCode) + 1) As Integer
                                            Next ilPass
                                            Exit For
                                        End If
                                    Next illoop
                                End If
                            Next ilVsf
                        End If
                    End If
                    For ilVsf = LBound(ilProcVefCode) To UBound(ilProcVefCode) - 1 Step 1
                        If tlVef.iCode <> ilProcVefCode(ilVsf) Then
                            tmVefSrchKey.iCode = ilProcVefCode(ilVsf)
                            ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        'slName = Trim$(tlVef.sname)
                        ilVefCode = tlVef.iCode
                        If Val(slPctTrade) = 100 Then
                            ilStartPass = 2
                            ilEndPass = 2
                        ElseIf Val(slPctTrade) = 0 Then
                            ilStartPass = 1
                            ilEndPass = 1
                        Else
                            ilStartPass = 1
                            ilEndPass = 2
                        End If
                        tmSmf.iOrigSchVef = ilProcVefCode(ilVsf)    'tmSdf.iVefCode
                        If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (ilVehicle <= ilConvUpper) Then
                            '11-30-04 access smf by key2 instead of key0 for speed
                            'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                            'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                            'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                            'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                            'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            tmSmfSrchKey2.lCode = tmSdf.lCode
                            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                                If (tmSmf.lSdfCode = tmSdf.lCode) Then
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            Loop
                        End If
                        For ilPass = ilStartPass To ilEndPass Step 1
                            slGross = ".00"
                            slNet = ".00"
                            slCommission = ".00"
                            If tmSdf.sSpotType <> "X" Then
                                'Select Case tmClf.sPriceType
                                    'Case "T"    'True
                                        If (tmSdf.sPriceType <> "N") And (tmSdf.sPriceType <> "P") Then
                                            'gPDNToStr tmClf.sActPrice, 2, slGross
'                                            ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slGross)
                                            ilRet = gGetFlightWhichPrice(tmSdf, tmClf, hmCff, hmSmf, slGross, slWhichPrice)
                                            If InStr(slGross, ".") <> 0 Then
                                                If (ilPass = 1) And (Val(slPctTrade) <> 0) Then
                                                    slGross = gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")
                                                ElseIf (ilPass = 2) And (Val(slPctTrade) <> 100) Then
                                                    'slGross = gDivStr(gMulStr(slGross, slPctTrade), "100")
                                                    slGross = gSubStr(slGross, gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")) 'gDivStr(gMulStr(slGross, slPctTrade), "100")
                                                End If
                                                slGross = gVirtVefSpotPrice(hmVef, hmVsf, tmClf.iVefCode, tmSmf.iOrigSchVef, slGross)
                                                'If (tmChf.iAgfCode > 0) And (ilPass = 1) Then    '(Val(slPctTrade) <> 100) Then
                                                If (tmChf.iAgfCode > 0) And ((ilPass = 1) Or ((ilPass = 2) And (tmChf.sAgyCTrade = "Y"))) Then   '(Val(slPctTrade) <> 100) Then
                                                    If tmClf.iVefCode <> ilPrevVeh Then     'only read vehicles when not in mem.
                                                        ilPrevVeh = tmClf.iVefCode
                                                        tmVefSrchKey.iCode = tmClf.iVefCode
                                                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                    End If
                                                    If tmChf.iAgfCode <> tmAgf.iCode Then
                                                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                                                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            'gPDNToStr tmAgf.sComm, 2, slAgyComm
                                                            slAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                                                            slNet = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyComm)), "100.00")
                                                            slCommission = gSubStr(slGross, slNet)
                                                            'adjust net for producers fee if applicable
                                                            mSpotSalesNetNet slNet, ilMatchSS
                                                            'If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
                                                            '    slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                            '    slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                            'End If
                                                        End If
                                                    Else
                                                        'gPDNToStr tmAgf.sComm, 2, slAgyComm
                                                        slAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                                                        slNet = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyComm)), "100.00")
                                                        slCommission = gSubStr(slGross, slNet)
                                                        'adjust net for producers fee if applicable
                                                        mSpotSalesNetNet slNet, ilMatchSS
                                                        'If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
                                                        '    slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                        '    slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                        'End If
                                                    End If
                                                Else   'direct, handle net-net
                                                    slNet = slGross
                                                    'adjust net for producers fee if applicable
                                                    mSpotSalesNetNet slNet, ilMatchSS
                                                    'If RptSelCb!rbcSelCInclude(1).Value Then         'net net option
                                                    '     slAgyComm = gIntToStrDec(tmVef.iProdPct(1), 2)
                                                    '     slNet = gDivStr(gMulStr(slNet, slAgyComm), "100.00")
                                                     'End If

                                                End If
                                            End If
                                        End If
                                'End Select
                            End If
                            If Not ilVehOnly Then
                                ilSBucket = 1
                            Else
                                ilSBucket = 2
                            End If
                            ilEBucket = 3
                            For ilBucket = ilSBucket To ilEBucket Step 1
                                llFound = -1
                                If ilBucket = 1 Then
                                    llBucketDate = llDate
                                    For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                        If (tlspotsale(llLoop).lDate = llBucketDate) And (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) Then
                                            llFound = llLoop
                                            Exit For
                                        End If
                                    Next llLoop
                                ElseIf ilBucket = 2 Then
                                    llBucketDate = 99900
                                    For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                        If (Val(tlspotsale(llLoop).sDate) = llBucketDate) And (tlspotsale(llLoop).iVefCode = ilProcVefCode(ilVsf)) Then
                                            llFound = llLoop
                                            Exit For
                                        End If
                                    Next llLoop
                                Else
                                    llBucketDate = 99910
                                    For llLoop = LBound(tlspotsale) To UBound(tlspotsale) - 1 Step 1
                                        If (Val(tlspotsale(llLoop).sDate) = llBucketDate) Then
                                            llFound = llLoop
                                            Exit For
                                        End If
                                    Next llLoop
                                End If
                                If llFound = -1 Then
                                    ReDim Preserve tlspotsale(0 To UBound(tlspotsale) + 1) As SPOTSALE
                                    llFound = UBound(tlspotsale) - 1

                                    If ilBucket = 1 Then
                                        gUnpackDateForSort tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                    Else
                                        slDate = Trim$(str$(llBucketDate))
                                    End If
                                    If tlVef.iCode <> ilProcVefCode(ilVsf) Then
                                        tmVefSrchKey.iCode = ilProcVefCode(ilVsf)
                                        ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    End If
                                    If ilBucket = 1 Then
                                        tlspotsale(llFound).sKey = tlVef.sName & "|" & "|" & "|" & slDate
                                        tlspotsale(llFound).iVefCode = ilProcVefCode(ilVsf) 'tmSdf.iVefCode
                                        'tlSpotSale(ilFound).sVehName = slName
                                    ElseIf ilBucket = 2 Then
                                        tlspotsale(llFound).sKey = tlVef.sName & "|" & "|" & "|" & Trim$(str$(llBucketDate))
                                        tlspotsale(llFound).iVefCode = ilProcVefCode(ilVsf) 'tmSdf.iVefCode
                                        'tlSpotSale(ilFound).sVehName = slName
                                    Else
                                        tlspotsale(llFound).sKey = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & "|" & "|" & "|" & Trim$(str$(llBucketDate))
                                        tlspotsale(llFound).iVefCode = 0 'tmSdf.iVefCode
                                        tlspotsale(llFound).sVehName = "Grand Totals"
                                    End If
                                    tlspotsale(llFound).sSOFName = ""   'Not used
                                    tlspotsale(llFound).sAdvtName = ""  'Not Used
                                    tlspotsale(llFound).lCntrNo = 0 'Not used
                                    If ilBucket = 1 Then
                                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                                        llDate = gDateValue(slDate)
                                        tlspotsale(llFound).lDate = llDate
                                        tlspotsale(llFound).sDate = Left$(slDate, Len(slDate) - 3)
                                    Else
                                        tlspotsale(llFound).lDate = 0
                                        tlspotsale(llFound).sDate = Trim$(str$(llBucketDate))
                                    End If
                                    tlspotsale(llFound).lCNoSpots = 0       '5-5-17
                                    tlspotsale(llFound).sCGross = ".00"
                                    tlspotsale(llFound).sCCommission = ".00"
                                    tlspotsale(llFound).sCNet = ".00"
                                    tlspotsale(llFound).lTNoSpots = 0       '5-5-17
                                    tlspotsale(llFound).sTGross = ".00"
                                    tlspotsale(llFound).sTCommission = ".00"
                                    tlspotsale(llFound).sTNet = ".00"
                                End If
                                If ilPass = 2 Then  'Val(slPctTrade) = 100 Then
                                    If ilStartPass = 2 Then 'Trade only
                                        tlspotsale(llFound).lTNoSpots = tlspotsale(llFound).lTNoSpots + ilUorC      '5-5-17
                                    End If
                                    tlspotsale(llFound).sTGross = gAddStr(Trim$(tlspotsale(llFound).sTGross), slGross)
                                    tlspotsale(llFound).sTCommission = gAddStr(Trim$(tlspotsale(llFound).sTCommission), slCommission)
                                    tlspotsale(llFound).sTNet = gAddStr(Trim$(tlspotsale(llFound).sTNet), slNet)
                                Else
                                    tlspotsale(llFound).lCNoSpots = tlspotsale(llFound).lCNoSpots + ilUorC      '5-5-17
                                    tlspotsale(llFound).sCGross = gAddStr(Trim$(tlspotsale(llFound).sCGross), slGross)
                                    tlspotsale(llFound).sCCommission = gAddStr(Trim$(tlspotsale(llFound).sCCommission), slCommission)
                                    tlspotsale(llFound).sCNet = gAddStr(Trim$(tlspotsale(llFound).sCNet), slNet)
                                End If
                            Next ilBucket
                            If ilPass = 2 Then  'Val(slPctTrade) = 100 Then
                            '    If ilStartPass = 2 Then 'Trade only
                            '        llTVehNoSpots = llTVehNoSpots + 1
                            '    End If
                            '    slTVehGross = gAddStr(slTVehGross, slGross)
                            '    slTVehCommission = gAddStr(slTVehCommission, slCommission)
                            '    slTVehNet = gAddStr(slTVehNet, slNet)
                                If ilStartPass = 2 Then 'Trade only
                                    llTAllNoSpots = llTAllNoSpots + ilUorC
                                End If
                                slTAllGross = gAddStr(slTAllGross, slGross)
                                slTAllCommission = gAddStr(slTAllCommission, slCommission)
                                slTAllNet = gAddStr(slTAllNet, slNet)
                            Else
                            '    llCVehNoSpots = llCVehNoSpots + ilUorC
                            '    slCVehGross = gAddStr(slCVehGross, slGross)
                            '    slCVehCommission = gAddStr(slCVehCommission, slCommission)
                            '    slCVehNet = gAddStr(slCVehNet, slNet)
                                llCAllNoSpots = llCAllNoSpots + ilUorC
                                slCAllGross = gAddStr(slCAllGross, slGross)
                                slCAllCommission = gAddStr(slCAllCommission, slCommission)
                                slCAllNet = gAddStr(slCAllNet, slNet)
                            End If
                        Next ilPass
                    Next ilVsf
                End If
            Next llUpperSDF
            'If (ilVehicle <= ilConvUpper) And (UBound(tmPLSdf) - 1 < LBound(tmPLSdf)) Then
            '    'Add vehicle totals
            '    ReDim Preserve tlSpotSale(0 To UBound(tlSpotSale) + 1) As SPOTSALE
            '    ilFound = UBound(tlSpotSale) - 1
            '    tlSpotSale(ilFound).sKey = tlVef.sname & "|" & "|" & "|" & "99900"
            '    tlSpotSale(ilFound).ivefCode = ilVefCode
            '    'tlSpotSale(ilFound).sVehName = slName
            '    tlSpotSale(ilFound).sSOFName = ""   'Not used
            '    tlSpotSale(ilFound).sAdvtName = ""  'Not Used
            '    tlSpotSale(ilFound).lCntrNo = 0 'Not used
            '    tlSpotSale(ilFound).lDate = 0
            '    tlSpotSale(ilFound).sDate = ""
            '    tlSpotSale(ilFound).iCNoSpots = llCVehNoSpots
            '    tlSpotSale(ilFound).sCGross = slCVehGross
            '    tlSpotSale(ilFound).sCCommission = slCVehCommission
            '    tlSpotSale(ilFound).sCNet = slCVehNet
            '    tlSpotSale(ilFound).iTNoSpots = llTVehNoSpots
            '    tlSpotSale(ilFound).sTGross = slTVehGross
            '    tlSpotSale(ilFound).sTCommission = slTVehCommission
            '    tlSpotSale(ilFound).sTNet = slTVehNet
            'End If
        End If
    Next ilVehicle

    'final totals don't need  this
    'ReDim Preserve tlSpotSale(0 To UBound(tlSpotSale) + 1) As SPOTSALE
    'ilFound = UBound(tlSpotSale) - 1
    'tlSpotSale(ilFound).sKey = "~~~~~~~~~~~~~~~~~~~~" & "|" & "|" & "|" & "99920"
    'tlSpotSale(ilFound).ivefCode = -1
    'tlSpotSale(ilFound).sVehName = "Cash + Trade Totals"
    'tlSpotSale(ilFound).sSOFName = ""
    'tlSpotSale(ilFound).sAdvtName = ""
    'tlSpotSale(ilFound).lCntrNo = 0 'Not used
    'tlSpotSale(ilFound).lDate = 0
    'tlSpotSale(ilFound).sDate = ""
    'tlSpotSale(ilFound).iCNoSpots = llCAllNoSpots + llTAllNoSpots
    'tlSpotSale(ilFound).sCGross = gAddStr(slCAllGross, slTAllGross)
    'tlSpotSale(ilFound).sCCommission = gAddStr(slCAllCommission, slTAllCommission)
    'tlSpotSale(ilFound).sCNet = gAddStr(slCAllNet, slTAllNet)
    'tlSpotSale(ilFound).iTNoSpots = 0
    'tlSpotSale(ilFound).sTGross = ""
    'tlSpotSale(ilFound).sTCommission = ""
    'tlSpotSale(ilFound).sTNet = ""
    'llNoRecsToProc = UBound(tlSpotSale)' - 1
    llUpperSDF = UBound(tlspotsale)
    If llUpperSDF > 0 Then
        ArraySortTyp fnAV(tlspotsale(), 0), llUpperSDF, 0, LenB(tlspotsale(0)), 0, LenB(tlspotsale(0).sKey), 0
    End If
    'outer loop - one loop per page
    llUpperSDF = LBound(tlspotsale)
'put in for loop write a record for each record tlSpotSale
    If llUpperSDF >= UBound(tlspotsale) Then
        ilDBRet = 1
    Else
        ilDBRet = BTRV_ERR_NONE
        'ilDummy = LLDefineVariableExt(hdJob, "Logo", sgLogoPath & "RptLogo.Bmp", LL_DRAWING, "")
        'ilDummy = LLDefineVariableExtHandle(hdJob, "CSILogo", Traffic!imcCSILogo, LL_DRAWING_HBITMAP, "")
        'ilDummy = LLDefineVariableExt(hdJob, "ReportDates", slDateRange, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "ReportInclude", slIncludeTitle, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "TitleNetC", slTitleNetC, LL_TEXT, "")
        'ilDummy = LLDefineVariableExt(hdJob, "TitleNetT", slTitleNetT, LL_TEXT, "")
    End If
    While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0
        ilAnyOutput = True
        ilMaxRowPerPage = 49    '50    'number of rows without lines'LLPrintGetRemainingItemsPerTable(hdJob, ":Ordered")
        ilNoLinesPerPage = 0
        ilNoRowsPrt = 0
        ilNewPage = True
'VB6**        ilret = LLPrintEnableObject(hdJob, ":Spots", True)
'VB6**        ilret = LLPrint(hdJob)
'VB6**        ilret = LLPrintEnableObject(hdJob, ":Spots", True)
        For illoop = 0 To 9 Step 1
            slField(illoop) = ""
        Next illoop
        While (ilDBRet = BTRV_ERR_NONE) And ilErrorFlag = 0
            If Not ilNewPage Then
                'If (tlSpotSale(ilIndex).ivefCode <> tlSpotSale(ilIndex - 1).ivefCode) Or ((ilNoRowsPrt + 1) > (ilMaxRowPerPage - ilNoLinesPerPage \ 7)) Then
                '    ilDummy = LLDefineFieldExt(hdJob, "Vehicle", slField(0), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "Date", slField(1), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CNoSpots", slField(2), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CGross", slField(3), LL_TEXT, "")
                '   ilDummy = LLDefineFieldExt(hdJob, "CCommission", slField(4), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "CNet", slField(5), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TNoSpots", slField(6), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TGross", slField(7), LL_TEXT, "")
                '    'ilDummy = LLDefineFieldExt(hdJob, "TCommission", slField(8), LL_TEXT, "")
                '    ilDummy = LLDefineFieldExt(hdJob, "TNet", slField(9), LL_TEXT, "")
                '    'ilret = LlPrintFields(hdJob)
                '    If ((ilNoRowsPrt + 1) > (ilMaxRowPerPage - ilNoLinesPerPage \ 7)) Then
                '        ilret = LL_WRN_REPEAT_DATA
                '    Else
                '        ilNoLinesPerPage = ilNoLinesPerPage + 1
                '        For ilLoop = 0 To 9 Step 1
                '            slField(ilLoop) = ""
                '        Next ilLoop
                '    End If
                'End If
            Else
                ilNewPage = False
            End If
'VB6**            If ilRet <> LL_WRN_REPEAT_DATA Then
                If slField(0) = "" Then
                    slField(0) = Trim$(tlspotsale(llUpperSDF).sVehName)
                Else
                    If tlspotsale(llUpperSDF).lDate > 0 Then
                        slField(0) = slField(0) & Chr$(10)
                    Else
                        ilTtlTest = 1
                        slField(0) = slField(0) & Chr$(10) & "Totals"
                    End If
                End If
                If tlspotsale(llUpperSDF).lDate > 0 Then
                    If slField(1) = "" Then
                        slField(1) = tlspotsale(llUpperSDF).sDate
                    Else
                        slField(1) = slField(1) & Chr$(10) & tlspotsale(llUpperSDF).sDate
                    End If
                Else
                    If slField(1) = "" Then
                        slField(1) = ""
                    Else
                        slField(1) = slField(1) & Chr$(10)
                    End If
                End If
                gPackDateLong tlspotsale(llUpperSDF).lDate, tmGrf.iStartDate(0), tmGrf.iStartDate(1)
                If tlspotsale(llUpperSDF).lCNoSpots > 0 Then        '5-5-17
                    If slField(2) = "" Then
                        slField(2) = Trim$(str$(tlspotsale(llUpperSDF).lCNoSpots))      '5-5-17
                    Else
                        slField(2) = slField(2) & Chr$(10) & Trim$(str$(tlspotsale(llUpperSDF).lCNoSpots))      '5-5-17
                    End If
                    slStr = Trim$(tlspotsale(llUpperSDF).sCGross)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(3) = "" Then
                        slField(3) = slStr
                    Else
                        slField(3) = slField(3) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llUpperSDF).sCCommission)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(4) = "" Then
                        slField(4) = slStr
                    Else
                        slField(4) = slField(4) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llUpperSDF).sCNet)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(5) = "" Then
                        slField(5) = slStr
                    Else
                        slField(5) = slField(5) & Chr$(10) & slStr
                    End If
                Else
                    If slField(2) = "" Then
                        slField(2) = " "
                    Else
                        slField(2) = slField(2) & Chr$(10) & " "
                    End If
                    If slField(3) = "" Then
                        slField(3) = " "
                    Else
                        slField(3) = slField(3) & Chr$(10) & " "
                    End If
                    If slField(4) = "" Then
                        slField(4) = " "
                    Else
                        slField(4) = slField(4) & Chr$(10) & " "
                    End If
                    If slField(5) = "" Then
                        slField(5) = " "
                    Else
                        slField(5) = slField(5) & Chr$(10) & " "
                    End If
                End If
       '***********
                If (tlspotsale(llUpperSDF).lTNoSpots > 0) Or (gCompNumberStr(Trim$(tlspotsale(llUpperSDF).sTGross), ".00") <> 0) Then       '5-5-17
                    If (tlspotsale(llUpperSDF).lTNoSpots > 0) Then      '5-5-17
                        If slField(6) = "" Then
                            slField(6) = Trim$(str$(tlspotsale(llUpperSDF).lTNoSpots))      '5-5-17
                        Else
                            slField(6) = slField(6) & Chr$(10) & Trim$(str$(tlspotsale(llUpperSDF).lTNoSpots))      '5-5-17
                        End If
                    Else
                        If slField(6) = "" Then
                            slField(6) = " "
                        Else
                            slField(6) = slField(6) & Chr$(10) & " "
                        End If
                    End If
       'numspots
       'Trim$(Str$(tlSpotSale(llUpperSDF).iTNoSpots
       'tmGrf.lDollars(6) = tlspotsale(llUpperSDF).iTNoSpots
       tmGrf.lDollars(5) = tlspotsale(llUpperSDF).lTNoSpots     '5-5-17
                    slStr = Trim$(tlspotsale(llUpperSDF).sTGross)
        'gross
        'slstr
        'tmGrf.lDollars(7) = gStrDecToLong(slStr, 2)
        tmGrf.lDollars(6) = gStrDecToLong(slStr, 2)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(7) = "" Then
                        slField(7) = slStr
                    Else
                        slField(7) = slField(7) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llUpperSDF).sTCommission)
        'tmGrf.lDollars(9) = gStrDecToLong(slStr, 2)
        tmGrf.lDollars(8) = gStrDecToLong(slStr, 2)
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(8) = "" Then
                        slField(8) = slStr
                    Else
                        slField(8) = slField(8) & Chr$(10) & slStr
                    End If
                    slStr = Trim$(tlspotsale(llUpperSDF).sTNet)
        'tmGrf.lDollars(8) = gStrDecToLong(slStr, 2)
        tmGrf.lDollars(7) = gStrDecToLong(slStr, 2)
        'Net
        'slstr
                    gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, slStr
                    If slField(9) = "" Then
                        slField(9) = slStr
                    Else
                        slField(9) = slField(9) & Chr$(10) & slStr
                    End If
                Else
                    If slField(6) = "" Then
                        slField(6) = " "
                    Else
                        slField(6) = slField(6) & Chr$(10) & " "
                    End If
                    If slField(7) = "" Then
                        slField(7) = " "
                    Else
                        slField(7) = slField(7) & Chr$(10) & " "
                    End If
                    If slField(8) = "" Then
                        slField(8) = " "
                    Else
                        slField(8) = slField(8) & Chr$(10) & " "
                    End If
                    If slField(9) = "" Then
                        slField(9) = " "
                    Else
                        slField(9) = slField(9) & Chr$(10) & " "
                    End If
                End If
'                ilNoRowsPrt = ilNoRowsPrt + 1
                llUpperSDF = llUpperSDF + 1
                llRecNo = llRecNo + 1
                'notify the user (how far have we come?)
                'ilret = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * llRecNo / llNoRecsToProc))
                DoEvents
                'tell L&L to print the table line
                'next data set if no error or warning
                If llUpperSDF >= UBound(tlspotsale) Then
                    ilDBRet = 1
                End If
'VB6**            End If
            '********************
            'tmGrf.iGenTime(0) = igNowTime(0)
            'tmGrf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            tmGrf.iVefCode = tlspotsale(llUpperSDF - 1).iVefCode
            'tmGrf.lDollars(1) = tlspotsale(llUpperSDF - 1).lDate
            'tmGrf.lDollars(2) = tlspotsale(llUpperSDF - 1).iCNoSpots

            ''agency commission wont be over $21million
            'tmGrf.lDollars(4) = gStrDecToLong(tlspotsale(llUpperSDF - 1).sCCommission, 2)

            tmGrf.lDollars(0) = tlspotsale(llUpperSDF - 1).lDate
            tmGrf.lDollars(1) = tlspotsale(llUpperSDF - 1).lCNoSpots        '5-5-17

            'agency commission wont be over $21million
            tmGrf.lDollars(3) = gStrDecToLong(tlspotsale(llUpperSDF - 1).sCCommission, 2)

            sl21MillionGross = tlspotsale(llUpperSDF - 1).sCGross
            sl21MillionNet = tlspotsale(llUpperSDF - 1).sCNet
            il21MillionTest = True
            Do While il21MillionTest = True
                'Create 1 record for each $21million of the gross.
                'test if vehicle gross is over $21million.  put anything over $21 million into another field to
                'be combined in crystal report
                If gCompNumberStr(sl21MillionGross, "21000000.00") > 0 Then    'vehicle total more than $21million
                    'over 21million
                    'tmGrf.lDollars(3) = 2100000000      '$21million (with pennies)
                    tmGrf.lDollars(2) = 2100000000      '$21million (with pennies)
                    sl21MillionGross = gSubStr(sl21MillionGross, "21000000.00")
                Else     'less or equal to $21 million
                    'tmGrf.lDollars(3) = gStrDecToLong(sl21MillionGross, 2)
                    tmGrf.lDollars(2) = gStrDecToLong(sl21MillionGross, 2)
                    sl21MillionGross = ".00"
                    il21MillionTest = False
                End If

                If gCompNumberStr(sl21MillionNet, "21000000.00") > 0 Then    'vehicle total more than $21million
                    'tmGrf.lDollars(5) = 2100000000      '$21million (with pennies)
                    tmGrf.lDollars(4) = 2100000000      '$21million (with pennies)
                    sl21MillionNet = gSubStr(sl21MillionNet, "21000000.00")
                Else
                    'tmGrf.lDollars(5) = gStrDecToLong(sl21MillionNet, 2)
                    tmGrf.lDollars(4) = gStrDecToLong(sl21MillionNet, 2)
                    sl21MillionNet = ".00"
                End If

                If ilVehOnly = False Then   ' Crystal flag to show totals or not
                    'tmGrf.lDollars(10) = 0
                    tmGrf.lDollars(9) = 0
                Else
                    'tmGrf.lDollars(10) = 1
                    tmGrf.lDollars(9) = 1
                End If
                tmGrf.sGenDesc = slTitleNetC         'net or net-net  for report header

                If (RptSelCb!rbcSelCSelect(0).Value) Then       'none (vs date, advt, sales source)
                    If (ilTtlTest = 0) Or tmGrf.iVefCode <> 0 Then
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If

                If Not (RptSelCb!rbcSelCSelect(0).Value) Then   'none (vs date, advt, sales source)
                    If (ilTtlTest = 0) Or (RptSelCb!rbcSelCSelect(2).Value = True) And tmGrf.iVefCode <> 0 Then
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If
                'repeeat for next set of $21 million; init the other counting fields to avoid overstating values
'                tmGrf.lDollars(2) = 0           'cash # spots
'                tmGrf.lDollars(3) = 0           'cash gross
'                tmGrf.lDollars(4) = 0           'cash comm
'                tmGrf.lDollars(5) = 0           'cash net
'                tmGrf.lDollars(6) = 0           'trade spots
'                tmGrf.lDollars(7) = 0           'trade gross
'                tmGrf.lDollars(8) = 0           'trade net
'                tmGrf.lDollars(9) = 0           'trade comm
                tmGrf.lDollars(1) = 0           'cash # spots
                tmGrf.lDollars(2) = 0           'cash gross
                tmGrf.lDollars(3) = 0           'cash comm
                tmGrf.lDollars(4) = 0           'cash net
                tmGrf.lDollars(5) = 0           'trade spots
                tmGrf.lDollars(6) = 0           'trade gross
                tmGrf.lDollars(7) = 0           'trade net
                tmGrf.lDollars(8) = 0           'trade comm
            Loop



            ilTtlTest = 0
        Wend  ' inner loop

        'if error or warning: different reactions:
        If ilRet < 0 Then
'VB6**            If ilRet <> LL_WRN_REPEAT_DATA Then
'VB6**                ilErrorFlag = ilRet
'VB6**            End If
        End If
    Wend    ' while not EOF
    'ilret = LLPrintEnableObject(hdJob, ":Spots", True)
    'ilDummy = LLDefineFieldExt(hdJob, "Vehicle", slField(0), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "Date", slField(1), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CNoSpots", slField(2), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CGross", slField(3), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CCommission", slField(4), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "CNet", slField(5), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TNoSpots", slField(6), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TGross", slField(7), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TCommission", slField(8), LL_TEXT, "")
    'ilDummy = LLDefineFieldExt(hdJob, "TNet", slField(9), LL_TEXT, "")
    'ilret = LlPrintFields(hdJob)


    Screen.MousePointer = vbDefault
    Erase slField
    Erase ilProcVefCode
    Erase imSpotSaleVefCode
    Erase tlspotsale
    Erase tmPLSdf
    If RptSelCb!rbcSelCInclude(1).Value Then        '5-9-17
        Erase tmPifPct, tmPifKey
    End If

    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmGrf)
    btrDestroy hmSsf
    btrDestroy hmSmf
    btrDestroy hmAdf
    btrDestroy hmSdf
    btrDestroy hmVsf
    btrDestroy hmVef
    btrDestroy hmAgf
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmGrf
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntrSchdSpotChk                *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check if a scheduled contract   *
'*                     contains the correct number of  *
'*                     spots--see gCntrSchdSpotChk     *
'*                                                     *
'*          9/2/97 dh Fix to use multiple DP definitions
'*          4-6-05 ignore BB spots for dispcrepancy
'*******************************************************
Function mCntrSchdSpotChk(ilDispOnly As Integer, ilClf As Integer, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'
'   ilRet = gCntrSchdSpotChk(ilClf As Integer, slStartDate)
'   Where:
'       ilDispOnly(I)- True=Discrepancy contracts only, False=All lines that fall within data span
'       ilClf(I)- Index into tgClfCB (which contains the line to be checked)
'                 tgCffCB must contain the flights for the line
'       slStartDate(I)- start date of check
    Dim ilRet As Integer
    Dim slDate As String
    Dim slCffStartDate As String
    Dim slCffEndDate As String
    Dim slSdfDate As String
    Dim slSchDate As String
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim llLnEarliestDate As Long
    Dim llLnLatestDate As Long
    Dim llMonDate As Long
    Dim llSunDate As Long
    Dim llDate As Long
    Dim llChkStartDate As Long
    Dim llChkEndDate As Long
    Dim ilCff As Integer
    Dim ilCffSpots As Integer
    Dim ilSdfSpots As Integer
    'Dim ilSdfIndex As Integer
    Dim llSdfIndex As Long
    Dim ilDay As Integer
    Dim slSdfTime As String
    Dim llSdfTime As Long
    ReDim llStartTime(0 To 6) As Long
    ReDim llEndTime(0 To 6) As Long
    Dim slOrigMissedDate As String
    Dim ilCVsf As Integer
    Dim ilVefFound As Integer
    Dim ilVefCode As Integer
    Dim ilTDay As Integer
    Dim ilDateFound As Integer
    Dim ilTime As Integer
    Dim ilFound As Integer
    Dim ilFlag As Integer
    'ReDim tlSdfExt(1 To 1) As SDFEXT
    'Get tmCRdf
    tmRdfSrchKey.iCode = tmClf.iRdfCode
    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        mCntrSchdSpotChk = False
        Exit Function
    End If
    tmVefSrchKey.iCode = tmClf.iVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mCntrSchdSpotChk = False
        Exit Function
    End If
    If tmVef.sType = "V" Then
        tmVsfSrchKey.lCode = tmVef.lVsfCode
        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            mCntrSchdSpotChk = False
            Exit Function
        End If
        For ilCVsf = UBound(tmVsf.iFSCode) To LBound(tmVsf.iFSCode) + 1 Step -1
            tmVsf.iFSCode(ilCVsf) = tmVsf.iFSCode(ilCVsf - 1)
            tmVsf.iNoSpots(ilCVsf) = tmVsf.iNoSpots(ilCVsf - 1)
        Next ilCVsf
        tmVsf.iFSCode(LBound(tmVsf.iFSCode)) = tmClf.iVefCode
        tmVsf.iNoSpots(LBound(tmVsf.iFSCode)) = 1
    Else
        For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
            tmVsf.iFSCode(ilCVsf) = 0
        Next ilCVsf
        tmVsf.iFSCode(LBound(tmVsf.iFSCode)) = tmClf.iVefCode
        tmVsf.iNoSpots(LBound(tmVsf.iFSCode)) = 1
    End If
    'Check all missed spots times- if in error correct
    ilCff = tgClfCB(ilClf).iFirstCff
    mCntrSchdSpotChk = True
    llLnEarliestDate = 0
    llLnLatestDate = 0
    Do While ilCff <> -1
        gUnpackDate tgCffCB(ilCff).CffRec.iStartDate(0), tgCffCB(ilCff).CffRec.iStartDate(1), slCffStartDate
        gUnpackDate tgCffCB(ilCff).CffRec.iEndDate(0), tgCffCB(ilCff).CffRec.iEndDate(1), slCffEndDate
        llCffStartDate = gDateValue(slCffStartDate)
        llCffEndDate = gDateValue(slCffEndDate)
        If llLnEarliestDate = 0 Then
            llLnEarliestDate = llCffStartDate
            llLnLatestDate = llCffEndDate
        Else
            If llCffStartDate < llLnEarliestDate Then
                llLnEarliestDate = llCffStartDate
            End If
            If llCffEndDate > llLnLatestDate Then
                llLnLatestDate = llCffEndDate
            End If
        End If
        If llCffEndDate < llCffStartDate Then   'Cancel before start
            For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                If tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
                    If RptSelCb!ckcSelC3(0).Value = vbChecked Then        '9-08-05
                        mCntrSchdSpotChk = False                        'consider it an error , show it
                    Else
                        gUnpackDateLong tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), llChkStartDate
                        If (llChkStartDate >= llStartDate) And (llChkStartDate <= llEndDate) Then
                            mCntrSchdSpotChk = False
                        End If
                    End If
                    'If (tlSdfExt(llSdfIndex).sSpotType <> "X") Then       '8-17-05
                    '    mCntrSchdSpotChk = False        'only if not a fill, set it to discrep
                    'End If
                    tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1  'Date invalid
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        '11-30-04 access smf by key2 instead of key0 for speed
                        'Obtain original dates
                        'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                        'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                        'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                        'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                        'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        tmSmfSrchKey2.lCode = tmSdf.lCode
                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                            If tmSmf.lSdfCode = tmSdf.lCode Then
                                gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                                tlSdfExt(llSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        Loop
                    End If
                End If
            Next llSdfIndex
            If (llCffStartDate >= llStartDate) And (llCffEndDate <= llEndDate) Then
                If Not ilDispOnly Then
                    mCntrSchdSpotChk = False
                End If
            End If
            Exit Function
        End If
        slDate = Format$(llCffStartDate, "m/d/yy")
        slDate = gObtainPrevMonday(slDate)  'First flight start date might be within week- back up to monday
        llMonDate = gDateValue(slDate)
        slDate = gObtainNextSunday(slDate)
        llSunDate = gDateValue(slDate)
        If llSunDate > llCffEndDate Then
            llSunDate = llCffEndDate
        End If
        Do  'Current line week loop
            If llMonDate > llEndDate Then
                Exit Do
            End If
            If llSunDate >= llStartDate Then
                If Not ilDispOnly Then
                    'Show line as it has dates within scan
                    mCntrSchdSpotChk = False
                End If
                If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then  'Weekly buy
                    ilCffSpots = tgCffCB(ilCff).CffRec.iSpotsWk + tgCffCB(ilCff).CffRec.iXSpotsWk
                    llChkStartDate = llMonDate
                    llChkEndDate = llSunDate
                    '6/8/16: Replaced GoSub
                    'GoSub lObtainCount
                    If Not mObtainCount(tlSdfExt(), slOrigMissedDate, llDate, llStartDate, llEndDate, llChkStartDate, llChkEndDate, slSchDate, slSdfTime, llSdfTime, ilCff, llStartTime(), llEndTime(), ilFlag, ilCffSpots) Then
                        mCntrSchdSpotChk = False
                    End If
                Else    'Daily buy
                    ilCffSpots = 0
                    For ilTDay = 0 To 6 Step 1
                        If (llMonDate + ilTDay >= llStartDate) And (llMonDate + ilTDay <= llEndDate) Then
                            ilCffSpots = tgCffCB(ilCff).CffRec.iDay(ilTDay)
                            llChkStartDate = llMonDate + ilTDay
                            llChkEndDate = llMonDate + ilTDay
                            '6/8/16: Replaced GoSub
                            'GoSub lObtainCount
                            If Not mObtainCount(tlSdfExt(), slOrigMissedDate, llDate, llStartDate, llEndDate, llChkStartDate, llChkEndDate, slSchDate, slSdfTime, llSdfTime, ilCff, llStartTime(), llEndTime(), ilFlag, ilCffSpots) Then
                                mCntrSchdSpotChk = False
                            End If
                        End If
                    Next ilTDay
                End If
            End If
            llMonDate = llSunDate + 1
            slDate = Format$(llMonDate, "m/d/yy")
            slDate = gObtainNextSunday(slDate)
            llSunDate = gDateValue(slDate)
            If llSunDate > llCffEndDate Then
                llSunDate = llCffEndDate
            End If
        Loop While llMonDate <= llCffEndDate
        ilCff = tgCffCB(ilCff).iNextCff
    Loop
    'If all flights are within date span- then extra spots can be checked
    'Any spot with a matching line number was not accounted for as the line number is
    'set to a negative when referenced
    'If (llLnEarliestDate >= llStartDate) And (llLnLatestDate <= llEndDate) Then
    '    For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
    '        If tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine Then
    '            mCntrSchdSpotChk = False
    '            tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1  'Date invalid
    '        End If
    '    Next ilSdfIndex
    'End If
    For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
            gUnpackDateLong tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ilVefFound = False
                If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    'tmSmfSrchKey.lChfCode = tmClf.lChfCode
                    'tmSmfSrchKey.iLineNo = tmClf.iLine
                    ''slDate = Format$(llChkStartDate, "m/d/yy")
                    ''gPackDate slDate, ilDate0, ilDate1
                    'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                    'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    'DL: 4-20-05 use key2 instead of key0 for speed
                    tmSmfSrchKey2.lCode = tmSdf.lCode
                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                        If tmSmf.lSdfCode = tmSdf.lCode Then
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                            tlSdfExt(llSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                            For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(ilCVsf) > 0 Then
                                    If tmSmf.iOrigSchVef = tmVsf.iFSCode(ilCVsf) Then
                                        ilVefFound = True
                                        Exit For
                                    End If
                                End If
                            Next ilCVsf
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    Loop
                Else
                    For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilCVsf) > 0 Then
                            If tlSdfExt(llSdfIndex).iVefCode = tmVsf.iFSCode(ilCVsf) Then
                                ilVefFound = True
                                Exit For
                            End If
                        End If
                    Next ilCVsf
                End If
                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If tmSdf.sTracer = "*" Then
                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
                        ilVefFound = True
                    End If
                End If
                If Not ilVefFound Then
                'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                    If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H4  'Vehicle invalid
                        mCntrSchdSpotChk = False
                    End If
                End If
                '6/8/16: Replaced GoSub
                'GoSub lSetStatus
                If Not mSetStatus(tlSdfExt(), llSdfIndex) Then
                    mCntrSchdSpotChk = False
                End If
                If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                    If tmSmf.lSdfCode = tmSdf.lCode Then
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                        tlSdfExt(llSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                        'Test Date of original missed date
                        ilDateFound = False
                        ilCff = tgClfCB(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDateLong tgCffCB(ilCff).CffRec.iStartDate(0), tgCffCB(ilCff).CffRec.iStartDate(1), llCffStartDate
                            gUnpackDateLong tgCffCB(ilCff).CffRec.iEndDate(0), tgCffCB(ilCff).CffRec.iEndDate(1), llCffEndDate
                            If (tlSdfExt(llSdfIndex).lMdDate >= llCffStartDate) And (tlSdfExt(llSdfIndex).lMdDate <= llCffEndDate) Then
                                ilDay = gWeekDayLong(tlSdfExt(llSdfIndex).lMdDate)
                                If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                    If (tgCffCB(ilCff).CffRec.iDay(ilDay) <> 0) Or (tgCffCB(ilCff).CffRec.sXDay(ilDay) = "Y") Then
                                        ilDateFound = True
                                    End If
                                Else
                                    If (tgCffCB(ilCff).CffRec.iDay(ilDay) <> 0) Then
                                        ilDateFound = True
                                    End If
                                End If
                                Exit Do
                            End If
                            ilCff = tgCffCB(ilCff).iNextCff
                        Loop
                        If Not ilDateFound Then
                            'illegal Date
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
 'Doug****************
                                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
                                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    If tmSdf.sTracer = "*" Then
                                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
                                        ilDateFound = True
                                    End If
                                End If
                                'mCntrSchdSpotChk = False
                                'tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
 'Doug^^^^^^^^^^^^^^^^
                            End If
                        End If
                        mGetLegalTimes slOrigMissedDate, tmSmf.iGameNo, llStartTime(), llEndTime()
                        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slSdfTime
                        llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))

                        ilFound = False
                        For ilTime = 0 To 6 Step 1
                            If (llStartTime(ilTime) >= 0 And llEndTime(ilTime) > 0) Then
                                'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
                                If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilTime
                        If (tlSdfExt(llSdfIndex).sSpotType <> "X") And (Not ilFound) Then
                            mCntrSchdSpotChk = False
                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H2
                        End If
                    End If
                Else
'Doug**********
                    ilFlag = False  'Set a flag variable
                    If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        If tmSdf.sTracer = "*" Then
                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
                            ilFlag = True
                        End If
                    End If
                    If Not ilFlag Then
                        If ilDispOnly Then
                            If (tlSdfExt(llSdfIndex).sSchStatus <> "G") Then
                                If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                    mCntrSchdSpotChk = False
                                    tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1  'Date invalid
                                End If
                            End If
                        Else
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                mCntrSchdSpotChk = False
                                tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1  'Date invalid
                            End If
                        End If
                    End If
 'Doug^^^^^^^^^^^
                End If
                tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
            End If
        End If
    Next llSdfIndex
    'Remove missed date counted flag
    For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(llSdfIndex).lMdDate < 0 Then
            tlSdfExt(llSdfIndex).lMdDate = -tlSdfExt(llSdfIndex).lMdDate    'Used negative to indicate missed counted
        End If
    Next llSdfIndex
    Exit Function
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetLegalTimes                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine range if legal times  *
'*                     for a contract line             *
'*                     current time- see mGetLegalTimes*
'*                     in CntrSchd.Bas                 *
'*                                                     *
'*******************************************************
Sub mGetLegalTimes(slDate As String, ilGameNo As Integer, llStartTime() As Long, llEndTime() As Long)
    Dim ilSsfInMem As Integer
    Dim ilRet As Integer
    Dim llRecPos As Long
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilRPRet As Integer
    Dim ilTimeIndex As Integer
    Dim slTime As String
    Dim illoop As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTimes As Integer
    Dim ilTBTime As Integer
    Dim llTBStartTime As Long
    Dim llTBEndTime As Long

    ilTBTime = 0
    For ilTimes = 0 To 6 Step 1
        llStartTime(ilTimes) = 0
        llEndTime(ilTimes) = 0
    Next ilTimes
    llDate = gDateValue(slDate)
    ilDay = gWeekDayLong(llDate)
    gPackDate slDate, ilDate0, ilDate1
    If (tmRdf.iLtfCode(0) <> 0) Or (tmRdf.iLtfCode(1) <> 0) Or (tmRdf.iLtfCode(2) <> 0) Then
        'Read Ssf for date- test for library
        ilSsfInMem = False
        If (lgSsfDate(ilDay) = llDate) Then
            ''If (tgSsf(ilDay).sType = "O") And (tgSsf(ilDay).iVefcode = tmClf.iVefcode) And (tgSsf(ilDay).iStartTime(0) = 0) And (tgSsf(ilDay).iStartTime(1) = 0) Then
            'If (tgSsf(ilDay).iType = 0) And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iStartTime(0) = 0) And (tgSsf(ilDay).iStartTime(1) = 0) Then
            If (tgSsf(ilDay).iType = ilGameNo) And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iStartTime(0) = 0) And (tgSsf(ilDay).iStartTime(1) = 0) Then
                ilSsfInMem = True
                ilRet = BTRV_ERR_NONE
                llRecPos = lgSsfRecPos(ilDay)
            End If
        End If
        If Not ilSsfInMem Then
            imSsfRecLen = Len(tgSsf(ilDay)) 'Max size of variable length record
            'tgSsfSrchKey.sType = "O" 'slType
            tgSsfSrchKey.iType = ilGameNo   '0 'slType
            tgSsfSrchKey.iVefCode = tmClf.iVefCode
            tgSsfSrchKey.iDate(0) = ilDate0
            tgSsfSrchKey.iDate(1) = ilDate1
            tgSsfSrchKey.iStartTime(0) = 0
            tgSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tgSsf(ilDay), imSsfRecLen, tgSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            ilRPRet = gSSFGetPosition(hmSsf, llRecPos)
        End If
        ''Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilDay).sType = "O") And (tgSsf(ilDay).iVefcode = tmClf.iVefcode) And (tgSsf(ilDay).iDate(0) = ilDate0) And (tgSsf(ilDay).iDate(1) = ilDate1)
        'Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilDay).iType = 0) And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iDate(0) = ilDate0) And (tgSsf(ilDay).iDate(1) = ilDate1)
        Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilDay).iType = ilGameNo) And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iDate(0) = ilDate0) And (tgSsf(ilDay).iDate(1) = ilDate1)
            lgSsfDate(ilDay) = llDate
            lgSsfRecPos(ilDay) = llRecPos
            For illoop = 1 To tgSsf(ilDay).iCount Step 1
               LSet tmProg = tgSsf(ilDay).tPas(ADJSSFPASBZ + illoop)
                If tmProg.iRecType = 1 Then 'Program subrecord
                    If (tmProg.iLtfCode = tmRdf.iLtfCode(0)) Or (tmProg.iLtfCode = tmRdf.iLtfCode(1)) Or (tmProg.iLtfCode = tmRdf.iLtfCode(1)) Then
                        gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                        llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
                        gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "1", slTime
                        llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
                        Exit Do
                    End If
                End If
            Next illoop
            'If (tgSsf(ilDay).iNextTime(0) = 1) And (tgSsf(ilDay).iNextTime(1) = 0) Then
                Exit Do
            'Else
            '    imSsfRecLen = Len(tgSsf(ilDay)) 'Max size of variable length record
            '    ilRet = gSSFGetNext(hmSsf, tgSsf(ilDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            '    ilRPRet = gSSFGetPosition(hmSsf, llRecPos)
            'End If
        Loop
    Else    'Time buy- check if override times defined (if so, use them as bump times)
        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
            For ilTimeIndex = UBound(tmRdf.iStartTime, 2) To LBound(tmRdf.iStartTime, 2) Step -1
                If (tmRdf.iStartTime(0, ilTimeIndex) <> 1) Or (tmRdf.iStartTime(1, ilTimeIndex) <> 0) Then
                    'gUnpackTime tmRdf.iStartTime(0, ilTimeIndex), tmRdf.iStartTime(1, ilTimeIndex), "A", "1", slTime
                    'llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
                    'gUnpackTime tmRdf.iEndTime(0, ilTimeIndex), tmRdf.iEndTime(1, ilTimeIndex), "A", "1", slTime
                    'llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
                    'ilTBTime = ilTBTime + 1
                    gUnpackTimeLong tmRdf.iStartTime(0, ilTimeIndex), tmRdf.iStartTime(1, ilTimeIndex), False, llTBStartTime
                    gUnpackTimeLong tmRdf.iEndTime(0, ilTimeIndex), tmRdf.iEndTime(1, ilTimeIndex), True, llTBEndTime
                    If ilTBTime <= UBound(llStartTime) Then
                        If llTBStartTime <= llTBEndTime Then
                            llStartTime(ilTBTime) = llTBStartTime
                            llEndTime(ilTBTime) = llTBEndTime
                            ilTBTime = ilTBTime + 1
                        Else
                            llStartTime(ilTBTime) = llTBStartTime
                            llEndTime(ilTBTime) = 86400
                            ilTBTime = ilTBTime + 1
                            If ilTBTime <= UBound(llStartTime) Then
                                llStartTime(ilTBTime) = 0
                                llEndTime(ilTBTime) = llTBEndTime
                                ilTBTime = ilTBTime + 1
                            End If
                        End If
                    End If
                End If
            Next ilTimeIndex
        Else
            'gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slTime
            'llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
            'gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slTime
            'llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llTBStartTime
            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llTBEndTime
            If llTBStartTime <= llTBEndTime Then
                llStartTime(ilTBTime) = llTBStartTime
                llEndTime(ilTBTime) = llTBEndTime
            Else
                llStartTime(ilTBTime) = llTBStartTime
                llEndTime(ilTBTime) = 86400
                ilTBTime = ilTBTime + 1
                llStartTime(ilTBTime) = 0
                llEndTime(ilTBTime) = llTBEndTime
            End If
        End If
    End If
End Sub

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
    'ilSSFType = 0
    ilSSFType = tlSdf.iGameNo
    ilSpotSeqNo = 0
    If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        'If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefcode <> tlSdf.iVefcode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
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
        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefcode = tlSdf.iVefcode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
            ilEvtIndex = 1
            Do
                If ilEvtIndex > tmSsf.iCount Then
                    imSsfRecLen = Len(tmSsf)
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefcode = tlSdf.iVefcode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
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
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Sub mObtainCopy(slProduct As String, slZone As String, slCart As String, slISCI As String)
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           slProduct(O)- Product (different zones separated by Chr(10)
'                         first product obtained from tmChf if time zone
'           slZone(O)- Zones
'           slCart(O)- Carts (different zones separated by Chr(10))
'           slISCI(O)- ISCI (different zones separated by Chr(10))
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
        slProduct = Trim$(tmChf.sProduct)
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

'*****************************************************************************
'*
'*      Procedure Name:mObtainMissedForMG
'*
'*             Created:10/09/93      By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:Obtain the missed Sdf records
'*                     to be reported for Order only
'*        3/12/98 Add time filter
'*        4/14/98 Eliminate subscript out of range when
'*                running by As Ordered for advt subtots
'*                (time testing was taking from tmPlSdf
'*                and not the SDF record
'*        6/11/01 D.S. Added new type SmfInfo. Was not finding
'*                all of the make goods due to two different
'*                search keys being used during an extended
'                 read operation.
'*        9/13/04 When site set as "S" (show ordered, updated ordered)
'*                option to produced As Ordered was not properly generated
'        6-25-06 test day selectivity for as ordered (mg/outsides ) on Spot Sales
'*****************************************************************************
Sub mObtainMissedForMG(ilSortType As Integer, ilVefCode As Integer, slStartDate As String, slEndDate As String, llStartTime As Long, llEndTime As Long, ilByOrderOrAir As Integer, ilCostType As Integer, llContrCode As Long)
'   Where:
'       ilSortType(I)- 0- For SalesBy Advt; 1=Sales by Vehicle (tmVef must contain vehicle name)
'       ilVefCode(I)- Vehicle Code
'       slStartDate(I)- Start Date
'       slEndDate(I)- End Date
'       ilByOrderOrAir(I)- 0=Order; 1=Aired, 2 = as aired, pkg ordered (4-6-99)
'       ilCostType(I) - bit map of spot costs to include (n/c, adu, bonus, etc)
'       llContrCode(I) - selective contract (chfcode)
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llTime As Long
    Dim slMktRank As String
    'Dim ilUpper As Integer
    Dim llUpper As Long             '12-5-17 prevent overflow error
    Dim ilOk As Integer
    Dim slStr As String
    Dim slPrice As String
    Dim ilSpotSeqNo As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE
    'Dim ilLoop As Integer
    Dim llLoop As Long
    Dim llUpperSDF As Long

    If ilByOrderOrAir = 1 Then  '9-5-13 need to see if mg was in different month for as ordered run    Or ilByOrderOrAir = 2 Then             'only process for As Ordered
        Exit Sub
    End If
    ReDim tmSmfInfo(0 To 0) As SMFINFO

    If ilSortType = 0 Then
        llUpper = UBound(tmSpotSOF)
    Else
        llUpperSDF = UBound(tmPLSdf)
    End If
    btrExtClear hmSmf   'Clear any previous extend operation
    ilExtLen = Len(tmSmf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSmf   'Clear any previous extend operation
    'tmSmfSrchKey.lChfCode = 0
    'tmSmfSrchKey.iLineNo = 0
    tmSmfSrchKey5.iOrigSchVef = ilVefCode
    'gPackDate slStartDate, tmSmfSrchKey.iMissedDate(0), tmSmfSrchKey.iMissedDate(1)
    gPackDate slStartDate, tmSmfSrchKey5.iMissedDate(0), tmSmfSrchKey5.iMissedDate(1)       '4-5-10
    'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE)   '4-5-10 speed up spot sales by using different key
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSmf, llNoRec, -1, "UC", "SMF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        'ilOffset = gFieldOffset("Smf", "SmfOrigSchVEF")
        'ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        ilOffSet = gFieldOffset("smf", "SmfOrigSchVef")
        ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
            ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmSmf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                'Get Sdf
                tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilOk = True

                    'If tgSpf.sInvAirOrder <> "O" And tgSpf.sInvAirOrder <> "S" Then    '9-13-04 Test Airing Vehicle, 5-7-04 previously testing equal to "O"
                    '8-19-10 Determine spots based on the user selection, not site
                    If ilByOrderOrAir <> 0 And ilByOrderOrAir <> "2" Then    '9-13-04 Test Airing Vehicle, 5-7-04 previously testing equal to "O"
                        If tmSdf.iVefCode <> ilVefCode Then
                            ilOk = False
                        End If
                    Else    'Test Order Vehicle (tgspf.sinvairorder = S or O)
                        If (tmSmf.iOrigSchVef <> ilVefCode) Or tmSdf.sSpotType = "X" Then       '4-9-10 fills will follow scheduled vehicle, not orig vehicle
                            ilOk = False
                        End If
                    End If
                    '8-12-10 test for single contract selection
                    If ((llContrCode > 0) And (llContrCode <> tmSdf.lChfCode)) Then
                        ilOk = False            'not matching contr
                    End If
                Else
                    ilOk = False
                End If
                'see if date of missed side should be excluded
                gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                ilRet = gWeekDayStr(slDate)
                If RptSelCb!ckcSelC8(ilRet) = vbUnchecked Then
                    ilOk = False
                End If
                If ilOk Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then          'always ignore psas & promos
                            ilOk = False
                    End If

                    If tmChf.sType = "C" And Not (RptSelCb!ckcSelC6(0).Value = vbChecked) Then  'include std cntrs?
                        ilOk = False
                    End If
                    If tmChf.sType = "V" And Not (RptSelCb!ckcSelC6(1).Value = vbChecked) Then   'include reserves?
                        ilOk = False
                    End If
                    If tmChf.sType = "T" And Not (RptSelCb!ckcSelC6(2).Value = vbChecked) Then   'include remnants?
                        ilOk = False
                    End If
                    If tmChf.sType = "R" And Not (RptSelCb!ckcSelC6(3).Value = vbChecked) Then   'direct response?
                        ilOk = False
                    End If
                    If tmChf.sType = "Q" And Not (RptSelCb!ckcSelC6(4).Value = vbChecked) Then   'per inquiry?
                        ilOk = False
                    End If

                End If
                If ilOk Then
                    tmSmfInfo(UBound(tmSmfInfo)).tSmf = tmSmf
                    tmSmfInfo(UBound(tmSmfInfo)).iAdfCode = tmChf.iAdfCode
                    tmSmfInfo(UBound(tmSmfInfo)).iSlfCode = tmChf.iSlfCode(0)
                    tmSmfInfo(UBound(tmSmfInfo)).lChfCode = tmChf.lCode
                    tmSmfInfo(UBound(tmSmfInfo)).lCntrNo = tmChf.lCntrNo
                    ReDim Preserve tmSmfInfo(0 To UBound(tmSmfInfo) + 1) As SMFINFO
                End If
                ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Loop
            Loop
            For llLoop = 0 To UBound(tmSmfInfo) - 1 Step 1
                ilOk = True
                tmSmf = tmSmfInfo(llLoop).tSmf
                tmChf.iAdfCode = tmSmfInfo(llLoop).iAdfCode
                tmChf.iSlfCode(0) = tmSmfInfo(llLoop).iSlfCode
                tmChf.lCode = tmSmfInfo(llLoop).lChfCode
                tmChf.lCntrNo = tmSmfInfo(llLoop).lCntrNo
                tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                'get line first, to send to filter routine
                tmClfSrchKey.lChfCode = tmSdf.lChfCode
                tmClfSrchKey.iLine = tmSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                    '4/6/99 OK if reporting by ordered only, or if reporting by aired/packages and its a hidden line to package, its OK
                    If (ilByOrderOrAir = 0) Or (ilByOrderOrAir = 2 And tmClf.sType = "H") Then
                        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                        tmSdf.iVefCode = tmClf.iVefCode
                        If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                            mTestCostType ilOk, ilCostType, slPrice, tmSdf.sSchStatus   '10-18-10 test mg/out
                        End If
                    Else
                        ilOk = False
                    End If
                Else
                    ilOk = False
                End If

                If ilOk Then
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        'gUnpackTimeLong tmPLSdf(ilUpper).tSdf.iTime(0), tmPLSdf(ilUpper).tSdf.iTime(1), False, lltime
                    If llTime < llStartTime Or llTime >= llEndTime Then
                        ilOk = False
                    End If
                End If
                If ilOk Then

                    tmSdf.sSchStatus = "M"
                    'If tgSpf.sInvAirOrder <> "O" Then    'Test Airing Vehicle
                    '8-19-10 Determine spots based on the user selection, not site
                    If ilByOrderOrAir <> 0 Then    'Test Airing Vehicle
                        tmSdf.iVefCode = tmSmf.iOrigSchVef
                    End If
                    tmSdf.iDate(0) = tmSmf.iMissedDate(0)
                    tmSdf.iDate(1) = tmSmf.iMissedDate(1)
                    tmSdf.iTime(0) = tmSmf.iMissedTime(0)
                    tmSdf.iTime(1) = tmSmf.iMissedTime(1)
                    If ilSortType = 0 Then
                        'Build Key
                        tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If tmSlf.iSofCode <> 0 Then
                            tmSofSrchKey.iCode = tmSlf.iSofCode
                            ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                tmSof.sName = "Missing"
                                tmSof.iMktRank = 9999
                            End If
                        Else
                            tmSof.sName = "Missing"
                            tmSof.iMktRank = 9999
                        End If
                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        tmSpotSOF(llUpper).tSdf = tmSdf
                        slMktRank = Trim$(str$(tmSof.iMktRank))
                        Do While Len(slMktRank) < 4
                            slMktRank = "0" & slMktRank
                        Loop
                        tmSpotSOF(llUpper).sKey = slMktRank & "|" & tmAdf.sName & "|" & Trim$(str$(tmChf.lCntrNo)) & "|" & tmSof.sName
                        ReDim Preserve tmSpotSOF(0 To llUpper + 1) As SPOTTYPESORT
                        llUpper = llUpper + 1
                    Else
                        tmPLSdf(llUpperSDF).tSdf = tmSdf
                        ilSpotSeqNo = mGetSeqNo(tmPLSdf(llUpperSDF).tSdf)
                        tmPLSdf(llUpperSDF).sKey = tmVef.sName
                        gUnpackDateForSort tmPLSdf(llUpperSDF).tSdf.iDate(0), tmPLSdf(llUpperSDF).tSdf.iDate(1), slDate
                        tmPLSdf(llUpperSDF).sKey = Trim$(tmPLSdf(llUpperSDF).sKey) & "|" & slDate
                        If (tmPLSdf(llUpperSDF).tSdf.sSchStatus = "S") Or (tmPLSdf(llUpperSDF).tSdf.sSchStatus = "G") Or (tmPLSdf(llUpperSDF).tSdf.sSchStatus = "O") Then
                            tmPLSdf(llUpperSDF).sKey = Trim$(tmPLSdf(llUpperSDF).sKey) & "|A"
                        Else
                            tmPLSdf(llUpperSDF).sKey = Trim$(tmPLSdf(llUpperSDF).sKey) & "|Z"
                        End If
                        gUnpackTimeLong tmPLSdf(llUpperSDF).tSdf.iTime(0), tmPLSdf(llUpperSDF).tSdf.iTime(1), False, llTime
                        slStr = Trim$(str$(llTime))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop
                        If ilSpotSeqNo < 10 Then
                            slStr = slStr & "0" & Trim$(str$(ilSpotSeqNo))
                        Else
                            slStr = slStr & Trim$(str$(ilSpotSeqNo))
                        End If
                        tmPLSdf(llUpperSDF).sKey = Trim$(tmPLSdf(llUpperSDF).sKey) & "|" & slStr
                        ReDim Preserve tmPLSdf(0 To llUpperSDF + 1) As SPOTTYPESORT
                        llUpperSDF = llUpperSDF + 1
                    End If
                    'End If
                End If
            Next llLoop
        End If
    End If
    Erase tmSmfInfo
    Exit Sub
    ilRet = err.Number
    Resume Next
End Sub

'******************************************************************
'*                                                                *
'*      Procedure Name:mObtainSdf                                 *
'*                                                                *
'*             Created:10/09/93      By:D. LeVine                 *
'*            Modified: 11/20/96     By:d.h.                      *
'*                                                                *
'*            Comments:Obtain the Sdf records to be               *
'*                     reported                                   *
'*                                                                *
'           3/12/98  Add time filter                              *
'           4/6/99 For packages, test to bill as ordered          *
'                  when billing as aired                          *
'           6/18/99 Included missed when requested for            *
'              option as aired/pkg ordered for NONE.              *
'              previously, package missed not included.           *
'           7/2/99 more problems as described on 6/18/99.         *
'               NONe didnt work properly, advt option  OK         *
'           7-19-04 Include/ exclude contract/Network spots       *
'           10-21-04 handle more than 32000+ for a vehicle for
'                   1 year
'           3-3-10 Ensure that the spots printed on Spots by DAte
'           and time are in the same order as Spot Screen.  Create
'           list box with the sdf codes and an index number.  The
'           index number will be used as a subsort below spot time
'           in the crystal report
'******************************************************************
Sub mObtainSdf(ilVefCode As Integer, slStartDate As String, slEndDate As String, llStartTime As Long, llEndTime As Long, ilSpotType As Integer, ilBillType As Integer, ilIncludePSA As Integer, ilMissedType As Integer, ilISCIOnly As Integer, ilCostType As Integer, ilByOrderOrAir As Integer, ilIncludeType As Integer, llContrCode As Long, ilLocal As Integer, ilFeed As Integer, ilPropPrice As Integer, ilListIndex As Integer, tlCntTypes As CNTTYPES)
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
'       ilByOrderOrAir(I)- 0=Order; 1=Aired   , 3 = as aired, pkg ordered (option 3 added 4-6-99)
'       ilIncludeType - true to test contract type inclusions, else false to ignore test
'       llContrCode - if selective contract, code # (else 0 for all)
'       ilLocal - true to include local spots
'       ilFeed - true to include network (feed) spots
'       ilPropPrice - true to show proposal price vs actual price
'       ilListIndex - report index representing report selected
    Dim slDate As String
    Dim llTime As Long
    Dim ilRet As Integer
    Dim ilReturn As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim llUpper As Long
    Dim ilOk As Integer
    Dim ilSpotSeqNo As Integer
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim tlCff As CFF
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLoopOnDate As Long
    Dim ilDate(0 To 1) As Integer
    Dim ilVefIndex As Integer
    Dim ilType As Integer
    Dim ilEvt As Integer
    Dim ilSpot As Integer
    Dim illoop As Integer
    Dim llUpperSortSeq  As Long
    Dim llMissedDate As Long
    Dim ilGameSelect As Integer
    Dim ilLineSelect As Integer
    Dim llChfSelect As Long
    Dim ilVpfIndex As Integer
    Dim slLLDate As String
    Dim llLLDate As Long
    Dim slSpotDate As String
    
    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    
        'get the SSF records for this vehicle so that the order of the spots within the spot screen can be maintained
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
                                   
    ilVefIndex = gBinarySearchVef(ilVefCode)
    
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

    ilType = 0
    
    '5-9-11 Remove all the invalid bb spots that doesnt belong
    ilGameSelect = 0
    ilLineSelect = 0
    llChfSelect = llContrCode
    ilRet = gRemoveBBSpots(hmSdf, ilVefCode, ilGameSelect, slStartDate, slEndDate, llChfSelect, ilLineSelect)

    'llUpperSortSeq = UBound(tmSeqSortType)
    If ilListIndex = CNT_SPTSBYDATETIME Then        'only spots by date and time need to be in exact order as the spot screen
        llUpperSortSeq = UBound(tmSeqSortType)
        For llLoopOnDate = llStartDate To llEndDate
            slDate = Format$(llLoopOnDate, "m/d/yy")
            gPackDate slDate, ilDate(0), ilDate(1)
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            'tmSsfSrchKey.sType = slType
            If tgMVef(ilVefIndex).sType <> "G" Then
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                tmSsfSrchKey.iDate(0) = ilDate(0)
                tmSsfSrchKey.iDate(1) = ilDate(1)
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Else
                tmSsfSrchKey2.iVefCode = ilVefCode
                tmSsfSrchKey2.iDate(0) = ilDate(0)
                tmSsfSrchKey2.iDate(1) = ilDate(1)
                ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                ilType = tmSsf.iType
            End If
            
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate(0)) And (tmSsf.iDate(1) = ilDate(1)))
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmAvail.iRecType = 2 Then
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        If llTime >= llStartTime And llTime < llEndTime Then      'event within entered time parameters
                            For ilSpot = 1 To tmAvail.iNoSpotsThis
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot + ilEvt)
                                slStr = Trim$(str(tmSpot.lSdfCode))
                                Do While Len(slStr) < 10
                                    slStr = "0" & slStr
                                Loop
                                tmSeqSortType(llUpperSortSeq).sKey = slStr     'internal spot code for sorting
                                tmSeqSortType(llUpperSortSeq).lSdfCode = tmSpot.lSdfCode
                                tmSeqSortType(llUpperSortSeq).iSeqNo = ilSpot
                                ReDim Preserve tmSeqSortType(0 To llUpperSortSeq + 1) As SEQSORTTYPE
                                llUpperSortSeq = llUpperSortSeq + 1
                                'RptSelCb!lbcLnCode.AddItem slStr         'Add ID to list box
                                'RptSelCb!lbcLnCode.ItemData(RptSelCb!lbcLnCode.NewIndex) = ilSpot
                            Next ilSpot
                        End If
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'increment past spots
                    End If
                    ilEvt = ilEvt + 1                                       'increment to next event
                Loop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
        Next llLoopOnDate
    End If
    
    llUpper = UBound(tmPLSdf)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        '7-21-04  Exclude/include contract/feed spots
        tlLongTypeBuff.lCode = 0
        If Not ilLocal Or Not ilFeed Then           'either local or feed spots are to be excluded
            If ilLocal Then                         'include local only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
            ElseIf ilFeed Then                      'include feed only
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
        ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
            Loop
            tmPLSdf(llUpper).sLiveCopy = ""
            Do While ilRet = BTRV_ERR_NONE
                ilOk = mSpotDateFilter(tmPLSdf(llUpper), llStartDate, llEndDate, llStartTime, llEndTime, ilSpotType, ilBillType, ilIncludePSA, ilMissedType, ilISCIOnly, ilCostType, ilByOrderOrAir, ilIncludeType, llContrCode, ilPropPrice, llLLDate, tlCntTypes)
                If ilOk Then
                    If ilListIndex = CNT_HILORATE Then
                                               'key = vehicle name, contract code, line #
                        tmPLSdf(llUpper).sKey = tmVef.sName
                        slStr = Trim$(str$(tmSdf.lChfCode))
                        Do While Len(slStr) < 9
                            slStr = "0" & slStr
                        Loop
                        slStr = slStr & "|" & Trim$(str$(tmSdf.iLineNo))
                        Do While Len(slStr) < 4
                            slStr = "0" & slStr
                        Loop
                                                
                        '6-3-10 gather flight info
                        ilRet = gGetSpotFlight(tmPLSdf(llUpper).tSdf, tmClf, hmCff, hmSmf, tlCff)
                        tmPLSdf(llUpper).sDyWk = tlCff.sDyWk
                        For illoop = 0 To 6
                            tmPLSdf(llUpper).iDay(illoop) = tlCff.iDay(illoop)
                            tmPLSdf(llUpper).sXDay(illoop) = tlCff.sXDay(illoop)
                        Next illoop

                    Else
                        ilSpotSeqNo = mGetSeqNo(tmPLSdf(llUpper).tSdf)
                        tmPLSdf(llUpper).sKey = tmVef.sName
                        gUnpackDateForSort tmPLSdf(llUpper).tSdf.iDate(0), tmPLSdf(llUpper).tSdf.iDate(1), slDate
                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slDate
                        If (tmPLSdf(llUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "O") Then
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|A"
                        Else
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|Z"
                        End If
                        gUnpackTimeLong tmPLSdf(llUpper).tSdf.iTime(0), tmPLSdf(llUpper).tSdf.iTime(1), False, llTime
                        slStr = Trim$(str$(llTime))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop
                        If ilSpotSeqNo < 10 Then
                            slStr = slStr & "0" & Trim$(str$(ilSpotSeqNo))
                        Else
                            slStr = slStr & Trim$(str$(ilSpotSeqNo))
                        End If
                    End If
                    tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slStr
                    ReDim Preserve tmPLSdf(0 To llUpper + 1) As SPOTTYPESORT
                    llUpper = llUpper + 1
                End If
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
    ilRet = err.Number
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSdfBySOF                 *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Sdf records to be    *
'*                     reported
'       3/12/98 Add filter by Time                     *
'       4/6/99 For packages, test to bill as ordered
'                  when billing as aired
'       6/18/99 As aired/pkg ordered:  handle missed
'               properly
'*      10-23-04 handle 32000+ spots                   *
'*******************************************************
Sub mObtainSdfBySOF(ilVefCode As Integer, slStartDate As String, slEndDate As String, llStartTime As Long, llEndTime As Long, ilMissedType As Integer, ilCostType As Integer, ilByOrderOrAir As Integer, llContrCode As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilUpper                                                                               *
'******************************************************************************************
'   Where:
'       ilVefCode(I)- Vehicle Code
'       slStartDate(I)- Start Date
'       slEndDate(I)- End Date
'       ilMissedType(I)-
'       ilCostType(I) - bit map -inclusion of different spot spot types (n/c, adu, bonus, fill, etc)
'       ilByOrderOrAir(I)- 0=Order; 1=Aired   , 2 = as aired, pkg ordered (option 3 added 4-6-99)
'       llContrCode(I) - selective contract (chfCode)
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slMktRank As String
    Dim slPrice As String
    Dim llUpper As Long         '10-23-04
    Dim ilOk As Integer
    Dim llTime As Long                  'Spot time
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim slDate As String
    Dim llMissedDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    ReDim tmSpotSOF(0 To 0) As SPOTTYPESORT
   
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    
    llUpper = LBound(tmSpotSOF)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)

        '7-19-04  Exclude network spots
        tlLongTypeBuff.lCode = 0
        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
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
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                '4/6/99 get line first, to send to filter routine
                tmClfSrchKey.lChfCode = tmSdf.lChfCode
                tmClfSrchKey.iLine = tmSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                '8-12-10 test for single contract selection
               If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((llContrCode = 0) Or (llContrCode > 0 And llContrCode = tmSdf.lChfCode)) Then
                    ilOk = True
                Else
                    ilOk = False
                End If

                'check day of week selectivity
                 gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                ilRet = gWeekDayStr(slDate)
                If RptSelCb!ckcSelC8(ilRet) = vbUnchecked Then
                    ilOk = False
                End If

                '3-30-05 Ignore BB spots
                If tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C" Then
                    ilOk = False
                Else
                    If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                        If ilByOrderOrAir = 0 Then
                            If (tmSdf.sSchStatus = "S") Or (tmSdf.sSpotType = "X") Then     '4-9-10 For as ordered, fill spots follow the sched vehicle, not orig vehicle
                                'ilOk = True
                            Else
                                ilOk = False
                            End If
                        Else
                            'ilOK = True
                            If ilByOrderOrAir = 2 Then          'aired/pkg ordered only
                                If tmClf.sType = "H" And (tmSdf.sSchStatus <> "S") Then
                                    'ilOk = False        'ignore the mgs and outsides here for packages
                                    'need to get the smf to see what if the original missed date was within the requested dates
                                    tmSmfSrchKey2.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    If ilRet = BTRV_ERR_NONE Then
                                        '8-19-10 check if original missed within requested period
                                        gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                                        If (llMissedDate < llStartDate Or llMissedDate > llEndDate) Or (tmClf.iVefCode <> tmSdf.iVefCode) Or ((llMissedDate >= llStartDate And llMissedDate <= llEndDate) And (tmClf.iVefCode = tmSdf.iVefCode)) Then        '8-4-16 if different vehicles, ignore because it will be based on ordered
                                            ilOk = False
                                        End If
                                    Else
                                        ilOk = False
                                    End If
                                End If
                            End If

                        End If
                    Else        'spot missed, cancelled or hidden
                        'If ilByOrderOrAir = 0 Then
                        '    If (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "C") Then
                        '        ilOk = False
                        '    Else
                        '        ilOk = True
                        '    End If
                        'Else
                        If ilByOrderOrAir = 2 Then
                            'alter ilmissedtype based if line is pkg or conventional
                            'show pkg as ordered, show conventional as aired for missed/hidden/cancelled spot
                                'If tmClf.sType = "H" And (tmPLSdf(ilUpper).tSdf.sSchStatus <> "S") Then
                                If tmClf.sType <> "H" Then          'not hidden (for pkg), must be conventional
                                    If (tmSdf.sSchStatus = "H") Then    '6/18/99
                                        If (ilMissedType And &H4) <> &H4 Then
                                            ilOk = False
                                        End If
                                    ElseIf (tmSdf.sSchStatus = "C") Then
                                        If (ilMissedType And &H2) <> &H2 Then
                                            ilOk = False
                                        End If
                                    Else
                                        If (ilMissedType And &H1) <> &H1 Then
                                            ilOk = False
                                        End If
                                    End If
                                End If
                        Else        'by ordered or aired for a missed/hidden/cancelled spot
                        'End If
                            If ilByOrderOrAir = 0 Then          '8-26-03 ordered method was not including the missed spots
                                'ilOk = True                 'ignore whether to exclude/include the missed--its by ordered so include them all
                                '8-16-10 if already false, dont reset to true to include the spot
                            Else                            'aired method, choose which type not aired to included
                                ilOk = False
                                If (tmSdf.sSchStatus = "H") Then
                                    If (ilMissedType And &H4) = &H4 Then
                                        ilOk = True
                                    End If
                                ElseIf (tmSdf.sSchStatus = "C") Then
                                    If (ilMissedType And &H2) = &H2 Then
                                        ilOk = True
                                    End If
                                Else
                                    If (ilMissedType And &H1) = &H1 Then
                                        ilOk = True
                                    End If
                                End If
                            End If
                        End If
                        'End If
                    End If
                End If

                If ilOk Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then          'always ignore psas & promos
                        ilOk = False
                    End If

                    If tmChf.sType = "C" And Not (RptSelCb!ckcSelC6(0).Value = vbChecked) Then  'include std cntrs?
                        ilOk = False
                    End If
                    If tmChf.sType = "V" And Not (RptSelCb!ckcSelC6(1).Value = vbChecked) Then   'include reserves?
                        ilOk = False
                    End If
                    If tmChf.sType = "T" And Not (RptSelCb!ckcSelC6(2).Value = vbChecked) Then   'include remnants?
                        ilOk = False
                    End If
                    If tmChf.sType = "R" And Not (RptSelCb!ckcSelC6(3).Value = vbChecked) Then   'direct response?
                        ilOk = False
                    End If
                    If tmChf.sType = "Q" And Not (RptSelCb!ckcSelC6(4).Value = vbChecked) Then   'per inquiry?
                        ilOk = False
                    End If

                    If (ilOk) Then                          'test time filters
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        If llTime < llStartTime Or llTime >= llEndTime Then
                            ilOk = False
                        End If
                    End If

                    If ilOk Then                'filter out spots
                        'get line first, to send to filter routine
                        'tmClfSrchKey.lchfcode = tmSdf.lchfcode
                        'tmClfSrchKey.iLine = tmSdf.iLineNo
                        'tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        'tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        'ilret = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        'Do While (ilret = BTRV_ERR_NONE) And (tmClf.lchfcode = tmSdf.lchfcode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
                        '    ilret = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        'Loop
                        'If (ilret = BTRV_ERR_NONE) And (tmClf.lchfcode = tmSdf.lchfcode) And (tmClf.iLine = tmSdf.iLineNo) Then
                            ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
'                            tmSdf.iVefCode = tmClf.iVefCode            '5-9-17   retain the spots vehicle code
                            If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                                mTestCostType ilOk, ilCostType, slPrice, tmSdf.sSchStatus   '10-18-10 test mg/out
                                '5-25-05 test for inclusion/exclusion of BB spots
                                If (tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C") And (ilCostType And SPOT_BB) <> SPOT_BB Then
                                    ilOk = False
                                End If
                            End If
                        'Else
                        '    ilOK = False
                        'End If
                    End If
                    If ilOk Then
                        'Build Key
                        tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If tmSlf.iSofCode <> 0 Then
                            tmSofSrchKey.iCode = tmSlf.iSofCode
                            ilRet = btrGetEqual(hmSof, tmSof, imSofRecLen, tmSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                tmSof.sName = "Missing"
                                tmSof.iMktRank = 9999
                            End If
                        Else
                            tmSof.sName = "Missing"
                            tmSof.iMktRank = 9999
                        End If
                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        tmSpotSOF(llUpper).tSdf = tmSdf
                        slMktRank = Trim$(str$(tmSof.iMktRank))
                        Do While Len(slMktRank) < 4
                            slMktRank = "0" & slMktRank
                        Loop
                        tmSpotSOF(llUpper).sKey = slMktRank & "|" & tmAdf.sName & "|" & Trim$(str$(tmChf.lCntrNo)) & "|" & tmSof.sName
                        ReDim Preserve tmSpotSOF(0 To llUpper + 1) As SPOTTYPESORT
                        llUpper = llUpper + 1
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

    ilRet = err.Number
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSelSdf                   *
'           Spots by Advertiser                        *
'           MG Revenue                                 *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Sdf records to be    *
'*                     reported                        *
'*                                                     *
'           7-19-04 Exclude/include local/feed         *
'           4-22-05 change table array tohandle 32000+
'*******************************************************
Sub mObtainSelSdf(ilWhichKey As Integer, ilKeyField As Integer, slStartDate As String, slEndDate As String, slStartEDate As String, slEndEDate As String, ilSelType As Integer, ilCostType As Integer, ilLocal As Integer, ilFeed As Integer, ilIncludeCodes As Integer, ilCodes() As Integer, ilUseAcqRate As Integer, tlCntTypes As CNTTYPES, Optional llTestCntrNo As Long = 0)
'   where:
'       ilWhichKey - INDEXKEY1 (by vehicle) or INDEXKEY7 (by advt)
'       ilKeyfield - if INDEXKEY1 then field is vehicle, otherwise advt code
'       slStartDate(I)- Spot start date
'       slEndDate(I)- Spot end date
'       slStartEDate(I)- Contract entered start date
'       slEndEDate(I)- Contract entered End date
'       ilSelType(I)-0=Advertiser, 1=Agency; 2=Salesperson; 3=No selection
'       tmSelChf(I)- contains the selections
'       tmSelAgf(I)
'       tmSelSlf(I)
'       ilCostType(I) - bit string of spots types to include
'                   bit 0 = charged, 1 = .00, 2 = adu, 3 = bonus, 4 = extra,
'                       5 = fill, 6 = n/c, 7 = mg, 8 = recapturable, 9 = spinoff
'       ilLocal - true to include local spots
'       ilFeed - true to include network (feed) spots
'       ilUseAcqRate = true if using barters and user selected to show rates along with acq.
    Dim slDate As String
    Dim llTime As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    'Dim ilUpper As Integer      '4-22-05 change to long
    Dim llUpper As Long          '4-22-05
    Dim illoop As Integer
    Dim ilOk As Integer
    Dim llStartEDate As Long
    Dim llEndEDate As Long
    Dim slEnteredDate As String         'common date entered field from contract hdr or feed spot record
    Dim llContrNumber As Long           'common contr # field from contr hdr or feed spot record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim ilAgfCode As Integer         '11-25-09 Need to test for direct advertiser selection for AGency sort
    Dim ilGameSelect As Integer
    Dim ilLineSelect As Integer
    Dim llChfSelect As Long
    Dim slLLDate As String
    Dim llLLDate As Long
    Dim slSpotDate As String
    Dim ilVpfIndex As Integer
    Dim ilVefInx As Integer

    If Len(Trim$(slStartEDate)) = 0 Then
        llStartEDate = 0
    Else
        llStartEDate = gDateValue(slStartEDate)
    End If
    If Len(Trim$(slEndEDate)) = 0 Then
        llEndEDate = 0
    Else
        llEndEDate = gDateValue(slEndEDate)
    End If

'   10-23-14  Move test later when theres a spot to get vehicle code .  search may be by advt, so the vehicle code isnt being sent via caller
    
'    ilVpfIndex = gBinarySearchVpf(ilVefCode)
'
'    'ignore  BB from days in future
'    If ilVpfIndex <> -1 Then
'        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slLLDate
'        If slLLDate = "" Then
'            slLLDate = Format(Now, "m/d/yy")
'        Else
'            If gDateValue(slLLDate) < gDateValue(Format(Now, "m/d/yy")) Then
'                slLLDate = Format(Now, "m/d/yy")
'            End If
'        End If
'        slLLDate = gIncOneDay(slLLDate)
'    Else
'        slLLDate = gIncOneDay(Format(Now, "m/d/yy"))
'    End If
    llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater
    
'    10-23-14 clear the billboards outside of this routine
'    '5-9-11 Remove all the invalid bb spots that doesnt belong
'    ilGameSelect = 0
'    ilLineSelect = 0
'    llChfSelect = 0
'    ilRet = gRemoveBBSpots(hmSdf, ilVefCode, ilGameSelect, slStartDate, slEndDate, llChfSelect, ilLineSelect)

    llUpper = UBound(tmPLSdf)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    If ilWhichKey = INDEXKEY1 Then
        tmSdfSrchKey1.iVefCode = ilKeyField
        gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = ""
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Else
        tmSdfSrchKey7.iAdfCode = ilKeyField
        gPackDate slStartDate, tmSdfSrchKey7.iDate(0), tmSdfSrchKey7.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey7, INDEXKEY7, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)

        '7-19-04  Exclude network spots
        tlLongTypeBuff.lCode = 0
        If Not ilLocal Or Not ilFeed Then           'either local or feed spots are to be excluded
            If ilLocal Then                         'include local only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
            ElseIf ilFeed Then                      'include feed only
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
            End If
        End If
        tlIntTypeBuff.iType = ilKeyField
        If ilWhichKey = INDEXKEY1 Then
            ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        Else
            ilOffSet = gFieldOffset("Sdf", "SdfAdfCode")
        End If
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
        ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'Build sort key
                ilOk = False
                '12-28-17 access the contract upfront, to do contract type selection
                tmChfSrchKey.lCode = tmPLSdf(llUpper).tSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation

                If ilSelType = 0 Or (ilWhichKey = INDEXKEY7 And ilSelType = 3) Then   'Advertiser, match the selected contracts built for the advts, or selective 1 advt
                'if selective 1 advt, need to check for vehicle selectivity
                    If ilWhichKey = INDEXKEY7 Then
                        For illoop = 0 To UBound(tmSelVef) - 1 Step 1
                            If tmSelVef(illoop) = tmPLSdf(llUpper).tSdf.iVefCode Then
                                ilOk = True
                                Exit For
                            End If
                        Next illoop
                    Else
                        If tmPLSdf(llUpper).tSdf.lChfCode = 0 Then       'network feed, find all net spots matching selected advt
                            'match the advt codes
                            For illoop = 0 To UBound(tmSelAdf) - 1 Step 1
                                If tmSelAdf(illoop) = tmPLSdf(llUpper).tSdf.iAdfCode Then
                                    ilOk = True
                                    Exit For
                                End If
                            Next illoop
                        Else
                            For illoop = 0 To UBound(tmSelChf) - 1 Step 1
                                If tmSelChf(illoop) = tmPLSdf(llUpper).tSdf.lChfCode Then
                                    ilOk = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                    End If
                ElseIf ilSelType = 1 Then   'Agency
                        '12-28-17 moving reading of contract header so all options can use the contract header to test contract types
'                        tmChfSrchKey.lCode = tmPLSdf(llUpper).tSdf.lChfCode
'                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        
'                       11-25-09 need to test for direct advertisers, see below
'                        For ilLoop = 0 To UBound(tmSelAgf) - 1 Step 1
'                            If tmSelAgf(ilLoop) = tmChf.iAgfCode Then
'                                ilOk = True
'                                Exit For
'                            End If
'                        Next ilLoop
                        
                        If tmChf.iAgfCode = 0 Then
                            ilAgfCode = -tmChf.iAdfCode
                        Else
                            ilAgfCode = tmChf.iAgfCode
                        End If
                        If tmChf.iAdfCode = 740 Then
                            ilRet = ilRet
                        End If
                        ilOk = gFilterAgyAdvCodes(ilAgfCode, ilIncludeCodes, ilCodes())

                ElseIf ilSelType = 2 Then    'Salesperson
                    tmChfSrchKey.lCode = tmPLSdf(llUpper).tSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    For illoop = 0 To UBound(tmSelSlf) - 1 Step 1
                        If tmSelSlf(illoop) = tmChf.iSlfCode(0) Then
                            ilOk = True
                            Exit For
                        End If
                    Next illoop
                Else
                    ilOk = True
                End If
                
                '12-28-17 filter out the contract types
                mFilterCntTypes tmChf, tlCntTypes, ilOk
                If llTestCntrNo <> 0 Then
                    If tmChf.lCntrNo <> llTestCntrNo Then ilOk = False
                End If
                    
                If ilOk Then            'test for spot cost inclusion
                    ilVefInx = gBinarySearchVef(tmPLSdf(llUpper).tSdf.iVefCode)
                    tmVef = tgMVef(ilVefInx)
                    tmPLSdf(llUpper).sLiveCopy = ""

                    If tmPLSdf(llUpper).tSdf.lChfCode > 0 Then      'network feed spot (vs contract spot) has no lines to access
                        'get line first, to send to filter routine
                        tmClfSrchKey.lChfCode = tmPLSdf(llUpper).tSdf.lChfCode
                        tmClfSrchKey.iLine = tmPLSdf(llUpper).tSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(llUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(llUpper).tSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(llUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(llUpper).tSdf.iLineNo) Then
                            If ilUseAcqRate Then
                                tmPLSdf(llUpper).sCostType = gLongToStrDec(tmClf.lAcquisitionCost, 2)
                            Else
                                ilRet = gGetSpotPrice(tmPLSdf(llUpper).tSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, tmPLSdf(llUpper).sCostType)
                            End If
                            tmPLSdf(llUpper).iVefCode = tmClf.iVefCode
                            tmPLSdf(llUpper).sLiveCopy = tmClf.sLiveCopy         '9-2-15
                            mTestCostType ilOk, ilCostType, tmPLSdf(llUpper).sCostType, tmPLSdf(llUpper).tSdf.sSchStatus    '10-18-10 test mg/out
                            '5-25-05 test for inclusion/exclusion of BB spots
                            If (tmPLSdf(llUpper).tSdf.sSpotType = "O" Or tmPLSdf(llUpper).tSdf.sSpotType = "C") And (ilCostType And SPOT_BB) <> SPOT_BB Then
                                ilOk = False
                            End If
                            'If Not ilOk Then
                                'ilOk = False
                            'End If
                        Else
                            ilOk = False
                        End If
                         'Test if Open or Close BB, ignore if in the future
                        If tmPLSdf(llUpper).tSdf.sSpotType = "O" Or tmPLSdf(llUpper).tSdf.sSpotType = "C" Then
                        
                            ilVpfIndex = gBinarySearchVpf(tmPLSdf(llUpper).tSdf.iVefCode)
    
                            'ignore  BB from days in future
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
                        
                            gUnpackDate tmPLSdf(llUpper).tSdf.iDate(0), tmPLSdf(llUpper).tSdf.iDate(1), slSpotDate
                            If gDateValue(slSpotDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
                                ilOk = False
                            End If
                        End If
                    Else                        'network feed spot
                        tmPLSdf(llUpper).sCostType = "Feed"
                    End If
                End If
                If ilOk Then
                    If (ilSelType = 0) Or (ilSelType = 3) Then   'Advertiser
                        If tmPLSdf(llUpper).tSdf.lChfCode > 0 Then
                            tmChfSrchKey.lCode = tmPLSdf(llUpper).tSdf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slEnteredDate
                            llContrNumber = tmChf.lCntrNo
                        Else
                            tmChfSrchKey.lCode = tmPLSdf(llUpper).tSdf.lFsfCode
                            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            gUnpackDate tmFsf.iEnterDate(0), tmFsf.iEnterDate(1), slEnteredDate
                            llContrNumber = 0
                        End If
                    End If
                    If llStartEDate <> 0 Then
                        'gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slDate
                        If gDateValue(slEnteredDate) < llStartEDate Then
                            ilOk = False
                        End If
                    End If
                    If llEndEDate <> 0 Then
                        'gUnpackDate tmChf.iPropDate(0), tmChf.iPropDate(1), slDate
                        If gDateValue(slEnteredDate) > llEndEDate Then
                            ilOk = False
                        End If
                    End If
                    If ilOk Then
                        If tmAdf.iCode <> tmPLSdf(llUpper).tSdf.iAdfCode Then
                            tmAdfSrchKey.iCode = tmPLSdf(llUpper).tSdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        tmPLSdf(llUpper).sKey = tmAdf.sName
                        slStr = Trim$(str$(llContrNumber))       'Trim$(Str$(tmChf.lCntrNo))
                        Do While Len(slStr) < 8
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slStr
                        slStr = Trim$(str$(tmPLSdf(llUpper).tSdf.iLineNo))
                        Do While Len(slStr) < 4
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slStr


                        If RptSelCb!rbcSelC4(1).Value Then      'sort by date, then station
                            'Date
                            gUnpackDateForSort tmPLSdf(llUpper).tSdf.iDate(0), tmPLSdf(llUpper).tSdf.iDate(1), slDate
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slDate

                            'Station
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & tmVef.sName
                        Else                            'sort by station, then date
                            'vehicle
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & tmVef.sName
                            'Date
                            gUnpackDateForSort tmPLSdf(llUpper).tSdf.iDate(0), tmPLSdf(llUpper).tSdf.iDate(1), slDate
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slDate
                        End If
                        If (tmPLSdf(llUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(llUpper).tSdf.sSchStatus = "O") Then
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|A"
                        Else
                            tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|Z"
                        End If
                        gUnpackTimeLong tmPLSdf(llUpper).tSdf.iTime(0), tmPLSdf(llUpper).tSdf.iTime(1), False, llTime
                        slStr = Trim$(str$(llTime))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(llUpper).sKey = Trim$(tmPLSdf(llUpper).sKey) & "|" & slStr
                        ReDim Preserve tmPLSdf(0 To llUpper + 1) As SPOTTYPESORT
                        llUpper = llUpper + 1
                    End If
                End If
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmPLSdf(llUpper).tSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub

    ilRet = err.Number
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSelPCF                   *
'*          Digital Lines by Advertiser / Vehicle      *
'*                                                     *
'*             Created:6/14/23       By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the PCF records to be    *
'*                     reported                        *
'*                                                     *
'*******************************************************
'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
Sub mObtainSelPcf(ilWhichKey, llItemCode As Long, slStartDate As String, slEndDate As String, tmSelAgf() As Integer, tmSelAdf() As Integer, tmSelChf() As Long, tlCntTypes As CNTTYPES, Optional llCntrNo As Long = 0)
'   where:
'       ilWhichKey      - 0=Contracts, 1=Agency, 5=Advertiser, 6=Vehicles
'       llItemCode      - the value of the Contr,Agency,Adv or Vehicle (Depending on ilWhichKey)
'       slStartDate     - Digital Line start date
'       slEndDate       - Digital Line end date
'       slCntrStartDate - Contract entered start date
'       slCntrEndDate   - Contract entered End date
'       ilSelType       - 0=Advertiser, 1=Agency; 2=Salesperson; 3=No selection
'       tmSelChf        - contains the selections
    Dim blValid As Boolean
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    Dim illoop As Long
    Dim llUpper As Long
    Dim blNeedComma As Boolean
    
    slSQLQuery = ""
    slSQLQuery = slSQLQuery & "SELECT "
    slSQLQuery = slSQLQuery & "  chf.chfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfBillCycle,"
    slSQLQuery = slSQLQuery & "  chf.chfCntrNo,"
    slSQLQuery = slSQLQuery & "  chf.chfExtCntrNo," 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 5)
    slSQLQuery = slSQLQuery & "  chf.chfType,"
    slSQLQuery = slSQLQuery & "  chf.chfStartDate,"
    slSQLQuery = slSQLQuery & "  chf.chfEndDate,"
    slSQLQuery = slSQLQuery & "  chf.chfAgfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfAdfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfProduct,"
    
    slSQLQuery = slSQLQuery & "  pcf.pcfCode,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPodCPMID,"
    slSQLQuery = slSQLQuery & "  pcf.pcfVefCode,"
    'slSQLQuery = slSQLQuery & "  pcf.pcfType,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPriceType,"
    slSQLQuery = slSQLQuery & "  pcf.pcfStartDate,"
    slSQLQuery = slSQLQuery & "  pcf.pcfEndDate,"
    'slSQLQuery = slSQLQuery & "  pcf.pcfImpressionGoal,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPodCPM,"
    slSQLQuery = slSQLQuery & "  pcf.pcfLen," 'Boostr Phase 2: Spot and Digital Line Combo report: show length for digital lines
    
    slSQLQuery = slSQLQuery & "  pcf.pcfRdfCode," 'Daypart
    slSQLQuery = slSQLQuery & "  pcf.pcfCxfCode," 'Line Comment
    
    slSQLQuery = slSQLQuery & "  pcf.pcfTotalCost"
    slSQLQuery = slSQLQuery & " FROM "
    slSQLQuery = slSQLQuery & "  CHF_Contract_Header chf"
    slSQLQuery = slSQLQuery & "  JOIN pcf_Pod_CPM_Cntr pcf on pcf.pcfChfCode = chf.chfCode"
    slSQLQuery = slSQLQuery & " WHERE "
    
    'TTP 10961 - Spot and Digital Line combo report
    'Contract Status
    slSQLQuery = slSQLQuery & "  chf.chfStatus in ("
    blNeedComma = False
    'Chf.sStatus = "H" tlCntTypes.iHold
    If tlCntTypes.iHold Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'H'" 'Holds
        blNeedComma = True
    End If
    'Chf.sStatus = "O" tlCntTypes.iOrder
    If tlCntTypes.iOrder Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'O'" 'Orders (Scheduled Order)
        blNeedComma = True
    End If
    'Chf.sType = "C" tlCntTypes.iStandard
    If tlCntTypes.iStandard Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'C'" 'include Standard Orders (Scheduled Order)
        blNeedComma = True
    End If
    'Chf.sType = "V" tlCntTypes.iReserv
    If tlCntTypes.iReserv Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'V'" 'include reservations ?
        blNeedComma = True
    End If
    'Chf.sType = "R" tlCntTypes.iDR
    If tlCntTypes.iDR Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'R'" 'include DR?
        blNeedComma = True
    End If
    'Chf.sType = "S"  tlCntTypes.iPSA
    If tlCntTypes.iPSA Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'S'" 'include PSA ?
        blNeedComma = True
    End If
    'Chf.sType = "M"  tlCntTypes.iPromo
    If tlCntTypes.iPromo Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'M'" 'include Promo?
        blNeedComma = True
    End If
    'Chf.sType = "T"  tlCntTypes.iRemnant
    If tlCntTypes.iRemnant Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'T'" 'include Remnant?
        blNeedComma = True
    End If
    'Chf.sType = "Q" tlCntTypes.iPI
    If tlCntTypes.iPI Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'Q'" 'include PI?
        blNeedComma = True
    End If
    slSQLQuery = slSQLQuery & ")"
    
    'TTP 10961 - Spot and Digital Line combo report
    If llCntrNo <> 0 Then
        slSQLQuery = slSQLQuery & "  AND chf.chfCntrNo = " & llCntrNo
    End If
    
    slSQLQuery = slSQLQuery & "  AND chf.chfDelete <> 'Y'"
    'slSQLQuery = slSQLQuery & "  AND chf.chfAdfCode= 30"
    If ilWhichKey = 0 Then 'by Contract
        slSQLQuery = slSQLQuery & "  AND chf.chfCntrNo = " & llItemCode
    ElseIf ilWhichKey = 1 Then 'by Agency
        slSQLQuery = slSQLQuery & "  AND chf.chfAgfCode = " & llItemCode
    ElseIf ilWhichKey = 5 Then 'by Advertiser
        slSQLQuery = slSQLQuery & "  AND chf.chfAdfCode = " & llItemCode
    ElseIf ilWhichKey = 6 Then 'by Vehicle
        slSQLQuery = slSQLQuery & "  AND pcf.pcfVefCode = " & llItemCode
    End If
    
    If slEndDate <> "" Then
        slSQLQuery = slSQLQuery & "  AND pcf.pcfStartDate <= '" & Format(slEndDate, "yyyy-mm-dd") & "'"
    End If
    If slStartDate <> "" Then
        slSQLQuery = slSQLQuery & "  AND pcf.pcfEndDate >= '" & Format(slStartDate, "yyyy-mm-dd") & "'"
    End If
    slSQLQuery = slSQLQuery & "  AND pcf.pcfStartDate <= pcf.pcfEndDate " 'No CBS lines
    slSQLQuery = slSQLQuery & "  AND pcf.pcfDelete <> 'Y'"
    
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        Do While Not rst.EOF
            'Check Contract ChfCode
            If RptSelCb!ckcAll.Value = vbUnchecked Then       'all advt?
                If tmSelChf(0) <> 0 Then
                    blValid = False
                    For illoop = 0 To UBound(tmSelChf) - 1 Step 1
                        If tmSelChf(illoop) = rst.Fields("chfCode").Value Then
                            blValid = True
                            Exit For
                        End If
                    Next illoop
                    If blValid = False Then GoTo nextRec
                End If
            End If
            'Check Vehicle
            If tmSelVef(0) <> 0 Then
                blValid = False
                For illoop = 0 To UBound(tmSelVef) - 1 Step 1
                    If tmSelVef(illoop) = rst.Fields("pcfVefCode").Value Then
                        blValid = True
                        Exit For
                    End If
                Next illoop
                If blValid = False Then GoTo nextRec
            End If
            'Check Advertiser
            If RptSelCb!ckcAll.Value = vbUnchecked Then       'all advt?
                If tmSelAdf(0) <> 0 Then       'network feed, find all net spots matching selected advt
                    blValid = False
                    For illoop = 0 To UBound(tmSelAdf) - 1 Step 1
                        If tmSelAdf(illoop) = rst.Fields("chfAdfCode").Value Then
                            blValid = True
                            Exit For
                        End If
                    Next illoop
                    If blValid = False Then GoTo nextRec
                End If
            End If
            
            If blValid = True Then
                llUpper = UBound(tmPLPcf)
                'Contract Info
                tmPLPcf(llUpper).lCntrNo = rst.Fields("chfCntrNo").Value
                'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 5)
                tmPLPcf(llUpper).lExtCntrNo = rst.Fields("chfExtCntrNo").Value
                tmPLPcf(llUpper).sCalType = rst.Fields("chfBillCycle").Value
                tmPLPcf(llUpper).sCntrStartDate = rst.Fields("chfStartDate").Value
                tmPLPcf(llUpper).sCntrEndDate = rst.Fields("chfEndDate").Value
                Select Case rst.Fields("chfType").Value
                    Case "C": tmPLPcf(llUpper).sContractType = "Standard"
                    Case "V": tmPLPcf(llUpper).sContractType = "Reservation"
                    Case "T": tmPLPcf(llUpper).sContractType = "Remnant"
                    Case "R": tmPLPcf(llUpper).sContractType = "Direct Response"
                    Case "Q": tmPLPcf(llUpper).sContractType = "Per Inquiry" 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 7)
                    Case "S": tmPLPcf(llUpper).sContractType = "PSA"
                    Case "M": tmPLPcf(llUpper).sContractType = "Promo"
                    Case Else: tmPLPcf(llUpper).sContractType = Trim(rst.Fields("chfType").Value)
                End Select
                'tmPLPcf(llUpper).sContractType = rst.Fields("chfType").Value
                
                tmPLPcf(llUpper).iAgfCode = rst.Fields("chfAgfCode").Value
                tmPLPcf(llUpper).iAdfCode = rst.Fields("chfAdfCode").Value
                tmPLPcf(llUpper).sProduct = rst.Fields("chfProduct").Value
                'Line Info
                tmPLPcf(llUpper).tPcf.lCode = rst.Fields("pcfCode").Value
                tmPLPcf(llUpper).tPcf.iVefCode = rst.Fields("pcfVefCode").Value
                tmPLPcf(llUpper).tPcf.lChfCode = rst.Fields("chfCode").Value
                tmPLPcf(llUpper).tPcf.lTotalCost = rst.Fields("pcfTotalCost").Value
                tmPLPcf(llUpper).tPcf.iPodCPMID = rst.Fields("pcfPodCPMID").Value
                tmPLPcf(llUpper).tPcf.iLen = rst.Fields("pcfLen").Value 'Boostr Phase 2: Spot and Digital Line Combo report: show length for digital lines
                gPackDate rst.Fields("pcfEndDate").Value, tmPLPcf(llUpper).tPcf.iEndDate(0), tmPLPcf(llUpper).tPcf.iEndDate(1)
                gPackDate rst.Fields("pcfStartDate").Value, tmPLPcf(llUpper).tPcf.iStartDate(0), tmPLPcf(llUpper).tPcf.iStartDate(1)
                tmPLPcf(llUpper).tPcf.sPriceType = rst.Fields("pcfPriceType").Value
                tmPLPcf(llUpper).tPcf.lCxfCode = rst.Fields("pcfCxfCode").Value
                tmPLPcf(llUpper).tPcf.iRdfCode = rst.Fields("pcfRdfCode").Value
                ReDim Preserve tmPLPcf(0 To llUpper + 1) As PCFTYPESORT
            End If
nextRec:
            rst.MoveNext
        Loop
    End If
    rst.Close
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCffRec                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Function mReadCffRec(ilClfIndex As Integer, ilAllVersions As Integer) As Integer
'
'   iRet = mReadCffRec(ilClfIndex, ilAllVersions)
'   Where:
'       ilClfIndex (I) - CLF index (starting at 0)
'       ilAllVersions(I)- True=All versions (ignore Delete Flag); False=Latest Version only
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpperBound As Integer
    Dim ilFirst As Integer
    Dim tlCff As CFF
    Dim tlCffExt As CFFEXT    'Flight extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    ilUpperBound = UBound(tgCffCB)
    ilFirst = True
    RptSelCb!lbcLnCode.Clear
    btrExtClear hmCff   'Clear any previous extend operation
    ilExtLen = Len(tlCffExt)  'Extract operation record size
    tmCffSrchKey.lChfCode = tgChfCB.lCode
    tmCffSrchKey.iClfLine = tgClfCB(ilClfIndex).ClfRec.iLine
    tmCffSrchKey.iCntRevNo = tgClfCB(ilClfIndex).ClfRec.iCntRevNo
    tmCffSrchKey.iPropVer = tgClfCB(ilClfIndex).ClfRec.iPropVer
    tmCffSrchKey.iStartDate(0) = 0
    tmCffSrchKey.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlCff.lChfCode = tgChfCB.lCode) And (tlCff.iClfLine = tgClfCB(ilClfIndex).ClfRec.iLine) Then
        'If (tlCff.iClfVersion = tgClfCB(ilClfIndex).ClfRec.iVersion) And (tlCff.sDelete <> "Y") Then
        '    gUnpackDateForSort tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        '    ilRet = btrGetPosition(hmCff, llRecPos)
        '    slStr = slStr & "\" & Trim$(Str$(llRecPos))
        '    RptSelCb!lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
        'End If
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCff, llNoRec, -1, "UC", "CFFEXTPK", CFFEXTPK) 'Set extract limits (all records)

        ilOffSet = gFieldOffset("Cff", "CffChfCode")
        ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChfCB.lCode, 4)
        If ilRet <> BTRV_ERR_NONE Then
            mReadCffRec = False
            Exit Function
        End If
        ilOffSet = gFieldOffset("Cff", "CffClfLine")
        ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClfCB(ilClfIndex).ClfRec.iLine, 2)
        If ilRet <> BTRV_ERR_NONE Then
            mReadCffRec = False
            Exit Function
        End If
        If ilAllVersions Then
            ilOffSet = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tgClfCB(ilClfIndex).ClfRec.iCntRevNo, 2)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
        Else
            ilOffSet = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClfCB(ilClfIndex).ClfRec.iCntRevNo, 2)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
            ilOffSet = gFieldOffset("Cff", "CffDelete")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
        End If
        ilOffSet = gFieldOffset("Cff", "CffStartDate")
        ilRet = btrExtAddField(hmCff, ilOffSet, ilExtLen)  'Extract start date
        If ilRet <> BTRV_ERR_NONE Then
            mReadCffRec = False
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hmCff)    'Extract record
        ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateForSort tlCffExt.iStartDate(0), tlCffExt.iStartDate(1), slStr
                slStr = slStr & "\" & Trim$(str$(llRecPos))
                RptSelCb!lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
                ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmCff   'Clear any previous extend operation
            For illoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
                slNameCode = RptSelCb!lbcLnCode.List(illoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmCff, tgCffCB(ilUpperBound).CffRec, imCffRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadCffRec = False
                    Exit Function
                End If
                If tgClfCB(ilClfIndex).iFirstCff = -1 Then
                    tgClfCB(ilClfIndex).iFirstCff = ilUpperBound
                Else
                    tgCffCB(ilUpperBound - 1).iNextCff = ilUpperBound
                End If
                tgCffCB(ilUpperBound).iNextCff = -1
                tgCffCB(ilUpperBound).lRecPos = llRecPos
                tgCffCB(ilUpperBound).iStatus = 1 'Old and retain
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgCffCB(0 To ilUpperBound) As CFFLIST
                tgCffCB(ilUpperBound).iStatus = -1 'Not Used
                tgCffCB(ilUpperBound).iNextCff = -1
                tgCffCB(ilUpperBound).lRecPos = 0
            Next illoop
        End If
    End If
    mReadCffRec = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfRec                     *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Function mReadChfRec(llCntrIndex As Long, ilDispOnly As Integer, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************
'   iRet = mReadChfRec(ilFirstTime)
'   Where:
'       ilFirstTime (I) - True=First time getting contract
'       iRet (O)- True if record read,
'                 False if not read
    Dim ilRet As Integer    'Return status
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slDate As String
    Dim llDate As Long
    Dim ilFound As Integer
    Dim ilClfIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCheckCntr As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llLoop As Long
    'ReDim tlSdfExt(1 To 1) As SDFEXT
    ReDim tlSdfExt(0 To 0) As SDFEXT
    Dim ilVersions As Integer
    Dim llVefCode As Long
    Dim ilIncludeType As Integer        '1-11-03

    ilFound = False
    If RptSelCb!ckcAll.Value = vbChecked Then
        If llCntrIndex = LBound(tmChfAdvtExt) Then
            If (lgStartingCntrNo > 0) And ilDispOnly Then
                tmChfSrchKey1.lCntrNo = lgStartingCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfCB, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                lgStartingCntrNo = 0
                For llLoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                    If tmChfAdvtExt(llLoop).lCode = tgChfCB.lCode Then
                        ilFound = True
                        llCntrIndex = llLoop
                        Exit For
                    End If
                Next llLoop
                If Not ilFound Then
                    mReadChfRec = False
                    Exit Function
                End If
                ilFound = False
            End If
        End If
    End If
    Do
        '
        'If ilFirstTime Then
        '    ilRet = btrGetFirst(hmChf, tgChfCB, imChfRecLen, INDEXKEY0, BTRV_LOCK_NONE)
        '    ilFirstTime = False
        'Else
        '    ilRet = btrGetNext(hmChf, tgChfCB, imChfRecLen, BTRV_LOCK_NONE)
        'End If
        If RptSelCb!ckcAll.Value = vbChecked Then
            'If llCntrIndex = 0 Then
            '    If (lgStartingCntrNo > 0) And ilDispOnly Then
            '        tmChfSrchKey1.lCntrNo = lgStartingCntrNo
            '        tmChfSrchKey1.iCntRevNo = 32000
            '        tmChfSrchKey1.iPropVer = 32000
            '        ilRet = btrGetGreaterOrEqual(hmCHF, tgChfCB, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            '        lgStartingCntrNo = 0
            '    Else
            '        ilRet = btrGetFirst(hmCHF, tgChfCB, imChfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            '    End If
            'Else
            '    ilRet = btrGetNext(hmCHF, tgChfCB, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            'End If
            If llCntrIndex >= UBound(tmChfAdvtExt) Then
                mReadChfRec = False
                Exit Function
            End If
            tmChfSrchKey.lCode = tmChfAdvtExt(llCntrIndex).lCode
            ilRet = btrGetEqual(hmCHF, tgChfCB, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfRec = False
                Exit Function
            End If
            If ((tgChfCB.sSchStatus = "F") Or (tgChfCB.sSchStatus = "M")) And (tgChfCB.sDelete <> "Y") Then
                If imUpdateCntrNo Then
                    Do
                        ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, lmSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        'tmSRec = tmSpf
                        'ilRet = gGetByKeyForUpdate("Spf", hmSpf, tmSRec)
                        'tmSpf = tmSRec
                        tmSpf.lDiscCurrCntrNo = tgChfCB.lCntrNo
                        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
                ilCheckCntr = True
            Else
                ilCheckCntr = False
            End If
        Else
            If llCntrIndex > RptSelCb!lbcSelection(0).ListCount - 1 Then
                mReadChfRec = False
                Exit Function
            End If
            ilCheckCntr = False
            If RptSelCb!lbcSelection(0).Selected(llCntrIndex) Then
                ilCheckCntr = True
                slNameCode = RptSelCb!lbcCntrCode.List(llCntrIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmChfSrchKey.lCode = Val(slCode)
                ilRet = btrGetEqual(hmCHF, tgChfCB, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    mReadChfRec = False
                    Exit Function
                End If
            End If
        End If
        If ilCheckCntr Then
            'If (Not RptSelCb!ckcAll.Value) Or ((tgChfCB.sStatus <> "M") And (tgChfCB.sStatus <> "P") And (tgChfCB.sStatus <> "N")) Then
            'If discrepancies only and manually moved, proposal or new contract- bypass
            '(sType is also tested to cover errors when sStatus was not set correctly)

            '1-11-03 Include the types Remnant, Psa,or Promo if automatic scheduling, or its one of these types that have already been fully scheduled
            '(changed scheduling status in Site)
            ilIncludeType = False
            If (tgChfCB.sType = "T" And tgSpf.sSchdRemnant = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            ElseIf (tgChfCB.sType = "M" And tgSpf.sSchdPromo = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            ElseIf (tgChfCB.sType = "S" And tgSpf.sSchdPSA = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            End If

            'If (tgChfCB.sSchStatus = "A") Or ((ilDispOnly) And ((tgChfCB.sSchStatus = "M") Or (tgChfCB.sSchStatus = "P") Or (tgChfCB.sSchStatus = "N") Or (tgChfCB.sType = "T") Or (tgChfCB.sType = "Q") Or (tgChfCB.sType = "M") Or (tgChfCB.sType = "S"))) Or (ilDispOnly And gIsCntrRemote(tgChfCB, hmVsf)) Then
            If (tgChfCB.sSchStatus = "A") Or ((ilDispOnly) And (Not ilIncludeType)) Or (ilDispOnly And tgChfCB.sType = "Q") Or (ilDispOnly And gIsCntrRemote(tgChfCB, hmVsf)) Then
                ilFound = False
            Else
                'Get all spots
                'RptSelCb!lbcSort.Clear
                slStartDate = Format$(llStartDate, "m/d/yy")
                slEndDate = Format$(llEndDate, "m/d/yy")
                'If lgMtfNoRecs > 0 Then
                '    llVefCode = -1
                'Else
                '    llVefCode = tgChfCB.lVefCode
                'End If
                llVefCode = 0   'Use key by Contract instead of Vehicle, Contract
                ReDim tlSdfExtSort(0 To 0) As SDFEXTSORT
                If tgChfCB.lVefCode > 0 Then                      'all same veh on this order
                    '5-9-11 Remove all the invalid bb spots that doesnt belong
                    'SDF opened in calling rtn
                    ilRet = gRemoveBBSpots(hmSdf, CInt(tgChfCB.lVefCode), 0, slStartDate, slEndDate, tgChfCB.lCode, 0)
                    ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                Else                                            'possibly multiple vehicles on order
                    If ((tgChfCB.sStatus = "M") Or (tgChfCB.sStatus = "P") Or (tgChfCB.sStatus = "N") Or (tgChfCB.sType = "T") Or (tgChfCB.sType = "Q") Or (tgChfCB.sType = "M") Or (tgChfCB.sType = "S")) Then
                        ''PSA/ Promo,.. can't have MG so only get spots for specified dates to reduce number of spots
                        'slStartDate = Format$(llStartDate, "m/d/yy")
                        'slEndDate = Format$(llEndDate, "m/d/yy")
                        ''ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, slStartDate, slEndDate, RptSelCb!lbcSort, tlSdfExt())
                        mRemoveBBSpotSetup slStartDate, slEndDate, tgChfCB
                        ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                    Else
                        'ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, "", "", RptSelCb!lbcSort, tlSdfExt())
                        mRemoveBBSpotSetup slStartDate, slEndDate, tgChfCB
                        ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                   End If
                End If
                gUnpackDate tgChfCB.iStartDate(0), tgChfCB.iStartDate(1), slDate

                llSDate = gDateValue(slDate)
                gUnpackDate tgChfCB.iEndDate(0), tgChfCB.iEndDate(1), slDate
                llEDate = gDateValue(slDate)
                If (llEDate >= llStartDate) And (llSDate <= llEndDate) Then
                    ilFound = True
                Else
                    'Determine if any spots are within date span as contract is not
                    'For ilLoop = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                    '    gUnpackDateLong tlSdfExt(ilLoop).iDate(0), tlSdfExt(ilLoop).iDate(1), llDate
                    For llLoop = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                        gUnpackDateLong tlSdfExt(llLoop).iDate(0), tlSdfExt(llLoop).iDate(1), llDate
                        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                            ilFound = True
                            Exit For
                        End If
                    'Next ilLoop
                    Next llLoop
                End If
            End If
        End If
        llCntrIndex = llCntrIndex + 1
    Loop While Not ilFound
    mReadChfRec = True
    ReDim tgClfCB(0 To 0) As CLFLIST
    tgClfCB(0).iStatus = -1 'Not Used
    tgClfCB(0).lRecPos = 0
    tgClfCB(0).iFirstCff = -1
    ReDim tgCffCB(0 To 0) As CFFLIST
    tgCffCB(0).iStatus = -1 'Not Used
    tgCffCB(0).lRecPos = 0
    tgCffCB(0).iNextCff = -1
    ilVersions = 2
    If mReadClfRec(ilVersions) Then
        For ilClfIndex = LBound(tgClfCB) To UBound(tgClfCB) - 1 Step 1
            If Not mReadCffRec(ilClfIndex, False) Then
                mReadChfRec = False
                Exit Function
            End If
        Next ilClfIndex

        ''Get all spots
        'RptSelCb!lbcSort.Clear
        'If tgChfCB.iVefCode > 0 Then
        '    ilRet = gObtainCntrSpot(tgChfCB.iVefCode, False, tgChfCB.lCode, -1, "", "", RptSelCb!lbcSort, tlSdfExt())
        'Else
        '    If ((tgChfCB.sStatus = "M") Or (tgChfCB.sStatus = "P") Or (tgChfCB.sStatus = "N") Or (tgChfCB.sType = "T") Or (tgChfCB.sType = "Q") Or (tgChfCB.sType = "M") Or (tgChfCB.sType = "S")) Then
        '        'PSA/ Promo,.. can't have MG so only get spots for specified dates to reduce number of spots
        '        slStartDate = Format$(llStartDate, "m/d/yy")
        '        slEndDate = Format$(llEndDate, "m/d/yy")
        '        ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, slStartDate, slEndDate, RptSelCb!lbcSort, tlSdfExt())
        '    Else
        '        ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, "", "", RptSelCb!lbcSort, tlSdfExt())
        '   End If
        'End If
    Else
        mReadChfRec = False
    End If
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadClfRec                     *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Function mReadClfRec(ilVersions As Integer) As Integer
'   iRet = mReadClfRec(ilVersions)
'   Where:
'       illVersions(I)- 0=All versions; 1=Latest version only; 2=Latest Fully Schedule Versions
'       iRet (O)- True if record read,
'                 False if not read
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim illoop As Integer
    Dim slStr As String
    Dim tlClf As CLF
    Dim tlClfExt As CLFEXT    'Contract line extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim slLine As String
    Dim slVersion As String
    Dim ilAddLine As Integer
    Dim ilOffSet As Integer

    ReDim tgClfCB(0 To 0) As CLFLIST
    ilUpperBound = UBound(tgClfCB)
    tgClfCB(ilUpperBound).iStatus = -1 'Not Used
    tgClfCB(ilUpperBound).lRecPos = 0
    tgClfCB(ilUpperBound).iFirstCff = -1
    RptSelCb!lbcLnCode.Clear
    btrExtClear hmClf   'Clear any previous extend operation
    ilExtLen = Len(tlClfExt)  'Extract operation record size
    tmClfSrchKey.lChfCode = tgChfCB.lCode
    tmClfSrchKey.iLine = 0
    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlClf.lChfCode = tgChfCB.lCode) Then 'And ((ilVersions = 0) Or (ilVersions = 2) Or (tlClf.sDelete <> "Y")) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmClf, llNoRec, -1, "UC", "CLFEXTPK", CLFEXTPK) 'Set extract limits (all records)
        If (ilVersions = 0) Then 'Or (ilVersion = 2) Then
            ilOffSet = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tgChfCB.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
        Else
            ilOffSet = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChfCB.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
            ilOffSet = gFieldOffset("Clf", "ClfDelete")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
        End If
        ilOffSet = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddField(hmClf, ilOffSet, ilExtLen - 3) 'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            mReadClfRec = False
            Exit Function
        End If
        ilOffSet = gFieldOffset("Clf", "ClfSchStatus")
        ilRet = btrExtAddField(hmClf, ilOffSet, 1) 'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            mReadClfRec = False
            Exit Function
        End If
        ilOffSet = gFieldOffset("Clf", "ClfPropVer")
        ilRet = btrExtAddField(hmClf, ilOffSet, 2) 'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            mReadClfRec = False
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                'Only show the latest line
                ilAddLine = True
                If ilVersions = 1 Then
                    For illoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
                        slNameCode = RptSelCb!lbcLnCode.List(illoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slLine)
                        ilRet = gParseItem(slNameCode, 2, "\", slVersion)
                        If tlClfExt.iLine = Val(slCode) Then
                            If tlClfExt.iCntRevNo > Val(slVersion) Then
                                RptSelCb!lbcLnCode.RemoveItem illoop
                            Else
                                ilAddLine = False
                            End If
                            Exit For
                        End If
                    Next illoop
                ElseIf ilVersions = 2 Then
                    'Manually schedule (M) are only shown when running spot placement
                    If (tlClfExt.sSchStatus = "F") Or (tlClfExt.sSchStatus = "M") Then
                        For illoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
                            slNameCode = RptSelCb!lbcLnCode.List(illoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slLine)
                            ilRet = gParseItem(slNameCode, 2, "\", slVersion)
                            If tlClfExt.iLine = Val(slCode) Then
                                If tlClfExt.iCntRevNo > Val(slVersion) Then
                                    RptSelCb!lbcLnCode.RemoveItem illoop
                                Else
                                    ilAddLine = False
                                End If
                                Exit For
                            End If
                        Next illoop
                    Else
                        ilAddLine = False
                    End If
                End If
                If ilAddLine Then
                    slStr = Trim$(str$(tlClfExt.iLine))
                    Do While Len(slStr) < 4
                        slStr = "0" & slStr
                    Loop
                    If ilVersions = 0 Then
                        slVersion = Trim$(str$(999 - tlClfExt.iCntRevNo))
                        Do While Len(slVersion) < 3
                            slVersion = "0" & slVersion
                        Loop
                    Else
                        slVersion = Trim$(str$(tlClfExt.iCntRevNo))
                    End If
                    slStr = slStr & "\" & slVersion
                    slStr = slStr & "\" & Trim$(str$(llRecPos))
                    RptSelCb!lbcLnCode.AddItem slStr
                End If
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmClf   'Clear any previous extend operation
            For illoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
                slNameCode = RptSelCb!lbcLnCode.List(illoop)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmClf, tgClfCB(ilUpperBound).ClfRec, imClfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadClfRec = False
                    Exit Function
                End If
                tgClfCB(ilUpperBound).iFirstCff = -1
                tgClfCB(ilUpperBound).lRecPos = llRecPos
                tgClfCB(ilUpperBound).iStatus = 1 'Old line
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgClfCB(0 To ilUpperBound) As CLFLIST
                tgClfCB(ilUpperBound).iStatus = -1 'Not Used
                tgClfCB(ilUpperBound).iFirstCff = -1
                tgClfCB(ilUpperBound).lRecPos = 0
            Next illoop
        End If
    End If
    mReadClfRec = True
    Exit Function
End Function

'                   mSetCostType - set up bit string in ilCostType for
'                   the types of spots to include
'                   Used in spots by ADvt and Spots by Date & Time
'                   Include: Charged, 0.00, ADU, Bonus, Extra, Fill
'                   No Charge, Nc MG, recaptureable, & spinoff
'                   <output> ilCosttype - as bit string
'     5-25-05 Add bit 10 for BB testing
Sub mSetCostType(ilCostType As Integer)
    ilCostType = 0
    If RptSelCb!ckcSelC5(0).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_CHARGE          'bit 0
    End If
    If RptSelCb!ckcSelC5(1).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_00
    End If
    If RptSelCb!ckcSelC5(2).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_ADU
    End If
    If RptSelCb!ckcSelC5(3).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_BONUS
    End If
    If RptSelCb!ckcSelC5(4).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_EXTRA
    End If
    If RptSelCb!ckcSelC5(5).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_FILL
    End If
    If RptSelCb!ckcSelC5(6).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_NC
    End If
    If RptSelCb!ckcSelC5(7).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_MG
    End If
    If RptSelCb!ckcSelC5(8).Value = vbChecked Then
        ilCostType = ilCostType Or SPOT_RECAP
    End If
    If RptSelCb!ckcSelC5(9).Value = vbChecked Then                     'bit 9
        ilCostType = ilCostType Or SPOT_SPINOFF
    End If
    If RptSelCb!ckcSelC5(10).Value = vbChecked Then                    'bit 10
        ilCostType = ilCostType Or SPOT_BB
    End If
End Sub

'*************************************************************************
'
'           mSpotSalesTitle
'               <Output>    ilMissedType : 0 = show standard spot
'                           (process non mg spot (sched, missed)
'                                           1 = process missed part of mg
'                                           2 = process mg part of mg
'                           slIncludeTitle - string for inclusion on
'                                            header
'
'          Created 1/21/97 D.H.
'
'*************************************************************************
Sub mSpotSalesTitle(ilMissedType As Integer, slInclude As String, slExclude As String)
    If RptSelCb!rbcSelC7(0).Value Then                'ordered
        slInclude = "Include- Ordered"
    ElseIf RptSelCb!rbcSelC7(1).Value Then          'aired
        slInclude = "Include- Aired"
    Else
        slInclude = "Include- Aired/Pkg Ordered"
    End If
    If RptSelCb!ckcSelC3(0).Value = vbChecked Then
        ilMissedType = 1
    End If
    If RptSelCb!ckcSelC3(1).Value = vbChecked Then
        ilMissedType = ilMissedType + 2
    End If
    If RptSelCb!ckcSelC3(2).Value = vbChecked Then
        ilMissedType = ilMissedType + 4
    End If
    gIncludeExcludeRbc RptSelCb!rbcSelC4(0), slInclude, slExclude, "Spots"
    gIncludeExcludeRbc RptSelCb!rbcSelC4(1), slInclude, slExclude, "Units"
    gIncludeExcludeCkc RptSelCb!ckcSelC3(0), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelCb!ckcSelC3(1), slInclude, slExclude, "Cancel"
    gIncludeExcludeCkc RptSelCb!ckcSelC3(2), slInclude, slExclude, "Hidden"
    gIncludeExcludeCkc RptSelCb!ckcSelC6(0), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelCb!ckcSelC6(1), slInclude, slExclude, "Resv"
    gIncludeExcludeCkc RptSelCb!ckcSelC6(2), slInclude, slExclude, "Rem"
    gIncludeExcludeCkc RptSelCb!ckcSelC6(3), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelCb!ckcSelC6(4), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(0), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(1), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(2), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(3), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(4), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(5), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(6), slInclude, slExclude, "NC"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(7), slInclude, slExclude, "MG"
    gIncludeExcludeCkc RptSelCb!ckcSelC5(9), slInclude, slExclude, "Spinoff"
End Sub

'                   sub mTestCostType - for Spots by Advt and spots by Date & time
'                   Include different types of spots test
'                   <input> ilCosttype - bit string based on user request of
'                           types of spots to include
'                           slStrCost - string defining cost of spot ($ value as string
'                                       or text such as ADU , bonus, etc.
'                           slSchstatus - spot sched status to test for MG/outside 10-18-10
'                   <output> ilOk - false if not a spot to report
'           3-24-03 Change test of extra vs fill.  Spot is set to + or - Fill
'           10-20-10 test incl/exc of mg spots (sched line rates or mg/outside spots.  fills that have sch status of "O"
'                    is not considered a mg/outside
Sub mTestCostType(ilOk As Integer, ilCostType As Integer, slStrCost As String, slSchStatus As String)
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
        'if $ mg/out, test for inclusion/exclusion of spot
        If ((Trim$(slSchStatus) = "O" Or Trim$(slSchStatus) = "G") And ((ilCostType And SPOT_MG) <> SPOT_MG)) Then
        '10-18-10 treat mg/outsides same as the schedule line rate of MG for inclusion/exclusion
            ilOk = False
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
    ElseIf (Trim$(slStrCost) = "MG") And ((ilCostType And SPOT_MG) <> SPOT_MG) Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Recapturable" And (ilCostType And SPOT_RECAP) <> SPOT_RECAP Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Spinoff" And (ilCostType And SPOT_SPINOFF) <> SPOT_SPINOFF Then
            ilOk = False
    End If
    If ((Trim$(slSchStatus) = "O" Or Trim$(slSchStatus) = "G") And ((ilCostType And SPOT_MG) <> SPOT_MG)) Then
    '10-18-10 treat mg/outsides same as the spot cost of MG for inclusion/exclusion
    'fills are considered OUtsides, but not for inclusion/exclusion of a mg type spot
        If ((Trim$(slStrCost) <> "- Fill") And (Trim$(slStrCost) <> "+ Fill")) Then
            ilOk = False
        End If
    End If
    Exit Sub
End Sub

Private Function mObtainCntrForDates(llStartDate As Long, llEndDate As Long) As Integer
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim llChf As Long
    Dim ilFound As Integer
    Dim llPrevChfCode As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim llRecPos As Long
    
    mObtainCntrForDates = False
    slCntrStatus = "HO"
    slCntrType = "C"
    ilHOType = 1
    sgCntrForDateStamp = ""
    slStartDate = Format$(llStartDate, "m/d/yy")
    slEndDate = Format$(llEndDate, "m/d/yy")
    ilRet = gObtainCntrForDate(RptSelCb, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
    If (ilRet <> CP_MSG_NOPOPREQ) And (ilRet <> CP_MSG_NONE) Then
        Exit Function
    End If

    llPrevChfCode = -1
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    gPackDate slStartDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
    tmSdfSrchKey4.lChfCode = 0
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)


        ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Function
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If llPrevChfCode <> tmSdf.lChfCode Then
                    ilFound = False
                    For llChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                        If tmSdf.lChfCode = tmChfAdvtExt(llChf).lCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next llChf
                    If Not ilFound Then
                        tmChfAdvtExt(UBound(tmChfAdvtExt)).lCode = tmSdf.lChfCode
                        ReDim Preserve tmChfAdvtExt(LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) + 1) As CHFADVTEXT
                    End If
                End If
                llPrevChfCode = tmSdf.lChfCode
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    mObtainCntrForDates = True
End Function

'               called by mObtainSDF to filter spots based on selectivity
Private Function mSpotDateFilter(tlPlSDF As SPOTTYPESORT, llStartDate As Long, llEndDate As Long, llStartTime As Long, llEndTime As Long, ilSpotType As Integer, ilBillType As Integer, ilIncludePSA As Integer, ilMissedType As Integer, ilISCIOnly As Integer, ilCostType As Integer, ilByOrderOrAir As Integer, ilIncludeType As Integer, llContrCode As Long, ilPropPrice As Integer, llLLDate As Long, tlCntTypes As CNTTYPES) As Integer
'   where:
'       tlPlSDF(I/O)- entry to store spot info, plus other information that may be to be shown on output
'       llStartDate - user requested start date
'       llEndDate - user requested end date
'       llStarttime - user requested start time
'       llEndTime - user requested end time
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
'       ilByOrderOrAir(I)- 0=Order; 1=Aired   , 3 = as aired, pkg ordered (option 3 added 4-6-99)
'       ilIncludeType - true to test contract type inclusions, else false to ignore test
'       llContrCode - if selective contract, code # (else 0 for all)
'       ilLocal - true to include local spots
'       ilFeed - true to include network (feed) spots
'       ilPropPrice - true to show proposal price vs actual price
    Dim ilReturn As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim llMissedDate As Long
    Dim llTime As Long
    Dim slSpotDate As String
    Dim ilOk As Integer
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String

    If tlPlSDF.tSdf.lChfCode > 0 Then      'network feed spot (vs contract spot) has no lines to access
        '4/6/99 get line first, to send to filter routine
        tmClfSrchKey.lChfCode = tlPlSDF.tSdf.lChfCode
        tmClfSrchKey.iLine = tlPlSDF.tSdf.iLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        ilReturn = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilReturn = BTRV_ERR_NONE) And (tmClf.lChfCode = tlPlSDF.tSdf.lChfCode) And (tmClf.iLine = tlPlSDF.tSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))   'And (tmClf.sSchStatus = "A")
            ilReturn = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilReturn = BTRV_ERR_NONE) And (tmClf.lChfCode = tlPlSDF.tSdf.lChfCode) And (tmClf.iLine = tlPlSDF.tSdf.iLineNo) Then
            ilOk = True
        Else
            ilOk = False
        End If
        tlPlSDF.sLiveCopy = tmClf.sLiveCopy    '5-31-12
    Else                                        'network feed spot
        tlPlSDF.sCostType = "Feed"
        ilOk = True
    End If

    'check day of week selectivity
     gUnpackDate tlPlSDF.tSdf.iDate(0), tlPlSDF.tSdf.iDate(1), slDate
    ilRet = gWeekDayStr(slDate)
    If RptSelCb!ckcSelC8(ilRet) = vbUnchecked Then
        ilOk = False
    End If

    'Build sort key
    'ilOK = True
    If ilByOrderOrAir = 0 Then
        'Schedule and Missed only
        If (tlPlSDF.tSdf.sSchStatus = "S") Or (tlPlSDF.tSdf.sSchStatus = "G") Or (tlPlSDF.tSdf.sSchStatus = "O") Then
            If (tlPlSDF.tSdf.sSchStatus <> "S") And (tlPlSDF.tSdf.sSpotType <> "X") Then      '4-9-10 fills follow vehicles where they are sched, not the orig vehicle
                ilOk = False
            End If
        Else
           
            If (tlPlSDF.tSdf.sSchStatus = "H") Then
                If (ilMissedType And &H4) <> &H4 Then
                    ilOk = False
                End If
            ElseIf (tlPlSDF.tSdf.sSchStatus = "C") Then
                If (ilMissedType And &H2) <> &H2 Then
                    ilOk = False
                End If
            End If
        End If
    Else                      'as aired, or as aired/pkg ordered
        If ilSpotType = 1 Then  'Scheduled only
            If (tlPlSDF.tSdf.sSchStatus <> "S") And (tlPlSDF.tSdf.sSchStatus <> "G") And (tlPlSDF.tSdf.sSchStatus <> "O") Then
                ilOk = False
            Else
                If ilBillType = 1 Then  'Billed only
                    If (tlPlSDF.tSdf.sBill <> "Y") Then
                        ilOk = False
                    End If
                ElseIf ilBillType = 2 Then  'Unbilled only
                    If (tlPlSDF.tSdf.sBill = "Y") Then
                        ilOk = False
                    End If
                ElseIf ilBillType = 0 Then  'Neither
                    ilOk = False
                End If
            End If
        ElseIf ilSpotType = 2 Then  'Missed only
            If (tlPlSDF.tSdf.sSchStatus = "S") Or (tlPlSDF.tSdf.sSchStatus = "G") Or (tlPlSDF.tSdf.sSchStatus = "O") Then
                ilOk = False
            Else
                If (tlPlSDF.tSdf.sSchStatus = "H") Then
                    If (ilMissedType And &H4) <> &H4 Then
                        ilOk = False
                    End If
                ElseIf (tlPlSDF.tSdf.sSchStatus = "C") Then
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
            If ilByOrderOrAir = 2 Then          'aired/pkg ordered only
                'If tmClf.sType = "H" And (tlPlSdf.tSdf.sSchStatus <> "S") Then
                If tmClf.sType <> "H" Then          'not hidden (for pkg), must be conventional
                    If (tlPlSDF.tSdf.sSchStatus <> "S") And (tlPlSDF.tSdf.sSchStatus <> "G") And (tlPlSDF.tSdf.sSchStatus <> "O") Then
                        If (tlPlSDF.tSdf.sSchStatus = "H") Then    '6/15/99
                            If (ilMissedType And &H4) <> &H4 Then
                                ilOk = False
                            End If
                        ElseIf (tlPlSDF.tSdf.sSchStatus = "C") Then
                            If (ilMissedType And &H2) <> &H2 Then
                                ilOk = False
                            End If
                        Else
                            If (ilMissedType And &H1) <> &H1 Then
                                ilOk = False
                            End If
                        End If
                    End If
                   ' ilOK = False        'ignore the mgs and outsides here for packages
                'else pkg, automatically include missed
                Else                        '7/2/99 hidden (pkg line), ignore here ifmg or out spot  (will be processed later in mobtainmissedformg)
                    'if missed, cancelled, hidden, need to be included for package
                    If (tlPlSDF.tSdf.sSchStatus = "O" Or tlPlSDF.tSdf.sSchStatus = "G") Then
                        'ilOk = False
                        'need to get the smf to see what if the original missed date was within the requested dates
                        tmSmfSrchKey2.lCode = tlPlSDF.tSdf.lCode
                        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            'check if original missed within requested period
                            gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
                            'If llMissedDate < llStartDate Or llMissedDate > llEndDate Then
                            If (llMissedDate < llStartDate Or llMissedDate > llEndDate) Or (tmClf.iVefCode <> tmSdf.iVefCode) Or ((llMissedDate >= llStartDate And llMissedDate <= llEndDate) And (tmClf.iVefCode = tmSdf.iVefCode)) Then        '8-4-16 if different vehicles, ignore because it will be based on ordered
                                ilOk = False
                            End If
                        Else
                            ilOk = False
                        End If
                    End If
                End If
            'End If
            Else
                If (tlPlSDF.tSdf.sSchStatus <> "S") And (tlPlSDF.tSdf.sSchStatus <> "G") And (tlPlSDF.tSdf.sSchStatus <> "O") Then
                    If (tlPlSDF.tSdf.sSchStatus = "H") Then
                        If (ilMissedType And &H4) <> &H4 Then
                            ilOk = False
                        End If
                    ElseIf (tlPlSDF.tSdf.sSchStatus = "C") Then
                        If (ilMissedType And &H2) <> &H2 Then
                            ilOk = False
                        End If
                    Else
                        If (ilMissedType And &H1) <> &H1 Then
                            ilOk = False
                        End If
                    End If
                Else
                    If ilBillType = 1 Then  'Billed only
                        If (tlPlSDF.tSdf.sBill <> "Y") Then
                            ilOk = False
                        End If
                    ElseIf ilBillType = 2 Then  'Unbilled only
                        If (tlPlSDF.tSdf.sBill = "Y") Then
                            ilOk = False
                        End If
                    ElseIf ilBillType = 0 Then  'Neither
                        ilOk = False
                    End If
                End If
            End If
        End If
    End If
    If (ilOk) Then                          'test time filters and x-mid flags
        gUnpackTimeLong tlPlSDF.tSdf.iTime(0), tlPlSDF.tSdf.iTime(1), False, llTime
        If llTime < llStartTime Or llTime >= llEndTime Then
            ilOk = False
        End If

    End If
    'If (ilOk) And (Not ilIncludePSA) Then
        If tlPlSDF.tSdf.lChfCode > 0 Then          'only test for psa spot if its a contract spot (vs network feed spot)
            tmChfSrchKey.lCode = tlPlSDF.tSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            '1-3-18 psa and promo tested below
'                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
'                        ilOk = False
'                    End If
        End If
    'End If
    If (ilOk And ilIncludeType) Then
        If tlPlSDF.tSdf.lChfCode > 0 Then          'only test contr types if its a contract spot (vs network feed spot)
            mFilterCntTypes tmChf, tlCntTypes, ilOk
            '1-3-18 use common subroutine to test cnt types
'                    If tmChf.sType = "C" And Not (RptSelCb!ckcSelC6(0).Value = vbChecked) Then  'include std cntrs?
'                        ilOk = False
'                    End If
'                    If tmChf.sType = "V" And Not (RptSelCb!ckcSelC6(1).Value = vbChecked) Then   'include reserves?
'                        ilOk = False
'                    End If
'                    If tmChf.sType = "T" And Not (RptSelCb!ckcSelC6(2).Value = vbChecked) Then   'include remnants?
'                        ilOk = False
'                    End If
'                    If tmChf.sType = "R" And Not (RptSelCb!ckcSelC6(3).Value = vbChecked) Then   'direct response?
'                        ilOk = False
'                    End If
'                    If tmChf.sType = "Q" And Not (RptSelCb!ckcSelC6(4).Value = vbChecked) Then   'per inquiry?
'                        ilOk = False
'                    End If
        End If
    End If
    If (ilOk) And (ilISCIOnly) Then
        If tlPlSDF.tSdf.lChfCode > 0 Then          'copy only applies to contract spots
        '**********  NEED TO IMPLEMENT NETWORK COPY   **************
            tmSdf = tlPlSDF.tSdf
            mObtainCopy slProduct, slZone, slCart, slISCI
            If Len(slISCI) <= 0 Then
                tmAdfSrchKey.iCode = tlPlSDF.tSdf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If tmAdf.sShowISCI <> "Y" Then
                    tmChfSrchKey.lCode = tlPlSDF.tSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If tmChf.iAgfCode > 0 Then     'agency exists
                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If tmAgf.sShowISCI <> "Y" Then
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
    If ilOk Then            'check for selective contract (zero if all contracts)
        If llContrCode <> 0 Then
            If llContrCode <> tlPlSDF.tSdf.lChfCode Then
                ilOk = False
            End If
        End If
    End If
    
    'Test if Open or Close BB, ignore if in the future
    If tlPlSDF.tSdf.sSpotType = "O" Or tlPlSDF.tSdf.sSpotType = "C" Then
        gUnpackDate tlPlSDF.tSdf.iDate(0), tlPlSDF.tSdf.iDate(1), slSpotDate
        If gDateValue(slSpotDate) >= llLLDate Then   'is the spot date >= to last log date?  If so, ignore
            ilOk = False
        End If
    End If
    
    If ilOk Then
        'get line first, to send to filter routine
        'tmClfSrchKey.lchfcode = tlPlSdf.tSdf.lchfcode
        'tmClfSrchKey.iLine = tlPlSdf.tSdf.iLineNo
        'tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        'tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        'ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        'Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lchfcode = tlPlSdf.tSdf.lchfcode) And (tmClf.iLine = tlPlSdf.tSdf.iLineNo) And (tmClf.sSchStatus = "A")
        '    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
        'If (ilRet = BTRV_ERR_NONE) And (tmClf.lchfcode = tlPlSdf.tSdf.lchfcode) And (tmClf.iLine = tlPlSdf.tSdf.iLineNo) Then
        If tlPlSDF.tSdf.lChfCode > 0 Then          'only test for spot costs if its a contract spot (vs network feed spot)
            ilRet = gGetSpotPrice(tlPlSDF.tSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, tlPlSDF.sCostType)
'                    tlPlSDF.iVefCode = tmClf.iVefCode
            If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                mTestCostType ilOk, ilCostType, tlPlSDF.sCostType, tlPlSDF.tSdf.sSchStatus    '10-18-10 test mg/out
                '5-25-05 test for inclusion/exclusion of BB spots
                If (tlPlSDF.tSdf.sSpotType = "O" Or tlPlSDF.tSdf.sSpotType = "C") And (ilCostType And SPOT_BB) <> SPOT_BB Then
                    ilOk = False
                End If
            End If

            If ilOk And ilPropPrice Then        'passed filters and need to see propprice vs actual spot price
                'if spot is a fill, dont use the prop price, its zero
                If InStr(tlPlSDF.sCostType, "Fill") = 0 Then       'may be a fill
                    ilRet = gGetSpotFlight(tlPlSDF.tSdf, tmClf, hmCff, hmSmf, tmCff)
                    tlPlSDF.sCostType = gLongToStrDec(tmCff.lPropPrice * 100, 2)
                End If
            End If
        Else
            tlPlSDF.sCostType = "Feed"
        End If
            
    End If
    mSpotDateFilter = ilOk
    Exit Function
End Function

Private Function mSetStatus(tlSdfExt() As SDFEXT, llSdfIndex As Long) As Integer
    mSetStatus = True
    '3-28-05 if open/close billboard, not an invalid length
    If (tlSdfExt(llSdfIndex).iLen <> tmClf.iLen) Then  'spot length & line length not equal, if spot type is blank its invalid;
                                                        'otherwise its a billboard

        If (tlSdfExt(llSdfIndex).sSpotType <> "X") And (tlSdfExt(llSdfIndex).sSpotType <> "O" And tlSdfExt(llSdfIndex).sSpotType <> "C") Then
            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H8  'Length invalid
            mSetStatus = False
        End If
    End If

End Function

Private Function mObtainCount(tlSdfExt() As SDFEXT, slOrigMissedDate As String, llDate As Long, llStartDate As Long, llEndDate As Long, llChkStartDate As Long, llChkEndDate As Long, slSchDate As String, slSdfTime As String, llSdfTime As Long, ilCff As Integer, llStartTime() As Long, llEndTime() As Long, ilFlag As Integer, ilCffSpots As Integer) As Integer
    Dim llSdfIndex As Long
    Dim ilVefFound As Integer
    Dim ilRet As Integer
    Dim ilCVsf As Integer
    Dim ilSdfSpots As Integer
    Dim ilVefCode As Integer
    Dim slSdfDate As String
    Dim ilDay As Integer
    Dim ilFound As Integer
    Dim ilTime As Integer
    
    mObtainCount = True
    For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
            ilVefFound = False
            If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                'tmSmfSrchKey.lChfCode = tmClf.lChfCode
                'tmSmfSrchKey.iLineNo = tmClf.iLine
                ''slDate = Format$(llChkStartDate, "m/d/yy")
                ''gPackDate slDate, ilDate0, ilDate1
                'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                'DL: 4-20-05 use key2 instead of key0 for speed
                tmSmfSrchKey2.lCode = tmSdf.lCode
                ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                    If tmSmf.lSdfCode = tmSdf.lCode Then
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                        tlSdfExt(llSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                        For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(ilCVsf) > 0 Then
                                If tmSmf.iOrigSchVef = tmVsf.iFSCode(ilCVsf) Then
                                    ilVefFound = True
                                    Exit For
                                End If
                            End If
                        Next ilCVsf
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                Loop
            Else
                For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilCVsf) > 0 Then
                        If tlSdfExt(llSdfIndex).iVefCode = tmVsf.iFSCode(ilCVsf) Then
                            ilVefFound = True
                            Exit For
                        End If
                    End If
                Next ilCVsf
            End If

            gUnpackDateLong tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    If tmSdf.sTracer = "*" Then
                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
                        ilVefFound = True
                    End If
                End If
                If (tlSdfExt(llSdfIndex).sSpotType = "O") Or (tlSdfExt(llSdfIndex).sSpotType = "C") Then
                    tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
                Else
                    If Not ilVefFound Then
                        '4/6/16: Show MG and Outside with Original vehicle is error as an error
                        'If (tlSdfExt(llSdfIndex).sSchStatus <> "G" And tlSdfExt(llSdfIndex).sSchStatus <> "O") Then
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H4  'Vehicle invalid
                                'mCntrSchdSpotChk = False
                                mObtainCount = False
                            End If
                        'End If
                        tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
                    End If
                End If

            End If
            '6/8/16: Replaced GoSub
            'GoSub lSetStatus
            If Not mSetStatus(tlSdfExt(), llSdfIndex) Then
                mObtainCount = False
            End If
        End If
    Next llSdfIndex
    'Scan scheduled spots- checking if from this date span
    For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        If tmVsf.iFSCode(ilCVsf) > 0 Then
            ilSdfSpots = 0
            For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                If (tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine) Then
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        ''Obtain original dates
                        'tmSmfSrchKey.lChfCode = tmClf.lChfCode
                        'tmSmfSrchKey.iLineNo = tmClf.iLine
                        ''slDate = Format$(llChkStartDate, "m/d/yy")
                        ''gPackDate slDate, ilDate0, ilDate1
                        'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                        'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                        'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        'DL: 4-20-05 use key2 instead of key0 for speed
                        tmSmfSrchKey2.lCode = tmSdf.lCode
                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                            If tmSmf.lSdfCode = tmSdf.lCode Then
                                ilVefCode = tmSmf.iOrigSchVef
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        Loop
                    Else
                        ilVefCode = tlSdfExt(llSdfIndex).iVefCode
                    End If
                End If
                If (tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine) And (tmVsf.iFSCode(ilCVsf) = ilVefCode) Then
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        gUnpackDate tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), slSchDate
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slSdfDate
                        If ((gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate)) Or ((gDateValue(slSchDate) >= llChkStartDate) And (gDateValue(slSchDate) <= llChkEndDate)) Then
                            'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                            '    mCntrSchdSpotChk = False
                            'End If
                            tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                            tlSdfExt(llSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                            If ((gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate)) Then
                                If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                    ilSdfSpots = ilSdfSpots + 1
                                End If
                                tlSdfExt(llSdfIndex).lMdDate = -tlSdfExt(llSdfIndex).lMdDate    'Use negative to indicate missed counted
                            End If
                            mGetLegalTimes slOrigMissedDate, tmSmf.iGameNo, llStartTime(), llEndTime()
                            gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slSdfTime
                            llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
                            'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
                                'illegal time
                                'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                '    mCntrSchdSpotChk = False
                                '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H2
                                'End If
                            'End If
                            'Check Day
                            ilDay = gWeekDayStr(slSdfDate)
                            If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) And (tgCffCB(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                    'illegal Date
                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                    '    mCntrSchdSpotChk = False
                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    'End If
                                End If
                            Else
                                If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) Then
                                    'illegal Date
                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                    '    mCntrSchdSpotChk = False
                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    'End If
                                End If
                            End If
                        End If
                    Else
                        gUnpackDate tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), slSdfDate
                        If (gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate) Then
                            'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                            '    mCntrSchdSpotChk = False
                            'End If
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                ilSdfSpots = ilSdfSpots + 1
                            End If
                            tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
                            'If scheduled or missed spot, check its time
                            mGetLegalTimes slSdfDate, tlSdfExt(llSdfIndex).iGameNo, llStartTime(), llEndTime()
                            gUnpackTime tlSdfExt(llSdfIndex).iTime(0), tlSdfExt(llSdfIndex).iTime(1), "A", "1", slSdfTime
                            llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))

                            ilFound = False
                            For ilTime = 0 To 6 Step 1
                                If (llStartTime(ilTime) >= 0 And llEndTime(ilTime) > 0) Then
                                    'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
                                    If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                            Next ilTime
                            If (tlSdfExt(llSdfIndex).sSpotType <> "X") And (Not ilFound) Then
                                'mCntrSchdSpotChk = False
                                mObtainCount = False
                                tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H2
                            End If
      'Doug***************
                            ilFlag = False  'Set a flag variable
                            If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
                                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                If tmSdf.sTracer = "*" Then
                                    tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
                                    ilFlag = True
                                End If
                            End If
                            If Not ilFlag Then
                                ilDay = gWeekDayStr(slSdfDate)
                                If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                    If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) And (tgCffCB(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                        'illegal Date
                                        If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                            'mCntrSchdSpotChk = False
                                            mObtainCount = False
                                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1
                                        End If
                                    End If
                                Else
                                    If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) Then
                                        'illegal Date
                                        If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                            'mCntrSchdSpotChk = False
                                            mObtainCount = False
                                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1
                                        End If
                                    End If
                                End If
                            End If
       'Doug^^^^^^^^^^^^^^^^
                        End If
                    End If
                Else
                    'Process MG that are scheduled prior to missed and are in different weeks
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        If (-tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine) And (tlSdfExt(llSdfIndex).lMdDate > 0) Then
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            ''Obtain original dates
                            'tmSmfSrchKey.lChfCode = tmClf.lChfCode
                            'tmSmfSrchKey.iLineNo = tmClf.iLine
                            ''slDate = Format$(llChkStartDate, "m/d/yy")
                            ''gPackDate slDate, ilDate0, ilDate1
                            'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                            'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                            'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            'DL: 4-20-05 use key2 instead of key0 for speed
                            tmSmfSrchKey2.lCode = tmSdf.lCode
                            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                                If tmSmf.lSdfCode = tmSdf.lCode Then
                                    ilVefCode = tmSmf.iOrigSchVef
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            Loop
                            If (tmVsf.iFSCode(ilCVsf) = ilVefCode) Then
                                If ((tlSdfExt(llSdfIndex).lMdDate >= llChkStartDate) And (tlSdfExt(llSdfIndex).lMdDate <= llChkEndDate)) Then
                                    If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                        ilSdfSpots = ilSdfSpots + 1
                                    End If
                                    ilDay = gWeekDayLong(tlSdfExt(llSdfIndex).lMdDate)
                                    If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                        If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) And (tgCffCB(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                            'illegal Date
                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                            '    mCntrSchdSpotChk = False
                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                            'End If
                                        End If
                                    Else
                                        If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) Then
                                            'illegal Date
                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                            '    mCntrSchdSpotChk = False
                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                            'End If
                                        End If
                                    End If
                                    tlSdfExt(llSdfIndex).lMdDate = -tlSdfExt(llSdfIndex).lMdDate    'Use negative to indicate missed counted
                                End If
                            End If
                        End If
                    End If
                End If
            Next llSdfIndex
            If tmVsf.iNoSpots(ilCVsf) * ilCffSpots <> ilSdfSpots Then
                'mCntrSchdSpotChk = False
                mObtainCount = False
            End If
        End If
    Next ilCVsf
End Function

'           12-28-17
'           mFilterCntTypes - filter the contract type selection
'           <input> tlChf - contract header buffer
'                   tlCntTypes - array of contract types incl/excl
'           <output> ilFoundSpot - set to false if exclude the type
Public Sub mFilterCntTypes(tlChf As CHF, tlCntTypes As CNTTYPES, ilFoundSpot)
    If tlChf.sStatus = "H" Then
         If Not tlCntTypes.iHold Then
             ilFoundSpot = False
         End If
     ElseIf tlChf.sStatus = "O" Then
         If Not tlCntTypes.iOrder Then
             ilFoundSpot = False
         End If
     End If

     If tlChf.sType = "C" Then          '3-16-10 wrong flag previously tested (S--->C for standard)
         If Not tlCntTypes.iStandard Then       'include Standard types?
             ilFoundSpot = False
         End If

     ElseIf tlChf.sType = "V" Then
         If Not tlCntTypes.iReserv Then      'include reservations ?
             ilFoundSpot = False
         End If

     ElseIf tlChf.sType = "R" Then
         If Not tlCntTypes.iDR Then       'include DR?
             ilFoundSpot = False
         End If
    ElseIf tlChf.sType = "S" Then
         If Not tlCntTypes.iPSA Then      'include PSA ?
             ilFoundSpot = False
         End If

     ElseIf tlChf.sType = "M" Then
         If Not tlCntTypes.iPromo Then       'include Promo?
             ilFoundSpot = False
         End If
         
    '9-10-14 Remnant and PI filters were not tested
    ElseIf tlChf.sType = "T" Then
         If Not tlCntTypes.iRemnant Then       'include Remnant?
             ilFoundSpot = False
         End If
    ElseIf tlChf.sType = "Q" Then
         If Not tlCntTypes.iPI Then       'include PI?
             ilFoundSpot = False
         End If
    End If
End Sub

Sub mBuildMonthlyDates(slStart As String, ilCalType As Integer, ilPeriods As Integer)
    ReDim lmStartDates(1) As Long
    Dim illoop As Integer
    Dim slDate As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
'Debug.Print " - mBuildMonthlyDates for " & slStart & " * " & ilPeriods
    ilMonth = Month(slStart)
    ilYear = Year(slStart)
'Debug.Print " - ";
    For illoop = 1 To ilPeriods
        If ilCalType = 4 Then 'Std Cal
            slDate = DateSerial(ilYear, ilMonth + (illoop - 1), 15)
            slDate = gObtainStartStd(slDate)
        End If
        
        If ilCalType = 1 Then 'monthly Cal
            slDate = DateSerial(ilYear, ilMonth + (illoop - 1), 1)
        End If
'Debug.Print slDate & ",";
        lmStartDates(illoop) = gDateValue(slDate)
        ReDim Preserve lmStartDates(UBound(lmStartDates) + 1) As Long
    Next illoop
'Debug.Print ""
End Sub

Function mNumberOfDaysRunningInMonth(llMonthStartDate As Long, llMonthEndDate As Long, llLineStartDate As Long, llLineEndDate As Long, Optional slReportStartDate As String, Optional slReportEndDate As String)
    Dim sStartDate As String
    Dim sEndDate As String
    
    sStartDate = IIF(llLineStartDate >= llMonthStartDate, Format(llLineStartDate, "ddddd"), Format(llMonthStartDate, "ddddd"))
    If slReportStartDate <> "" Then
        If DateValue(slReportStartDate) > DateValue(sStartDate) Then sStartDate = slReportStartDate
    End If
    
    sEndDate = IIF(llLineEndDate <= llMonthEndDate, Format(llLineEndDate, "ddddd"), Format(llMonthEndDate, "ddddd"))
    If slReportEndDate <> "" Then
        If DateValue(slReportEndDate) < DateValue(sEndDate) Then sEndDate = slReportEndDate
    End If
    
    mNumberOfDaysRunningInMonth = DateDiff("d", DateValue(sStartDate), DateValue(sEndDate)) + 1
    If mNumberOfDaysRunningInMonth < 0 Then mNumberOfDaysRunningInMonth = 0
    'Debug.Print " - mNumberOfDaysRunningInMonth:" & mNumberOfDaysRunningInMonth & " - Month:" & Format(llMonthStartDate, "ddddd") & "-" & Format(llMonthEndDate, "ddddd") & ", Line Dates::" & Format(llLineStartDate, "ddddd") & "-" & Format(llLineEndDate, "ddddd") & ", Report Dates::" & slReportStartDate & "-" & slReportEndDate
End Function

Sub mWriteExportHeader(ilListIndex As Integer, hlExport As Integer)
    Dim slOutput As String
    
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    'Digital Line calculation note: optional column that shows information about how the digital line revenue was calculated
    If ilListIndex = CNT_SPTCOMBO Then
        slOutput = ""
        '1. ContractNo
        slOutput = slOutput & MakeCSVField("ContractNo", True, True)
        '2. ExtContractNo
        slOutput = slOutput & MakeCSVField("ExtContractNo", True, True)
        '3. LineType
        slOutput = slOutput & MakeCSVField("Conventional/Hidden", True, True)
        '4. ContractType
        slOutput = slOutput & MakeCSVField("ContractType", True, True)
        '5. Agency
        slOutput = slOutput & MakeCSVField("Agency", True, True)
        '6. Advertiser
        slOutput = slOutput & MakeCSVField("Advertiser", True, True)
        '7. Product
        slOutput = slOutput & MakeCSVField("Product", True, True)
        '8. Line
        slOutput = slOutput & MakeCSVField("LineNumber", True, True)
        '9. Vehicle
        slOutput = slOutput & MakeCSVField("Vehicle", True, True)
        '10. Day
        slOutput = slOutput & MakeCSVField("Day", True, True)
        '11. SpotAirDate
        slOutput = slOutput & MakeCSVField("SpotAirDate", True, True)
        '12. SpotAirTime
        slOutput = slOutput & MakeCSVField("SpotAirTime", True, True)
        '13. SpotAudioType
        slOutput = slOutput & MakeCSVField("SpotAudioType", True, True)
        '14. ISCIcode
        slOutput = slOutput & MakeCSVField("ISCIcode", True, True)
        '15. Len
        slOutput = slOutput & MakeCSVField("Len", True, True)
        '16. DigitalLineStartDate
        slOutput = slOutput & MakeCSVField("DigitalLineStartDate", True, True)
        '17. DigitalLineEndDate
        slOutput = slOutput & MakeCSVField("DigitalLineEndDate", True, True)
        'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 - Issue 1, Net/Gross Order
        '18. Gross
        slOutput = slOutput & MakeCSVField("Gross", True, True)
        '19. Net
        slOutput = slOutput & MakeCSVField("Net", True, True)
        '20. Spot status: the status of the spot (scheduled, missed, hidden, cancelled)
        slOutput = slOutput & MakeCSVField("SpotStatus", True, True)
        '21. Spot price type: the spot price type from the contract line (charge, N/C, MG, ADU, Recap, Spinoff, Bonus, Package.)
        slOutput = slOutput & MakeCSVField("SpotPriceType", True, True)
        '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
        slOutput = slOutput & MakeCSVField("DaypartLoc", True, True)
        '23. Line comment: the line comment from the spot line or digital line
        slOutput = slOutput & MakeCSVField("LineComment", True, RptSelCb.ckcSelDigitalComments.Value = vbChecked)
        If RptSelCb.ckcSelDigitalComments.Value = vbChecked Then
            '24. Digital Formula comment: the line comment from the spot line or digital line
            slOutput = slOutput & MakeCSVField("FormulaComment", True, False)
        End If
    End If
    
    'Debug.Print slOutput
    Print #hlExport, slOutput
End Sub

Sub mWriteExportFile(ilListIndex As Integer, hlExport As Integer, hlChf As Integer, hlAgf As Integer, hlAdf As Integer, hlClf As Integer, hlRdf As Integer, tlExp As EXPWOINVLN)
    Dim slOutput As String
    Dim ilRet As Integer
    Dim sAgencyName As String
    Dim ilAgyCommPct As Integer
    Dim ilInx As Integer
    Dim sAdvertiserName As String
    Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
    Dim lCxfCode As Long
    
    ilAgyCommPct = 10000 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 10)
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    'Digital Line calculation note: optional column that shows information about how the digital line revenue was calculated
    If ilListIndex = CNT_SPTCOMBO Then
        If tlExp.lContract_Number = 0 And tlExp.lChfCode <> 0 Then
            'Lookup contract
            tmChfSrchKey.lCode = tlExp.lChfCode
            ilRet = btrGetEqual(hlChf, tgChfCB, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If
            
            '1. ContractNo
            tlExp.lContract_Number = tgChfCB.lCntrNo
            '2. ExtContractNo
            tlExp.lExternal_Version_Number = tgChfCB.lExtCntrNo
            '4. ContractType - the contract type from the contract header (standard, remnant, reservation, etc.)
            Select Case tgChfCB.sType
                Case "C": tlExp.sContract_Type = "Standard"
                Case "V": tlExp.sContract_Type = "Reservation"
                Case "T": tlExp.sContract_Type = "Remnant"
                Case "R": tlExp.sContract_Type = "Direct Response"
                Case "Q": tlExp.sContract_Type = "Per Inquiry" 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 7)
                Case "S": tlExp.sContract_Type = "PSA"
                Case "M": tlExp.sContract_Type = "Promo"
                Case Else: tlExp.sContract_Type = Trim(tgChfCB.sType)
            End Select
            
            '5. Agency
            tlExp.iAgency_ID = tgChfCB.iAgfCode
            '6. Advertiser
            tlExp.iAdvertiser_ID = tgChfCB.iAdfCode
            '7. Product
            tlExp.sProduct_Name = Trim(tgChfCB.sProduct)
        End If
        '-------------------
        'Lookup Agy
        If tlExp.iAgency_ID <> 0 Then
            ilAgyCommPct = 10000                    'assume gross requested
            ilInx = gBinarySearchAgf(tlExp.iAgency_ID)
            If ilInx >= 0 Then
                sAgencyName = Trim(tgCommAgf(ilInx).sName)
                ilAgyCommPct = tgCommAgf(ilInx).iCommPct
                'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 10)
                'If ilAgyCommPct <> 0 Then
                ilAgyCommPct = 10000 - ilAgyCommPct
                'End If
            End If
        End If
        '-------------------
        'Lookup Adv
        If tlExp.iAdvertiser_ID <> 0 Then
            ilInx = gBinarySearchAdf(tlExp.iAdvertiser_ID)
            If ilInx >= 0 Then
                sAdvertiserName = Trim(tgCommAdf(ilInx).sName)
            End If
        End If

        'Lookup CLF/PCF Line
        If tlExp.sLine_Type <> "Digital" Then
            '-------------------
            'AIRTIME SPOT
            If smLastCLF <> tlExp.lChfCode & "," & tlExp.iPkLineNo Then
                tmClfSrchKey1.lChfCode = tlExp.lChfCode
                tmClfSrchKey1.iVefCode = tlExp.iVehicle_ID
                ilRet = btrGetEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Do While tmClf.lChfCode = tlExp.lChfCode And tmClf.iLine = tlExp.iPkLineNo And ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid looping
                    '3. LineType - "conventional" (not in a package) or "hidden" (in a package)
                    If tmClf.sType = "S" Then
                        tlExp.sLine_Type = "Conventional"
                    End If
                    If tmClf.sType = "O" Or tmClf.sType = "A" Or tmClf.sType = "H" Then
                        tlExp.sLine_Type = "Hidden"
                    End If
                    
                    '13. SpotAudioType
                    Select Case tmClf.sLiveCopy
                        Case "R": tlExp.sSpotAudioType = "RC"
                        Case "M": tlExp.sSpotAudioType = "LP"
                        Case "S": tlExp.sSpotAudioType = "RP"
                        Case "P": tlExp.sSpotAudioType = "PC"
                        Case "Q": tlExp.sSpotAudioType = "PP"
                        Case "L": tlExp.sSpotAudioType = "LC"
                        Case Else: tlExp.sSpotAudioType = tmClf.sLiveCopy
                    End Select
                    
                    '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
                    'get the daypart record
                    tmRdfSrchKey.iCode = tmClf.iRdfCode
                    ilRet = btrGetEqual(hlRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                    If ilRet <> BTRV_ERR_NONE Then
                        tlExp.sDaypart = "Missing DP"
                    Else
                        tlExp.sDaypart = Trim$(tmRdf.sName)
                    End If
                    
                    '23. Line Comment
                    If tmClf.lCxfCode <> 0 Then
                        smLineComment = mGetcxfComment(tmClf.lCxfCode)
                    Else
                        smLineComment = ""
                    End If
                    
                    Exit Do
                    'ilRet = btrGetNext(hlClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                
                smLastCLF = tlExp.lChfCode & "," & tlExp.iPkLineNo
            End If
        Else
            '-------------------
            'DIGITAL LINE
            '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
            'get the daypart record
            tlExp.sLine_Type = "Conventional"
            
            If Val(tlExp.sDaypart) <> 0 Then
                tmRdfSrchKey.iCode = Val(tlExp.sDaypart)
                ilRet = btrGetEqual(hlRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                If ilRet <> BTRV_ERR_NONE Then
                    tlExp.sDaypart = "Missing DP"
                Else
                    tlExp.sDaypart = Trim$(tmRdf.sName)
                End If
            End If
            '-------------------
            '23. Line Comment
            If tlExp.lLineCxfComment <> 0 Then
                smLineComment = mGetcxfComment(tlExp.lLineCxfComment)
            Else
                smLineComment = ""
            End If
            
            tlExp.lLineCxfComment = 0
        End If
        '-------------------
        'Lookup Vehicle
        If tlExp.iVehicle_ID <> 0 Then
            ilInx = gBinarySearchVef(tlExp.iVehicle_ID)
            If ilInx >= 0 Then
                tlExp.sVehicle_Name = Trim(tgMVef(ilInx).sName)
            End If
        End If
        
        '------------------------------------------------------------
        'Generate output string
        slOutput = ""
        '1. ContractNo
        slOutput = slOutput & MakeCSVField(tlExp.lContract_Number, False, True)
        '2. ExtContractNo
        slOutput = slOutput & MakeCSVField(tlExp.lExternal_Version_Number, False, True)
        '3. LineType - "conventional" (not in a package) or "hidden" (in a package)
        slOutput = slOutput & MakeCSVField(tlExp.sLine_Type, True, True)
        '4. ContractType
        slOutput = slOutput & MakeCSVField(tlExp.sContract_Type, True, True)
        '5. Agency
        slOutput = slOutput & MakeCSVField(Trim(sAgencyName), True, True)
        '6. Advertiser
        slOutput = slOutput & MakeCSVField(Trim(sAdvertiserName), True, True)
        '7. Product
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sProduct_Name), True, True)
        '8. Line #
        slOutput = slOutput & MakeCSVField(tlExp.iPkLineNo, False, True)
        '9. Vehicle
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sVehicle_Name), True, True)
        '10. Day
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sDay), True, True)
        '11. SpotAirDate
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sSpotAirDate), True, True)
        '12. SpotAirTime
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sSpotAirTime), True, True)
        '13. SpotAudioType
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sSpotAudioType), True, True)
        '14. ISCIcode
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sISCIcode), True, True)
        '15. Len
        slOutput = slOutput & MakeCSVField(Trim(tlExp.iSpot_Length), False, True)
        '16. DigitalLineStartDate
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sLine_Start_Date), True, True)
        '17. DigitalLineEndDate
        slOutput = slOutput & MakeCSVField(Trim(tlExp.sLine_End_Date), True, True)
        'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 - Issue 1, Net/Gross Order
        '18. Gross
        slOutput = slOutput & MakeCSVField(tlExp.dTotal_Gross / 100, False, True)
        '19. Net
        slOutput = slOutput & MakeCSVField(Format((tlExp.dTotal_Gross * (ilAgyCommPct / 10000)) / 100, "0.00"), False, True) 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 9)
        '20. Spot status: the status of the spot (scheduled, missed, hidden, cancelled)
        slOutput = slOutput & MakeCSVField(tlExp.sStatus, True, True)
        '21. Spot price type: the spot price type from the contract line (charge, N/C, MG, ADU, Recap, Spinoff, Bonus, Package.)
        slOutput = slOutput & MakeCSVField(tlExp.sPrice_Type, True, True)
        '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
        slOutput = slOutput & MakeCSVField(tlExp.sDaypart, True, True)
        '23. Line comment: the line comment from the spot line or digital line
        slOutput = slOutput & MakeCSVField(smLineComment, True, RptSelCb.ckcSelDigitalComments.Value = vbChecked)
        If RptSelCb.ckcSelDigitalComments.Value = vbChecked Then
            '24. Digital Line Comment
            slOutput = slOutput & MakeCSVField(tlExp.sFormulaComment, True, False)
        End If
    End If
    
    'Debug.Print slOutput
    Print #hlExport, slOutput
End Sub

Sub mWriteExportFileComment(ilListIndex As Integer, hlExport As Integer, slComment As String)
    Dim slOutput As String
    Dim ilRet As Integer
    Dim sAgencyName As String
    Dim ilAgyCommPct As Integer
    Dim ilInx As Integer
    Dim sAdvertiserName As String
    Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
    Dim lCxfCode As Long
    
    ilAgyCommPct = 10000 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 10)
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    'Digital Line calculation note: optional column that shows information about how the digital line revenue was calculated
    If ilListIndex = CNT_SPTCOMBO Then
        '------------------------------------------------------------
        'Generate output string
        slOutput = ""
        '1. ContractNo
        slOutput = slOutput & MakeCSVField("", False, True)
        '2. ExtContractNo
        slOutput = slOutput & MakeCSVField("", False, True)
        '3. LineType - "conventional" (not in a package) or "hidden" (in a package)
        slOutput = slOutput & MakeCSVField("", True, True)
        '4. ContractType
        slOutput = slOutput & MakeCSVField("", True, True)
        '5. Agency
        slOutput = slOutput & MakeCSVField("", True, True)
        '6. Advertiser
        slOutput = slOutput & MakeCSVField("", True, True)
        '7. Product
        slOutput = slOutput & MakeCSVField("", True, True)
        '8. Line #
        slOutput = slOutput & MakeCSVField("", False, True)
        '9. Vehicle
        slOutput = slOutput & MakeCSVField("", True, True)
        '10. Day
        slOutput = slOutput & MakeCSVField("", True, True)
        '11. SpotAirDate
        slOutput = slOutput & MakeCSVField("", True, True)
        '12. SpotAirTime
        slOutput = slOutput & MakeCSVField("", True, True)
        '13. SpotAudioType
        slOutput = slOutput & MakeCSVField("", True, True)
        '14. ISCIcode
        slOutput = slOutput & MakeCSVField("", True, True)
        '15. Len
        slOutput = slOutput & MakeCSVField("", False, True)
        '16. DigitalLineStartDate
        slOutput = slOutput & MakeCSVField("", True, True)
        '17. DigitalLineEndDate
        slOutput = slOutput & MakeCSVField("", True, True)
        '18. Gross
        slOutput = slOutput & MakeCSVField("", False, True)
        '19. Net
        slOutput = slOutput & MakeCSVField("", False, True)
        '20. Spot status: the status of the spot (scheduled, missed, hidden, cancelled)
        slOutput = slOutput & MakeCSVField("", True, True)
        '21. Spot price type: the spot price type from the contract line (charge, N/C, MG, ADU, Recap, Spinoff, Bonus, Package.)
        slOutput = slOutput & MakeCSVField("", True, True)
        '22. Ordered daypart/ad location: the daypart name for spot records, the ad location name for digital line records (from the contract line)
        slOutput = slOutput & MakeCSVField("", True, True)
        '23. Line comment: the line comment from the spot line or digital line
        slOutput = slOutput & MakeCSVField("", True, RptSelCb.ckcSelDigitalComments.Value = vbChecked)
        If RptSelCb.ckcSelDigitalComments.Value = vbChecked Then
            '24. Digital Line Comment
            slOutput = slOutput & MakeCSVField(slComment, True, False)
        End If
    End If
    
    'Debug.Print slOutput
    Print #hlExport, slOutput
End Sub


Function MakeCSVField(sString, blQuoted As Boolean, blIncludeComma As Boolean) As String
    Dim slOutput As String
    If Mid(sString, 1, 1) = "+" Then sString = " " & sString
    If Mid(sString, 1, 1) = "-" Then sString = " " & sString
    If Mid(sString, 1, 1) = "=" Then sString = " " & sString
    
    If blQuoted = True Then slOutput = """" 'Add a Quote
    slOutput = slOutput & Replace(Replace(sString, ",", "_"), """", "'") 'Strip Commas / replace " with '
    If blQuoted = True Then slOutput = slOutput & """"  'Add a Quote
    If blIncludeComma Then slOutput = slOutput & "," 'Add a Comma
    MakeCSVField = slOutput
End Function

Function mGetcxfComment(llCxfCode As Long) As String
    Dim ilRet As Integer
    mGetcxfComment = ""
    tmCxfSrchKey.lCode = llCxfCode
    If tmCxfSrchKey.lCode <> 0 Then
        imCxfRecLen = Len(tmCxf)
        ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            mGetcxfComment = gStripChr0(Left$(tmCxf.sComment, 255))
            If Len(gStripChr0(tmCxf.sComment)) > 255 Then
                mGetcxfComment = mGetcxfComment & "..."
            End If
        End If
    End If
End Function
