Attribute VB_Name = "RPTGENCT"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptgenct.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
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
'In RptGen
'Type TYPESORT
'    sKey As String * 100
'    lRecPos As Long
'End Type
'In RptGen
'Type SPOTTYPESORT
'    sKey As String * 80 'Office Advertiser Contract
'    iVefcode As Integer 'line airing vehicle
'    sCostType As String * 12    'string of spot type (0,00, bonus, adu, recapturable, etc)
'    tSdf As SDF
'End Type
'Type COPYROTNO
'    iRotNo As Integer
'    sZone As String * 3
'End Type
'Type COPTSORTCT
'    sKey As String * 80 'Agency, City ID
'    iCopyStatus As Integer '0=No Copy; 1=Assigned; 2=Copy but not assigned; 3= Supersede; 4=Zone missing
'    tSdf As SDF
'    sVehName As String * 20
'End Type
'In RptGenCB
'Type SPOTSALE
'    sKey As String * 100    'Vehicle|sSofName|AdvtName|Date or 99999 if total
'    iVefcode As Integer
'    sVehName As String * 20
'    sSofName As String * 20
'    sAdvtName As String * 30
'    lCntrNo As Long
'    lDate As Long
'    sDate As String * 8
'    iCNoSpots As Long            'chged to long 5-12-99 from integer
'    sCGross As String * 12
'    sCCommission As String * 12
'    sCNet As String * 12
'    iTNoSpots As Integer
'    sTGross As String * 12
'    sTCommission As String * 12
'    sTNet As String * 12
'End Type
'Type DALLASFDSORT
'    sKey As String * 30
'    sRecord As String * 104
'End Type
'In RptGenCB
'Type VEHICLELLD
'    iVefcode As Integer             'vehicle code
'    iLLD(0 To 1) As Integer         'vehicles last log date
'End Type
Dim tmSort() As TYPESORT
'Dim tmCopyCntr() As COPYCNTRSORT
'Dim tmCopy() As COPTSORTCT
'Dim tmDallasFdSort() As DALLASFDSORT
Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract line flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim hmVsf As Integer            'Vehicle combo file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer            'Agency name file handle
Dim tmAgf As AGF                'AGF record image
Dim imAgfRecLen As Integer        'AGF record length
Dim hmRcf As Integer            'Rate card file handle
Dim tmRcf As RCF                'RCF record image
Dim imRcfRecLen As Integer        'RCF record length
Dim hmRdf As Integer            'Rate card program/time file handle
Dim tmRdf As RDF                'RDF record image
Dim tmRdfSrchKey As INTKEY0     'RDF record image
Dim imRdfRecLen As Integer      'RdF record length
Dim hmStf As Integer            'MG and outside Times file handle
Dim tmStf As STF                'STF record image
Dim imStfRecLen As Integer      'STF record length
                                '5-1-12 All occurences of TmASTf have been changed to the new field stored in the arry along with the entire stf record
'Dim tmAStf() As STF            '5-1-12 fix overflow error due to large number stored in line field
Dim tmAStf() As STFPLUS
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF
'Short Title
Dim hmSif As Integer        'Short Title file handle
Dim tmSif As SIF            'SIF record image
Dim imSifRecLen As Integer     'SIF record length
'Copy rotation
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
Dim imCrfRecLen As Integer     'CRF record length
'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
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
Dim imSlfRecLen As Integer        'SLF record length
Dim hmMnf As Integer            'MultiName file handle
Dim tmMnf As MNF                'MNF record image
Dim imMnfRecLen As Integer        'MNF record length
Dim hmSof As Integer            'Sales Office file handle
Dim tmSof As SOF                'SOF record image
Dim imSofRecLen As Integer        'SOF record length
Dim hmUrf As Integer            'User file handle
Dim tmUrf As URF                'URF record image
Dim tmUrfSrchKey As INTKEY0            'URF record image
Dim imUrfRecLen As Integer        'URF record length
Dim hmCxf As Integer            'Contract Header Comment file handle
Dim tmCxf As CXF                'CXF record image
Dim tmCxfSrchKey As LONGKEY0            'CXF record image
Dim imCxfRecLen As Integer        'CXF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmCbf As CBF                   'Portrait contract & Contract History
Dim hmCbf As Integer
Dim imCbfRecLen As Integer        'CBF record length

Dim tmFsf As FSF
Dim hmFsf As Integer
Dim tmFSFSrchKey As LONGKEY0       'Gen date and time
Dim imFsfRecLen As Integer        'Gen record length

Dim tmPrf As PRF
Dim hmPrf As Integer
Dim tmPrfSrchKey As LONGKEY0       'Gen date and time
Dim imPrfRecLen As Integer        'Gen record length

Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim lmMatchingCnts() As Long        'array of contr codes selected for contr rept
Dim lmSingleCntr As Long            '12-1-00

Type STFPLUS
    StfRec As STF
    llIndex As Long
End Type
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmmlChgRpt                     *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate Spot Tracking for     *
'*                     affiliate in commercial         *
'*                     change format                   *
'*      DH 10-25-00 converted to crystal
'*                                                     *
'*******************************************************
Sub gCmmlChgRptCt()
    Dim ilErrorFlag As Integer
    Dim llRecNo As Long
    Dim ilDBRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slCreateStartDate As String
    Dim slCreateEndDate As String
    Dim slAirStartDate As String
    Dim slAirEndDate As String
    Dim llNoRecsToProc As Long
    'Dim ilSortIndex As Integer
    Dim llSortIndex As Long
    Dim ilAllSpots As Integer
    'Dim ilLoop As Integer
    Dim llLoop As Long                  '8-5-09
    'Dim ilStartIndex As Integer
    'Dim ilEndIndex As Integer
    Dim llStartIndex As Long            '8-5-09
    Dim llEndIndex As Long              '8-5-09
    Dim llLastStartIndex As Long        '8-5-09
    'Dim ilLastStartIndex As Integer
    Screen.MousePointer = vbHourglass
    slCreateStartDate = RptSelCt!CSI_CalFrom.Text   'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom.Text   'Start date
    slCreateEndDate = RptSelCt!CSI_CalTo.Text       'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom1.Text    'End date
    slAirStartDate = RptSelCt!CSI_From1.Text        'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo.Text   'Start date
    slAirEndDate = RptSelCt!CSI_To1.Text            'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo1.Text   'End date
    ilAllSpots = 0
    If (RptSelCt!ckcSelC3(0).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 1
    End If
    If (RptSelCt!ckcSelC3(1).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 2
    End If
    If (RptSelCt!ckcSelC3(2).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 4
    End If
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tgChfCT)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVlfRecLen = Len(tmVlf)
    hmStf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imStfRecLen = Len(tmStf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCrfRecLen = Len(tmCrf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSifRecLen = Len(tmSif)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
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
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    tmAdf.iCode = 0
    tmVef.iCode = 0
    'define variables for the load check (yes, L&L now checks the definition file!)
    mObtainStf ilAllSpots, slCreateStartDate, slCreateEndDate, slAirStartDate, slAirEndDate
    'Remove pairs
    llNoRecsToProc = UBound(tmSort) 'tmSort(0 To x) so UBound is number of records
    llStartIndex = LBound(tmSort)
    llLastStartIndex = llStartIndex
        llRecNo = 1
        ilErrorFlag = 0
        'Get first record
        'Determine Start/End index-matching Date/Time
        If Not mDetermineStfIndex(llStartIndex, llEndIndex) Then
            ilDBRet = 1
        End If
        Do While ilDBRet = BTRV_ERR_NONE
            For llLoop = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                'llSortIndex = tmAStf(llLoop).iLineNo
                llSortIndex = tmAStf(llLoop).llIndex            '5-1-12
                'If tmAStf(llLoop).lChfCode <> tmChf.lCode Then
                If tmAStf(llLoop).StfRec.lChfCode <> tmChf.lCode Then
                    'tmChfSrchKey.lCode = tmAStf(llLoop).lChfCode
                    tmChfSrchKey.lCode = tmAStf(llLoop).StfRec.lChfCode '5-1-12
                    ilDBRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                'If tmAStf(llLoop).lSdfCode <= 0 Then    'ABC- old system
                If tmAStf(llLoop).StfRec.lSdfCode <= 0 Then 'ABC - old system
                    'tmSdf.iVefCode = tmAStf(llLoop).iVefCode
                    tmSdf.iVefCode = tmAStf(llLoop).StfRec.iVefCode
                    'tmSdf.lChfCode = tmAStf(llLoop).lChfCode
                    tmSdf.lChfCode = tmAStf(llLoop).StfRec.lChfCode
                    'tmSdf.iLineNo = tmAStf(llLoop).iLineNo
                    tmSdf.iLineNo = tmAStf(llLoop).StfRec.iLineNo
                    tmSdf.sPtType = "0"
                    tmSdf.iRotNo = 0
                Else
                    'tmSdfSrchKey3.lCode = tmAStf(llLoop).lSdfCode
                    tmSdfSrchKey3.lCode = tmAStf(llLoop).StfRec.lSdfCode
                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        'tmSdf.iVefCode = tmAStf(llLoop).iVefCode
                        tmSdf.iVefCode = tmAStf(llLoop).StfRec.iVefCode
                        'tmSdf.lChfCode = tmAStf(llLoop).lChfCode
                         tmSdf.lChfCode = tmAStf(llLoop).StfRec.lChfCode
                        'tmSdf.iLineNo = tmAStf(llLoop).iLineNo
                        tmSdf.iLineNo = tmAStf(llLoop).StfRec.iLineNo
                        tmSdf.sPtType = "0"
                        tmSdf.iRotNo = 0
                    End If
                End If
                If tmAdf.iCode <> tmChf.iAdfCode Then
                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                    ilDBRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If

                tmGrf.sGenDesc = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
                'tmGrf.iYear = tmAStf(llLoop).iLen
                tmGrf.iYear = tmAStf(llLoop).StfRec.iLen
                'tmGrf.iVefCode = tmAStf(llLoop).iVefCode
                tmGrf.iVefCode = tmAStf(llLoop).StfRec.iVefCode
                'tmGrf.iDate(0) = tmAStf(llLoop).iLogDate(0)
                'tmGrf.iDate(1) = tmAStf(llLoop).iLogDate(1)
                tmGrf.iDate(0) = tmAStf(llLoop).StfRec.iLogDate(0)
                tmGrf.iDate(1) = tmAStf(llLoop).StfRec.iLogDate(1)

                ilRet = gParseItem(tmSort(llSortIndex).sKey, 7, "|", slStr)
                gPackTime slStr, tmGrf.iTime(0), tmGrf.iTime(1) 'log time

                tmGrf.lChfCode = tmChf.lCntrNo
                tmGrf.iAdfCode = tmChf.iAdfCode
                'If tmAStf(llLoop).sAction = "A" Then
                If tmAStf(llLoop).StfRec.sAction = "A" Then
                    tmGrf.iCode2 = 2    ' "Add"
                Else
                    tmGrf.iCode2 = 1    ' "Remove"
                End If

                'tmGrf.iStartDate(0) = tmAStf(llLoop).iCreateDate(0)       'Action date
                'tmGrf.iStartDate(1) = tmAStf(llLoop).iCreateDate(1)
                tmGrf.iStartDate(0) = tmAStf(llLoop).StfRec.iCreateDate(0)       'Action date
                tmGrf.iStartDate(1) = tmAStf(llLoop).StfRec.iCreateDate(1)

                'tmGrf.iMissedTime(0) = tmAStf(llLoop).iCreateTime(0)     'Action Time
                'tmGrf.iMissedTime(1) = tmAStf(llLoop).iCreateTime(1)
                tmGrf.iMissedTime(0) = tmAStf(llLoop).StfRec.iCreateTime(0)     'Action Time
                tmGrf.iMissedTime(1) = tmAStf(llLoop).StfRec.iCreateTime(1)


                'tmGrf.iGenTime(0) = igNowTime(0)
                'tmGrf.iGenTime(1) = igNowTime(1)
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmGrf.lGenTime = lgNowTime
                tmGrf.iGenDate(0) = igNowDate(0)
                tmGrf.iGenDate(1) = igNowDate(1)
                'tmGrf.lCode4 = tmAStf(llLoop).lCode
                tmGrf.lCode4 = tmAStf(llLoop).StfRec.lCode

                'grf parameters for crystal
                'tmgrf.iGenTime = generation time
                'tmgrf.igendate = generation date
                'tmgrf.lcode4 = Spot tracking record code for sorting (if same dates & times, first in is printed)
                'tmgrf.lchfcode = Contract # (not code)
                'tmgrf.ivefCode = vehicle code
                'tmgrf.sGenDesc = short title or product
                'tmgrf.iadfcode = advertiser code
                'tmgrf.iYear = spot length
                'tmgrf.iCode2: 1 = remove, 2 = add
                'tmgrf.iDate = Log Date
                'tmgrf.iSTime = log time
                'tmgrf.iStartDate = creation date
                'tmgrf.iMissedTime = creation time
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            Next llLoop
            If ilRet = 0 Then
                llRecNo = llRecNo + llEndIndex - llLastStartIndex + 1
                llStartIndex = llEndIndex + 1
                llLastStartIndex = llStartIndex
                If Not mDetermineStfIndex(llStartIndex, llEndIndex) Then
                    ilDBRet = 1
                End If
            Else
                ilErrorFlag = ilRet
                mErrMsg ilErrorFlag
            End If
        Loop

    Screen.MousePointer = vbDefault
    Erase tmAStf
    Erase tmSort
    ilRet = btrClose(hmSif)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCrf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmStf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSdf)
    btrDestroy hmSif
    btrDestroy hmAdf
    btrDestroy hmCrf
    btrDestroy hmClf
    btrDestroy hmStf
    btrDestroy hmVLF
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmVsf
    btrDestroy hmSdf
    btrDestroy hmGrf
End Sub

'********************************************************************
'*
'*      Procedure Name:gCntrRptCt
'*
'*             Created:6/16/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Generate Contract report
'
'       <input> ilHistory = true if show all versions
'                           else false for current version
'               ilpreview - True if display, false = print
'                         - if print (update printables flag)
'*
'*          11/4/98 Setting of printables was not udpated
'
'       dh 10-19-00 Convert Contract History and
'                   Portrait Contract to Crystal
'       dh 12-1-00 Single contract option
'       dh 11-11-04 CBS not showing on History
'       dh 3-29-05 show open/close bb notation on line
'*********************************************************************
Sub gCntrRptCt(ilHistory As Integer, ilPreview As Integer)
    Dim ilRet As Integer
    Dim ilRPRet As Integer              'error return from get position
    Dim slStr As String
    Dim llDate As Long
    Dim llActiveStartDate As Long
    Dim llActiveEndDate As Long
    Dim llEnterStartDate As Long
    Dim llEnterEndDate As Long
    Dim slDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilPass As Integer
    Dim ilClf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCff As Integer
    Dim ilDay As Integer
    Dim ilSpotsPerWk As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilLoop As Integer
    Dim ilNoFlights As Integer
    Dim slInvalid As String
    Dim ilSelection As Integer  '0=Contract; 1=Agency; 2=Salesperson
    Dim llOverallStartDate As Long
    Dim llOverallEndDate As Long
    Dim slHeaderComment As String
    Dim ilPrintOnly As Integer
    Dim ilUpper As Integer
    Dim llCurrentRecd As Long
    Dim llContrCode As Long
    Dim ilFoundCnt As Integer
    Dim ilGetAll As Integer             'true = get history of all lines, false = get only current
    Dim slCntrType  As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim tlChfAdvtExt() As CHFADVTEXT
    Dim ilSeq As Integer
    '********* TMCBF fields used for Crystal report *************
    'cbfGenDate - generation date
    'cbfGenTime - generation time
    'cbfChfCode - contract code
    'cbfLineNo - schedule line #
    'cbfvefCode - schedule line vehicle code
    'cbfDysTms - Daypart days & times
    'cbfLen - Spot Length
    'cbfRate - integer ($) Rate for accumulating in Crystal
    'cbfAirWks - # weeks in flight
    'cbfTotalWks - total weeks whole order
    'cbfSortField1 - Flight start & end date string
    'cbfSortField2 - Flight days of week string
    'cbfSurvey - Entered date (History only)
    'cbfIntComent -
    'cbfDemos - User name string (cant use Crystal, encrypted field)
    'cbfCurrModSpots - total spots/week
    'cbfExtra2Byte - Seq # to keep track of schedule line for grouping in Crystal
    'cbfCntRevNo - Internal or External rev #
    'cbfBuyer - $ string (ADU, $, .00, etc)
    'cbfObbLen - open BB length
    'cbfCBBLen - closed BB length

    ilUpper = 0
    Screen.MousePointer = vbHourglass
    If ilHistory Then
        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/17.2019 added CSI calendar control for date entry --> edcSelCFrom.Text   'Start date
        If slDate = "" Then
            slDate = "1/5/1970" 'Monday
        End If
        slDate = gObtainPrevMonday(slDate)
        llActiveStartDate = gDateValue(slDate)
        slDate = RptSelCt!CSI_CalTo.Text        'Date: 12/17/2019 added CSI calendar control for date entry --> edcSelCTo.Text   'End date
        If (StrComp(slDate, "TFN", 1) = 0) Or (Len(slDate) = 0) Then
            llActiveEndDate = gDateValue("12/29/2069")    'Sunday
        Else
            slDate = gObtainNextSunday(slDate)
            llActiveEndDate = gDateValue(slDate)
        End If
        slDate = "1/5/1970" 'Monday
        slDate = gObtainPrevMonday(slDate)
        llEnterStartDate = gDateValue(slDate)
        llEnterEndDate = gDateValue("12/29/2069")    'Sunday
    Else
        slDate = RptSelCt!edcSelCFrom.Text   'Start date
        If slDate = "" Then
            slDate = "1/5/1970" 'Monday
        End If
        slDate = gObtainPrevMonday(slDate)
        llActiveStartDate = gDateValue(slDate)
        slDate = RptSelCt!edcSelCFrom1.Text   'End date
        If (StrComp(slDate, "TFN", 1) = 0) Or (Len(slDate) = 0) Then
            llActiveEndDate = gDateValue("12/29/2069")    'Sunday
        Else
            slDate = gObtainNextSunday(slDate)
            llActiveEndDate = gDateValue(slDate)
        End If
        slDate = RptSelCt!edcSelCTo.Text   'Start date
        If slDate = "" Then
            slDate = "1/5/1970" 'Monday
            slDate = gObtainPrevMonday(slDate)
        End If
        llEnterStartDate = gDateValue(slDate)
        slDate = RptSelCt!edcSelCTo1.Text   'End date
        If (StrComp(slDate, "TFN", 1) = 0) Or (Len(slDate) = 0) Then
            llEnterEndDate = gDateValue("12/29/2069")    'Sunday
        Else
            'slDate = gObtainNextSunday(slDate)
            llEnterEndDate = gDateValue(slDate)
        End If
    End If
    If ilHistory Then
        ilSelection = 0
    Else
        If RptSelCt!rbcSelCSelect(1).Value Then
            ilSelection = 1
        ElseIf RptSelCt!rbcSelCSelect(2).Value Then
            ilSelection = 2
        Else
            ilSelection = 0
        End If
        If RptSelCt!ckcSelC3(0).Value = vbChecked Then        'only want printables
            ilPrintOnly = True
        Else
            ilPrintOnly = False                 'get printables and non-printables
        End If
    End If
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tgChfCT)
    ReDim tgClfCT(0 To 0) As CLFLIST
    tgClfCT(0).iStatus = -1 'Not Used
    tgClfCT(0).lRecPos = 0
    tgClfCT(0).iFirstCff = -1
    ReDim tgCffCT(0 To 0) As CFFLIST
    tgCffCT(0).iStatus = -1 'Not Used
    tgCffCT(0).lRecPos = 0
    tgCffCT(0).iNextCff = -1
    slHeaderComment = ""
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tgClfCT(0).ClfRec)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tgCffCT(0).CffRec)
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
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRcf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRcfRecLen = Len(tmRcf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCxf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
        btrDestroy hmCxf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)
    hmCbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCbf
        btrDestroy hmSof
        btrDestroy hmCxf
        btrDestroy hmUrf
        btrDestroy hmMnf
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmRcf
        btrDestroy hmRdf
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Erase tmSdfExt
        Erase tmSdfExtSort
        Erase tgClfCT
        Erase tgCffCT
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)
    tmAdf.iCode = 0
    tmVef.iCode = 0
    tmAgf.iCode = 0
    tmUrf.iCode = 0
    tmSof.iCode = 0
    tmSof.sName = ""
    If ilHistory Then
        ilGetAll = True                 'gather all line (past and current) for this contract
    Else
        ilGetAll = False                'get only current stuff
    End If

        'ReDim llPrintedCnts(1 To 1) As Long
        ReDim llPrintedCnts(0 To 0) As Long
        'ReDim lmMatchingCnts(1 To 1) As Long            'stored contr numbers to process
        ReDim lmMatchingCnts(0 To 0) As Long            'stored contr numbers to process
        lmSingleCntr = 0                        '11-27099
        'ilFoundCnt = 1                          '12-1-00
        ilFoundCnt = 0                          '12-1-00
        If RptSelCt!edcTopHowMany <> "" Then
            lmSingleCntr = CLng(RptSelCt!edcTopHowMany)                  '11-27-00
        End If
        If lmSingleCntr > 0 Then            'get the contracts audo code
            tmChfSrchKey1.lCntrNo = lmSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)

            'ReDim Preserve lmMatchingCnts(1 To ilFoundCnt) As Long
            ReDim Preserve lmMatchingCnts(0 To ilFoundCnt) As Long
            If lmSingleCntr = tmChf.lCntrNo Then
                lmMatchingCnts(ilFoundCnt) = tmChf.lCode
                ilFoundCnt = ilFoundCnt + 1
            End If
        End If

        'ilPass = 1
        ilPass = 0
        If Not RptSelCt!ckcAllAAS.Value Or lmSingleCntr > 0 Then '12-1-00
            'For ilLoop = 0 To RptSelCt!lbcSelection(10).ListCount - 1 Step 1
            If lmSingleCntr = 0 Then        '12-1-00
                For ilLoop = 0 To RptSelCt!lbcSelection(0).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(0).Selected(ilLoop) Then                        'selected element
                        'slNameCode = RptSelCt!lbcManyCntCode.List(ilLoop)                   'pick up office code
                        slNameCode = RptSelCt!lbcCntrCode.List(ilLoop)                   'pick up office code
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)      'obtain budget name for comparisons
                        'ReDim Preserve lmMatchingCnts(1 To ilPass) As Long
                        ReDim Preserve lmMatchingCnts(0 To ilPass) As Long
                        lmMatchingCnts(ilPass) = Val(slCode)
                        ilPass = ilPass + 1
                    End If
                Next ilLoop
            End If
        Else
            slCntrType = ""
            slCntrStatus = "HOGN"              'only get holds and orders
            ilHOState = 2                  'get latest orders and revisions
            slStartDate = Format$(llActiveStartDate, "m/d/yy")
            slEndDate = Format$(llActiveEndDate, "m/d/yy")
            ilRet = gObtainCntrForDate(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
            If ilRet = 0 Then
                For llCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
                    'ReDim Preserve lmMatchingCnts(1 To ilPass) As Long
                    ReDim Preserve lmMatchingCnts(0 To ilPass) As Long
                    lmMatchingCnts(ilPass) = tlChfAdvtExt(llCurrentRecd).lCode
                    ilPass = ilPass + 1
                Next llCurrentRecd
            End If
        End If
        'Do                                      'llRecNo < llRecsRemaining
        For llCurrentRecd = LBound(lmMatchingCnts) To UBound(lmMatchingCnts) Step 1                                            'loop while llCurrentRecd < llRecsRemaining
            llContrCode = lmMatchingCnts(llCurrentRecd)
            ilFoundCnt = mGetAContractCt(hmCHF, hmClf, hmCff, llContrCode, llEnterStartDate, llEnterEndDate, llActiveStartDate, llActiveEndDate, ilGetAll)
            If ilFoundCnt Then                              'get a contract and test for printables,
                Screen.MousePointer = vbHourglass

                tmChf = tgChfCT
                llPrintedCnts(ilUpper) = tmChf.lCode
                ilUpper = ilUpper + 1
                'ReDim Preserve llPrintedCnts(1 To ilUpper) As Long         'table of contracts printed
                ReDim Preserve llPrintedCnts(0 To ilUpper) As Long         'table of contracts printed
                ilSeq = 0                   'running seq # each new line for Crystal sort
                If tmChf.lCxfCode <> 0 Then
                    tmCxf.sComment = ""
                    imCxfRecLen = Len(tmCxf) '5027
                    tmCxfSrchKey.lCode = tmChf.lCxfCode
                    ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = gStripChr0(tmCxf.sComment)
                        'If (tmCxf.iStrLen > 0) And (tmCxf.sShOrder = "Y") Then
                        If (slStr <> "") And (tmCxf.sShOrder = "Y") Then
                            tmCbf.lIntComment = tmChf.lCxfCode
                            'slHeaderComment = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                        Else
                            tmCbf.lIntComment = 0
                            'slHeaderComment = ""
                        End If
                    Else
                        tmCbf.lIntComment = 0
                        'slHeaderComment = ""
                    End If
                Else
                    tmCbf.lIntComment = 0
                    'slHeaderComment = ""
                End If

                'Calculate total # of weeks for order
                gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slSFlightDate
                gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slEFlightDate
                If slSFlightDate = "" Or slEFlightDate = "" Then
                    tmCbf.iTotalWks = 0
                Else
                    tmCbf.iTotalWks = (gDateValue(slEFlightDate) - gDateValue(slSFlightDate)) \ 7 + 1
                End If
                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                    If (ilHistory) Or (Not ilHistory And tmClf.sType <> "H") Then               'ignore hidden lines on BR Narrow version, show everything on History
                        ilCff = tgClfCT(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDate tgCffCT(ilCff).CffRec.iStartDate(0), tgCffCT(ilCff).CffRec.iStartDate(1), slSFlightDate
                            gUnpackDate tgCffCT(ilCff).CffRec.iEndDate(0), tgCffCT(ilCff).CffRec.iEndDate(1), slEFlightDate
                            If gDateValue(slEFlightDate) >= gDateValue(slSFlightDate) Then
                                If llOverallStartDate = 0 Then
                                    llOverallStartDate = gDateValue(slSFlightDate)
                                    llOverallEndDate = gDateValue(slEFlightDate)
                                Else
                                    If gDateValue(slSFlightDate) < llOverallStartDate Then
                                        llOverallStartDate = gDateValue(slSFlightDate)
                                    End If
                                    If gDateValue(slEFlightDate) > llOverallEndDate Then
                                        llOverallEndDate = gDateValue(slEFlightDate)
                                    End If
                                End If
                            End If
                            ilCff = tgCffCT(ilCff).iNextCff
                        Loop
                    End If
                Next ilClf
                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                    Screen.MousePointer = vbHourglass
                    tmClf = tgClfCT(ilClf).ClfRec
                    If (ilHistory) Or (Not ilHistory And tmClf.sType <> "H") Then               'ignore hidden lines on BR Narrow version, show everything on History
                        'User (tmClf.iUrfCode)
                        If tmClf.iUrfCode <> tmUrf.iCode Then
                            tmUrfSrchKey.iCode = tmClf.iUrfCode
                            ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmUrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                'tmUrf.sRept = "Missing"
                                tmUrf.sRept = ""        'cannot decrypt the text
                            End If
                        End If
                        tmCbf.sDemos = gDecryptField(Trim$(tmUrf.sRept))

                        tmRdfSrchKey.iCode = tmClf.iRdfCode
                        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        'DoEvents
                        'Build output arrays
                        'ReDim smOrdered(1 To 12, 1 To 1) As String
                        'Line number
                        'smOrdered(1, 1) = Trim$(Str$(tmClf.iLine))
                        tmCbf.lLineNo = tmClf.iLine
                        'Version #
                        If ilHistory Then                   'history-show internal rev #
                            'smOrdered(2, 1) = Trim$(Str$(tmClf.iCntRevNo))
                            tmCbf.iCntRevNo = tmClf.iCntRevNo
                        Else                                'narrow contract - show external
                            tmCbf.iCntRevNo = tmChf.iExtRevNo
                            'smOrdered(2, 1) = Trim$(Str$(tmChf.iExtRevNo))
                        End If
                        'Vehicle Name
                        If tmClf.iVefCode <> tmVef.iCode Then
                            tmVefSrchKey.iCode = tmClf.iVefCode
                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                tmVef.sName = "Missing"
                            End If
                        End If
                        tmCbf.iVefCode = tmClf.iVefCode
                        'smOrdered(3, 1) = Trim$(tmVef.sname)
                        'Times
                        If (tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0) Then
                            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                            tmCbf.sDysTms = slStartTime & "-" & slEndTime
                            'smOrdered(4, 1) = slStartTime & "-" & slEndTime
                        Else        'if no time override, show DP name only (days are shown in separate column)
                            tmCbf.sDysTms = tmRdf.sName
                            'smOrdered(4, 1) = tmRdf.sName
                        End If
                        'Length
                        tmCbf.iLen = tmClf.iLen
                        tmCbf.iOBBLen = tmClf.iBBOpenLen            '3-29-05
                        tmCbf.iCBBLen = tmClf.iBBCloseLen
                        tmCbf.sLineType = ""
                        If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then    'use closest avail, dont know if its open or close
                            tmCbf.sLineType = "C"       'flag to use closest to show on report
                        Else                            'open & closes are defined
                           tmCbf.sLineType = "S"        'use specific avail
                        End If
                        'smOrdered(5, 1) = Trim$(Str$(tmClf.iLen))
                        'Dates, days and number of spots per wk
                        ilNoFlights = 0
                        ilCff = tgClfCT(ilClf).iFirstCff
                        Do While ilCff <> -1
                            'ilNoFlights = ilNoFlights + 1
                            'If ilNoFlights > UBound(smOrdered, 2) Then
                            '    ReDim Preserve smOrdered(1 To 12, 1 To ilNoFlights) As String
                            'End If
                            gUnpackDate tgCffCT(ilCff).CffRec.iStartDate(0), tgCffCT(ilCff).CffRec.iStartDate(1), slSFlightDate
                            gUnpackDate tgCffCT(ilCff).CffRec.iEndDate(0), tgCffCT(ilCff).CffRec.iEndDate(1), slEFlightDate
                            If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                tmCbf.sSortField1 = "CBS"  '11-11-04 Dates of flight, show CBS for Cancel before start
                                tmCbf.sSortField2 = ""  'Days of week string
                                tmCbf.lCurrModSpots = 0 '# spots/wk
                                tmCbf.sBuyer = ""       '$ string
                                tmCbf.sSurvey = ""      'date entered
                                tmCbf.lRate = 0
                                'smOrdered(6, ilNoFlights) = ""
                                'smOrdered(7, ilNoFlights) = "Cancel Before Start"
                                'smOrdered(8, ilNoFlights) = ""
                                'smOrdered(9, ilNoFlights) = ""
                                'smOrdered(10, ilNoFlights) = ""
                                'Exit Do         11-11-04 fall thru to create the record to print
                            Else
                                tmCbf.sSortField1 = slSFlightDate & "-" & slEFlightDate
                                'smOrdered(6, ilNoFlights) = slSFlightDate & "-" & slEFlightDate
                            End If
                            slStr = ""
                            If (tgCffCT(ilCff).CffRec.iSpotsWk > 0) Or (tgCffCT(ilCff).CffRec.iXSpotsWk > 0) Or (tgCffCT(ilCff).CffRec.sDyWk = "W") Then
                                slStr = gDayNames(tgCffCT(ilCff).CffRec.iDay(), tgCffCT(ilCff).CffRec.sXDay(), 2, slInvalid)
                                tmCbf.sSortField2 = slStr
                                'smOrdered(7, ilNoFlights) = slStr
                                tmCbf.lCurrModSpots = CLng(tgCffCT(ilCff).CffRec.iSpotsWk) + tgCffCT(ilCff).CffRec.iXSpotsWk
                                'smOrdered(8, ilNoFlights) = Trim$(Str$(tgCffCt(ilCff).CffRec.iSpotsWk + tgCffCt(ilCff).CffRec.iXSpotsWk))
                                If Not ilHistory Then       'portrait contract
                                    slStr = Trim$(str$((gDateValue(slEFlightDate) - gDateValue(slSFlightDate) + 7) \ 7))
                                    tmCbf.iAirWks = Val(slStr)  'total weeks in flight
                                    tmCbf.sSurvey = slStr
                                    'smOrdered(10, ilNoFlights) = slStr
                                    If (tgCffCT(ilCff).CffRec.sPriceType = "T") And (Trim$(tmCbf.sSortField1) <> "CBS") Then
                                    'If (tgCffCt(ilCff).CffRec.sPriceType = "T") And (smOrdered(6, ilNoFlights) <> "CBS") Then
                                        slStr = gLongToStrDec(tgCffCT(ilCff).CffRec.lActPrice, 2)
                                        gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, tmCbf.sBuyer
                                        'smOrdered(9, ilNoFlights) = gLongtoStrDec(tgCffCt(ilCff).cFfRec.lActPrice, 2)
                                        'slLnSpotTotal = gAddStr(slLnSpotTotal, gMulStr(slStr, smOrdered(8, ilNoFlights)))
                                        'If tmClf.sType = "E" Then           'package price equals the totals weeks price (not per spot)
                                            'slLnDollarTotal = gAddStr(slLnDollarTotal, smOrdered(9, ilNoFlights))
                                        'Else
                                        '    slLnDollarTotal = gAddStr(slLnDollarTotal, gMulStr(gMulStr(slStr, smOrdered(8, ilNoFlights)), smOrdered(9, ilNoFlights)))
                                        'End If
                                   End If
                                End If
                            Else
                                ilSpotsPerWk = 0
                                slStr = ""
                                For ilDay = 0 To 6 Step 1
                                    ilSpotsPerWk = ilSpotsPerWk + tgCffCT(ilCff).CffRec.iDay(ilDay)
                                    slStr = slStr & str$(tgCffCT(ilCff).CffRec.iDay(ilDay))
                                Next ilDay
                                tmCbf.sSortField2 = slStr
                                'smOrdered(7, ilNoFlights) = slStr
                                tmCbf.lCurrModSpots = ilSpotsPerWk
                                'smOrdered(8, ilNoFlights) = Trim$(Str$(ilSpotsPerWk))
                                If Not ilHistory Then
                                    slStr = Trim$(str$((gDateValue(slEFlightDate) - gDateValue(slSFlightDate) + 7) \ 7))
                                    tmCbf.sSurvey = slStr
                                    tmCbf.iAirWks = Val(slStr)
                                    'smOrdered(10, ilNoFlights) = slStr
                                    slStr = "0"
                                    If (tgCffCT(ilCff).CffRec.sPriceType = "T") And (Trim$(tmCbf.sSortField1) <> "CBS") Then
                                    'If (tgCffCt(ilCff).CffRec.sPriceType = "T") And (smOrdered(6, ilNoFlights) <> "CBS") Then
                                        For llDate = gDateValue(slSFlightDate) To gDateValue(slEFlightDate) Step 1
                                            ilDay = gWeekDayLong(llDate)
                                            slStr = gAddStr(slStr, Trim$(str$(tgCffCT(ilCff).CffRec.iDay(ilDay))))
                                        Next llDate
                                        slStr = gLongToStrDec(tgCffCT(ilCff).CffRec.lActPrice, 2)
                                        gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, tmCbf.sBuyer
                                        'smOrdered(9, ilNoFlights) = gLongtoStrDec(tgCffCt(ilCff).cFfRec.lActPrice, 2)
                                        'slLnDollarTotal = gAddStr(slLnDollarTotal, gMulStr(slStr, smOrdered(9, ilNoFlights)))
                                    End If
                                End If
                            End If
                            'Rate
                            tmCbf.lRate = tgCffCT(ilCff).CffRec.lActPrice   'use integer rate for accumulating in crystal
                            If gDateValue(slEFlightDate) >= gDateValue(slSFlightDate) Then
                                 Select Case tgCffCT(ilCff).CffRec.sPriceType
                                    Case "T"    'True
                                        slStr = gLongToStrDec(tgCffCT(ilCff).CffRec.lActPrice, 2)
                                        'smOrdered(9, ilNoFlights) = gLongtoStrDec(tgCffCt(ilCff).cFfRec.lActPrice, 2)
                                        'format flight $
                                        gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, tmCbf.sBuyer
                                        'gFormatStr smOrdered(9, ilNoFlights), FMTCOMMA + FMTDOLLARSIGN, 2, smOrdered(9, ilNoFlights)

                                    Case "N"    'No Charge
                                        tmCbf.sBuyer = "N/C"
                                        'smOrdered(9, ilNoFlights) = "N/C"
                                    Case "M"    'MG Line
                                        tmCbf.sBuyer = "MG"
                                        'smOrdered(9, ilNoFlights) = "MG"
                                    Case "B"    'Bonus
                                        tmCbf.sBuyer = "Bonus"
                                        'smOrdered(9, ilNoFlights) = "Bonus"
                                    Case "S"    'Spinoff
                                        tmCbf.sBuyer = "Spinoff"
                                        'smOrdered(9, ilNoFlights) = "Spinoff"
                                    Case "P"    'Package
                                        slStr = gLongToStrDec(tgCffCT(ilCff).CffRec.lActPrice, 2)
                                        'smOrdered(9, ilNoFlights) = gLongtoStrDec(tgCffCt(ilCff).cFfRec.lActPrice, 2)
                                        'format flight $
                                        gFormatStr slStr, FMTCOMMA + FMTDOLLARSIGN, 2, tmCbf.sBuyer
                                        'gFormatStr smOrdered(9, ilNoFlights), FMTCOMMA + FMTDOLLARSIGN, 2, smOrdered(9, ilNoFlights)
                                    Case "R"    'Recapturable
                                        tmCbf.sBuyer = "Recapturable"
                                        'smOrdered(9, ilNoFlights) = "Recapturable"
                                    Case "A"    'ADU
                                        tmCbf.sBuyer = "ADU"
                                        'smOrdered(9, ilNoFlights) = "ADU"
                                End Select
                            Else
                                tmCbf.sBuyer = ""
                                'smOrdered(9, ilNoFlights) = ""
                            End If
                            If ilHistory Then
                                'gUnPackDate tmClf.iEntryDate(0), tmClf.iEntryDate(1), smOrdered(10, 1)
                                gUnpackDate tmClf.iEntryDate(0), tmClf.iEntryDate(1), tmCbf.sSurvey
                            End If
                            If ilHistory Then
                                'gUnPackDate tmClf.iEntryDate(0), tmClf.iEntryDate(1), smOrdered(10, 1)
                                gUnpackDate tmClf.iEntryDate(0), tmClf.iEntryDate(1), tmCbf.sSurvey
                            End If
                            tmCbf.sStatus = tmClf.sSchStatus
                            tmCbf.lChfCode = tmChf.lCode
                            'tmCbf.iGenTime(0) = igNowTime(0)
                            'tmCbf.iGenTime(1) = igNowTime(1)
                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                            tmCbf.lGenTime = lgNowTime
                            tmCbf.iGenDate(0) = igNowDate(0)
                            tmCbf.iGenDate(1) = igNowDate(1)
                            tmCbf.iExtra2Byte = ilSeq
                            'see above description for all CBF fields updated
                            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                            ilCff = tgCffCT(ilCff).iNextCff
                        Loop

                    End If                              'tmClf.stype <> "H"
                    ilSeq = ilSeq + 1
                Next ilClf
            End If                                      'ilfoundcnt
        Next llCurrentRecd                              'for lbound(llmatchingcnt) to ubound(llmatchingcnt)

        If ilPreview = 0 Then
            'Clear Print flag- this is in the export
            'For ilLoop = 1 To ilUpper Step 1
            For ilLoop = 0 To ilUpper Step 1
                Do
                    tmChfSrchKey.lCode = llPrintedCnts(ilLoop)
                    imCHFRecLen = Len(tmChf)
                    ilRet = btrGetEqual(hmCHF, tgChfCT, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        tgChfCT.sPrint = "P"
                        ilRet = btrUpdate(hmCHF, tgChfCT, imCHFRecLen)
                    End If
                Loop While ilRPRet = BTRV_ERR_CONFLICT
            Next ilLoop
        End If
    Erase tmSdfExt
    Erase tmSdfExtSort
    Erase tgClfCT
    Erase tgCffCT
    'Erase smOrdered
    'Erase smAired
    Erase llPrintedCnts
    Erase lmMatchingCnts
    Erase tlChfAdvtExt
    Screen.MousePointer = vbDefault
    ilRet = btrClose(hmCbf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCxf)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmCbf)
    btrDestroy hmSof
    btrDestroy hmCxf
    btrDestroy hmUrf
    btrDestroy hmMnf
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmAgf
    btrDestroy hmAdf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmRcf
    btrDestroy hmRdf
End Sub




'*******************************************************
'*
'*      Procedure Name:gTrakAffRpt
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Generate Spot Tracking for
'*                     affiliate report
'*      dh 10-24-00 Convert to Crystal
'*
'*      dh 7-27-04 Include feed spots
'*      dh 8-5-09 Create prepass for the Sales Spot Tracking report
'*******************************************************
Sub gTrakAffRptCt()
    Dim ilDBRet As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slTime As String
    Dim slCreateStartDate As String
    Dim slCreateEndDate As String
    Dim slAirStartDate As String
    Dim slAirEndDate As String
    Dim llNoRecsToProc As Long
    Dim ilSortIndex As Integer
    Dim llSortIndex As Long             '8-5-09
    Dim ilAllSpots As Integer
    Dim slShortTitle As String
    Dim ilListIndex As Integer
    
    ilListIndex = RptSelCt!lbcRptType.ListIndex

    Screen.MousePointer = vbHourglass
    slCreateStartDate = RptSelCt!CSI_CalFrom.Text       'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom.Text   'Start date
    slCreateEndDate = RptSelCt!CSI_CalTo.Text           'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text   'End date
    slAirStartDate = RptSelCt!CSI_From1.Text            'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCTo.Text   'Start date
    slAirEndDate = RptSelCt!CSI_To1.Text                'Date: 12/6/2019 added CSI calendar control for date entries -->edcSelCTo1.Text   'End date
    ilAllSpots = 0
    If (RptSelCt!ckcSelC3(0).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 1
    End If
    If (RptSelCt!ckcSelC3(1).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 2
    End If
    If (RptSelCt!ckcSelC3(2).Value = vbChecked) Then
        ilAllSpots = ilAllSpots + 4
    End If
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tgChfCT)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAdf
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
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVlfRecLen = Len(tmVlf)
    hmStf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imStfRecLen = Len(tmStf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCrfRecLen = Len(tmCrf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSif
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSifRecLen = Len(tmSif)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmAdf
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
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
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
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmStf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmGrf
        btrDestroy hmVsf
        btrDestroy hmSdf
        btrDestroy hmSif
        btrDestroy hmAdf
        btrDestroy hmCrf
        btrDestroy hmClf
        btrDestroy hmStf
        btrDestroy hmVLF
        btrDestroy hmVef
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)


    tmAdf.iCode = 0
    tmVef.iCode = 0

    mObtainStf ilAllSpots, slCreateStartDate, slCreateEndDate, slAirStartDate, slAirEndDate
    llNoRecsToProc = UBound(tmSort) 'tmSort(0 To x) so UBound is number of records
    llSortIndex = LBound(tmSort)
    If (llSortIndex <= llNoRecsToProc) And (llNoRecsToProc > 0) Then
        ilDBRet = btrGetDirect(hmStf, tmStf, imStfRecLen, tmSort(llSortIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        Do While ilDBRet = BTRV_ERR_NONE
            If tmStf.lSdfCode <= 0 Then    'ABC- old system
                tmSdf.iVefCode = tmStf.iVefCode
                tmSdf.lChfCode = tmStf.lChfCode
                tmSdf.iLineNo = tmStf.iLineNo
                tmSdf.sPtType = "0"
                tmSdf.iRotNo = 0
                tmSdf.lFsfCode = 0
            Else
                tmSdfSrchKey3.lCode = tmStf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    tmSdf.iVefCode = tmStf.iVefCode
                    tmSdf.lChfCode = tmStf.lChfCode
                    tmSdf.iLineNo = tmStf.iLineNo
                    tmSdf.sPtType = "0"
                    tmSdf.iRotNo = tmStf.iRotNo
                End If
            End If

            'If tmSdf.lChfCode = 0 Then              'feed spot when no contract code exists
            If tmStf.lChfCode = 0 Then
                tmFSFSrchKey.lCode = tmSdf.lFsfCode
                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    tmAdf.sName = ""
                End If
                slShortTitle = "(F)"           'default if no product to feed designation
                'get the product
                If tmFsf.lPrfCode > 0 Then
                    tmPrfSrchKey.lCode = tmFsf.lPrfCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slShortTitle = tmPrf.sName & Trim$(slShortTitle)        'concatenate product & (F) designation
                    End If
                End If
                tmGrf.lChfCode = 0     'contract # (not code)
                tmGrf.lLong = tmFsf.lCode
                tmGrf.iAdfCode = tmFsf.iAdfCode
            Else
                tmChfSrchKey.lCode = tmStf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    tmAdf.sName = ""
                End If
                If tmAdf.iCode <> tmChf.iAdfCode Then
                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                    ilDBRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
                slShortTitle = gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf)
                tmGrf.lChfCode = tmChf.lCntrNo      'contract # (not code)
                tmGrf.iAdfCode = tmChf.iAdfCode
                tmGrf.lLong = 0
            End If

            tmGrf.iVefCode = tmStf.iVefCode
            tmGrf.sGenDesc = Trim$(slShortTitle)
            'tmGrf.iAdfCode = tmSdf.iAdfCode
            tmGrf.iYear = tmStf.iLen
            If tmStf.sAction = "A" Then
                tmGrf.iCode2 = 2    '"Add"
            Else
                tmGrf.iCode2 = 1    '"Remove"
            End If
            tmGrf.iDate(0) = tmStf.iLogDate(0)   'log date
            tmGrf.iDate(1) = tmStf.iLogDate(1)
            ilRet = gParseItem(tmSort(llSortIndex).sKey, 7, "|", slStr)
            gPackTime slStr, tmGrf.iTime(0), tmGrf.iTime(1) 'log time

            gUnpackDate tmStf.iCreateDate(0), tmStf.iCreateDate(1), slDate
            tmGrf.iStartDate(0) = tmStf.iCreateDate(0)       'Action date
            tmGrf.iStartDate(1) = tmStf.iCreateDate(1)

            gUnpackTime tmStf.iCreateTime(0), tmStf.iCreateTime(1), "A", "2", slTime
            tmGrf.iMissedTime(0) = tmStf.iCreateTime(0)     'Action Time
            tmGrf.iMissedTime(1) = tmStf.iCreateTime(1)
            'tmGrf.iGenTime(0) = igNowTime(0)
            'tmGrf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            tmGrf.lCode4 = tmStf.lCode
            'grf parameters for crystal
            'tmgrf.iGenTime = generation time
            'tmgrf.igendate = generation date
            'tmgrf.lcode4 = Spot tracking record code for sorting (if same dates & times, first in is printed)
            'tmgrf.lchfcode = Contract # (not code)
            'tmgrf.ivefCode = vehicle code
            'tmgrf.sGenDesc = short title or product
            'tmgrf.iadfcode = advertiser code
            'tmgrf.iYear = spot length
            'tmgrf.iCode2: 1 = remove, 2 = add
            'tmgrf.iDate = Log Date
            'tmgrf.iSTime = log time
            'tmgrf.iStartDate = creation date
            'tmgrf.iMissedTime = creation time
            'tmgrf.long = FSF spot code if applicable

            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            If ilRet = 0 Then
                llSortIndex = llSortIndex + 1
                If llSortIndex <= llNoRecsToProc Then
                    ilDBRet = btrGetDirect(hmStf, tmStf, imStfRecLen, tmSort(llSortIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                Else
                    ilDBRet = 1
                End If
            End If
        Loop
    End If

    Screen.MousePointer = vbDefault
    Erase tmSort
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSif)
    ilRet = btrClose(hmCrf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmStf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmGrf)
    btrDestroy hmSdf
    btrDestroy hmSif
    btrDestroy hmCrf
    btrDestroy hmClf
    btrDestroy hmStf
    btrDestroy hmVLF
    btrDestroy hmVef
    btrDestroy hmAdf
    btrDestroy hmCHF
    btrDestroy hmGrf
End Sub
'*******************************************************
'*******************************************************
'*                                                     *
'*      Procedure Name:mDetermineStfIndex               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine next Stf index values *
'*                     Build tmAStf with records       *
'*                                                     *
'*******************************************************
Function mDetermineStfIndex(llStartIndex As Long, llEndIndex As Long) As Integer
    Dim ilRet As Integer
    'Dim ilUpper As Integer
    Dim llUpper As Long     '8-5-09
    'Dim ilLoop1 As Integer
    Dim llLoop1 As Long     '8-5-09
    'Dim ilLoop2 As Integer
    Dim llLoop2 As Long     '8-5-09
    'Dim ilIndex As Integer
    Dim llIndex As Long     '8-05-09
    Dim ilAnyRemoved As Integer
    Dim llTime0 As Long
    Dim llTime1 As Long
    Dim slVeh0 As String
    Dim slVeh1 As String
    Dim slTime As String
    Do
        'ReDim tmAStf(0 To 0) As STF
        ReDim tmAStf(0 To 0) As STFPLUS
        If llStartIndex >= UBound(tmSort) Then
            mDetermineStfIndex = False
            Exit Function
        End If
        ilRet = btrGetDirect(hmStf, tmAStf(0).StfRec, imStfRecLen, tmSort(llStartIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)     '5-1-12
        'tmAStf(0).iLineNo = llStartIndex 'Save index
        tmAStf(0).llIndex = llStartIndex 'Save index            5-1-12
        ilRet = gParseItem(tmSort(llStartIndex).sKey, 1, "|", slVeh0)
        ilRet = gParseItem(tmSort(llStartIndex).sKey, 3, "|", slTime)
        slVeh0 = Trim$(slVeh0)
        slTime = Trim$(slTime)
        llTime0 = Val(slTime)
        llUpper = 1
        'ReDim Preserve tmAStf(0 To llUpper) As STF
        ReDim Preserve tmAStf(0 To llUpper) As STFPLUS
        llEndIndex = llStartIndex + 1
        Do While llEndIndex < UBound(tmSort)
            ilRet = btrGetDirect(hmStf, tmAStf(llUpper).StfRec, imStfRecLen, tmSort(llEndIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            'tmAStf(llUpper).iLineNo = llEndIndex 'Save index
            tmAStf(llUpper).llIndex = llEndIndex 'Save index
            ilRet = gParseItem(tmSort(llEndIndex).sKey, 1, "|", slVeh1)
            ilRet = gParseItem(tmSort(llEndIndex).sKey, 3, "|", slTime)
            slVeh1 = Trim$(slVeh1)
            slTime = Trim$(slTime)
            llTime1 = Val(slTime)
            'If (StrComp(slVeh0, slVeh1, 0) <> 0) Or (tmAStf(0).iLogDate(0) <> tmAStf(llUpper).iLogDate(0)) Or (tmAStf(0).iLogDate(1) <> tmAStf(llUpper).iLogDate(1)) Or (llTime0 <> llTime1) Then
            If (StrComp(slVeh0, slVeh1, 0) <> 0) Or (tmAStf(0).StfRec.iLogDate(0) <> tmAStf(llUpper).StfRec.iLogDate(0)) Or (tmAStf(0).StfRec.iLogDate(1) <> tmAStf(llUpper).StfRec.iLogDate(1)) Or (llTime0 <> llTime1) Then
                Exit Do
            End If
            llUpper = llUpper + 1
            llEndIndex = llEndIndex + 1
            'ReDim Preserve tmAStf(0 To llUpper) As STF
            ReDim Preserve tmAStf(0 To llUpper) As STFPLUS
        Loop
        llEndIndex = llEndIndex - 1
        'Remove pars -Any matching Contract # and length and Add with remove
        ilAnyRemoved = False
        Do
            ilAnyRemoved = False
            For llLoop1 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                For llLoop2 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                    If llLoop1 <> llLoop2 Then
                        'If (tmAStf(llLoop1).lChfCode = tmAStf(llLoop2).lChfCode) And (tmAStf(llLoop1).iLen = tmAStf(llLoop2).iLen) And (tmAStf(llLoop1).sAction <> tmAStf(llLoop2).sAction) Then
                        If (tmAStf(llLoop1).StfRec.lChfCode = tmAStf(llLoop2).StfRec.lChfCode) And (tmAStf(llLoop1).StfRec.iLen = tmAStf(llLoop2).StfRec.iLen) And (tmAStf(llLoop1).StfRec.sAction <> tmAStf(llLoop2).StfRec.sAction) Then
                            ilAnyRemoved = True
                            llUpper = 0
                            For llIndex = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                                If (llIndex <> llLoop1) And (llIndex <> llLoop2) Then
                                    'tmAStf(llUpper) = tmAStf(llIndex)
                                    tmAStf(llUpper).StfRec = tmAStf(llIndex).StfRec    '5-1-12
                                    llUpper = llUpper + 1
                                End If
                            Next llIndex
                            'ReDim Preserve tmAStf(0 To llUpper) As STF
                            ReDim Preserve tmAStf(0 To llUpper) As STFPLUS
                            Exit For
                        End If
                    End If
                Next llLoop2
                If ilAnyRemoved Then
                    Exit For
                End If
            Next llLoop1
        Loop While ilAnyRemoved
        If llUpper <= 0 Then
            llStartIndex = llEndIndex + 1
        End If
    Loop While llUpper <= 0
    mDetermineStfIndex = True
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
    Dim ilLoop As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTimes As Integer
    Dim ilTBTime As Integer
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
            tgSsfSrchKey.iType = ilGameNo   '0
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
            For ilLoop = 1 To tgSsf(ilDay).iCount Step 1
               LSet tmProg = tgSsf(ilDay).tPas(ADJSSFPASBZ + ilLoop)
                If tmProg.iRecType = 1 Then 'Program subrecord
                    If (tmProg.iLtfCode = tmRdf.iLtfCode(0)) Or (tmProg.iLtfCode = tmRdf.iLtfCode(1)) Or (tmProg.iLtfCode = tmRdf.iLtfCode(1)) Then
                        gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                        llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
                        gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "1", slTime
                        llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
                        Exit Do
                    End If
                End If
            Next ilLoop
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
                    gUnpackTime tmRdf.iStartTime(0, ilTimeIndex), tmRdf.iStartTime(1, ilTimeIndex), "A", "1", slTime
                    llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
                    gUnpackTime tmRdf.iEndTime(0, ilTimeIndex), tmRdf.iEndTime(1, ilTimeIndex), "A", "1", slTime
                    llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
                    ilTBTime = ilTBTime + 1
                End If
            Next ilTimeIndex
        Else
            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slTime
            llStartTime(ilTBTime) = CLng(gTimeToCurrency(slTime, False))
            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slTime
            llEndTime(ilTBTime) = CLng(gTimeToCurrency(slTime, True))
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
    Dim ilRet As Integer
    Dim ilSSFType As Integer
    'slSsfType = "O" 'On Air
    '11/24/12
    'ilSSFType = 0 'On Air
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
'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           slProduct(O)- Product (different zones separated by Chr(10)
'                         first product obtained from tmChf if time zone
'           slZone(O)- Zones
'           slCart(O)- Carts (different zones separated by Chr(10))
'           slISCI(O)- ISCI (different zones separated by Chr(10))
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainStf                      *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Stf records to be    *
'*                     reported                        *
'*                                                     *
'*******************************************************
Sub mObtainStf(ilAllRec As Integer, slCreateStartDate As String, slCreateEndDate As String, slAirStartDate As String, slAirEndDate As String)
    ' ilAllRec(I)- Bit map Bit 1=Ready; Bit 2=Print; Bit 3=Delete
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim slDate As String
    Dim slAirDate As String
    Dim slEffDate As String
    Dim slTermDate As String
    Dim llDate As Long
    Dim llTime As Long
    Dim slTime As String
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilUpper As Integer
    Dim llUpper As Long         '8-5-09
    Dim tlVef As VEF
    Dim tlVefL As VEF
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilVlfFd As Integer
    Dim slAirTime As String
    Dim slActionDate As String
    Dim slActionTime As String
    Dim slActionType As String  '1=Removed; 2=Added
    ReDim tmSort(0 To 0) As TYPESORT
    Dim ilListIndex As Integer
    
    ilListIndex = RptSelCt!lbcRptType.ListIndex
    tlVef.iCode = 0
    llUpper = LBound(tmSort)
    btrExtClear hmStf   'Clear any previous extend operation
    ilExtLen = Len(tmStf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmStf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmStf, tmStf, imStfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmStf, llNoRec, -1, "UC", "STF", "") 'Set extract limits (all records)
        If ilAllRec <> 7 Then
            If (ilAllRec And 1) <> 0 Then
                tlCharTypeBuff.sType = "R"    'Extract all matching records
                ilOffSet = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartDate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") And ((ilAllRec And 2) = 0) And ((ilAllRec And 4) = 0) Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
            If (ilAllRec And 2) <> 0 Then
                tlCharTypeBuff.sType = "P"    'Extract all matching records
                ilOffSet = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartDate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") And ((ilAllRec And 4) = 0) Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
            If (ilAllRec And 4) <> 0 Then
                tlCharTypeBuff.sType = "D"    'Extract all matching records
                ilOffSet = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartDate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
        End If
        If slCreateStartDate <> "" Then
            gPackDate slCreateStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfCreateDate")
            If (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slCreateEndDate <> "" Then
            gPackDate slCreateEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfCreateDate")
            If (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirStartDate <> "" Then
            gPackDate slAirStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfLogDate")
            If (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirEndDate <> "" Then
            gPackDate slAirEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Stf", "StfLogDate")
            ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmStf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'Build sort record
                gUnpackDateForSort tmStf.iCreateDate(0), tmStf.iCreateDate(1), slActionDate
                gUnpackTimeLong tmStf.iCreateTime(0), tmStf.iCreateTime(1), False, llTime
                slActionTime = Trim$(str$(llTime))
                Do While Len(slActionTime) < 5
                    slActionTime = "0" & slActionTime
                Loop
                If tmStf.sAction = "R" Then
                    slActionType = "1"
                Else
                    slActionType = "2"
                End If
                'Obtain vehicle and determine if selling or conventional
                If tmVef.iCode <> tmStf.iVefCode Then
                    tmVefSrchKey.iCode = tmStf.iVefCode
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet = BTRV_ERR_NONE Then
                    If tmVef.sType = "S" And ilListIndex <> CNT_SPOTTRAK Then
                        ilVlfFd = False
                        'Map selling to airing vehicle
                        'First obtain effective date
                        gUnpackDate tmStf.iLogDate(0), tmStf.iLogDate(1), slAirDate
                        ilDay = gWeekDayStr(slAirDate)
                        If ilDay <= 4 Then
                            ilDay = 0
                        ElseIf ilDay = 5 Then
                            ilDay = 6
                        ElseIf ilDay = 6 Then
                            ilDay = 7
                        End If
                        ilEffDate0 = tmStf.iLogDate(0)
                        ilEffDate1 = tmStf.iLogDate(1)
                        tmVlfSrchKey.iSellCode = tmVef.iCode
                        tmVlfSrchKey.iSellDay = ilDay
                        tmVlfSrchKey.iEffDate(0) = ilEffDate0
                        tmVlfSrchKey.iEffDate(1) = ilEffDate1
                        tmVlfSrchKey.iSellTime(0) = 0
                        tmVlfSrchKey.iSellTime(1) = 6144  '24*256
                        tmVlfSrchKey.iSellPosNo = 32000
                        ilRet = btrGetLessOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        If (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = tmVef.iCode) And (tmVlf.iSellDay = ilDay) And (tmVlf.sStatus = "C") Then
                            ilEffDate0 = tmVlf.iEffDate(0)
                            ilEffDate1 = tmVlf.iEffDate(1)
                        Else
                            ilEffDate0 = 0
                            ilEffDate1 = 0
                        End If
                        tmVlfSrchKey.iSellCode = tmVef.iCode    'selling vehicle code number
                        tmVlfSrchKey.iSellDay = ilDay     '0=M-F, 6= Sa, 7=Su
                        tmVlfSrchKey.iEffDate(0) = ilEffDate0 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iEffDate(1) = ilEffDate1 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellTime(0) = tmStf.iLogTime(0) 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellTime(1) = tmStf.iLogTime(1) 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        tmVlfSrchKey.iSellPosNo = 0   'Unit (spot) no- currently zero
                        tmVlfSrchKey.iSellSeq = 0     'Sequence number start at 1
                        ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = tmVef.iCode) And (tmVlf.iSellDay = ilDay)
                            If tmVlf.sStatus = "C" Then
                                If (tmVlf.iSellTime(0) = tmStf.iLogTime(0)) And (tmVlf.iSellTime(1) = tmStf.iLogTime(1)) Then
                                    gUnpackDate tmVlf.iEffDate(0), tmVlf.iEffDate(1), slEffDate
                                    If gDateValue(slEffDate) <= gDateValue(slAirDate) Then
                                        gUnpackDate tmVlf.iTermDate(0), tmVlf.iTermDate(1), slTermDate
                                        If ((tmVlf.iTermDate(0) = 0) And (tmVlf.iTermDate(1) = 0)) Or (gDateValue(slAirDate) <= gDateValue(slTermDate)) Then
                                            ilVlfFd = True
                                            If tlVef.iCode <> tmVlf.iAirCode Then
                                                tmVefSrchKey.iCode = tmVlf.iAirCode
                                                ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                            Else
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                            If ilRet = BTRV_ERR_NONE Then
                                                'Create one sort record
                                                gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                                                llDate = gDateValue(slDate)
                                                gUnpackTime tmVlf.iAirTime(0), tmVlf.iAirTime(1), "A", "1", slAirTime
                                                gUnpackTimeLong tmVlf.iAirTime(0), tmVlf.iAirTime(1), False, llTime
                                                slTime = Trim$(str$(llTime))
                                                Do While Len(slTime) < 5
                                                    slTime = "0" & slTime
                                                Loop
                                                tmSort(llUpper).sKey = tlVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                                tmSort(llUpper).lRecPos = llRecPos
                                                llUpper = llUpper + 1
                                                ReDim Preserve tmSort(0 To llUpper) As TYPESORT
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If Not ilVlfFd Then
                            gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                            llDate = gDateValue(slDate)
                            gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                            gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                            slTime = Trim$(str$(llTime))
                            Do While Len(slTime) < 5
                                slTime = "0" & slTime
                            Loop
                            tmSort(llUpper).sKey = "~" & tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                            tmSort(llUpper).lRecPos = llRecPos
                            llUpper = llUpper + 1
                            ReDim Preserve tmSort(0 To llUpper) As TYPESORT
                        End If
                    Else
                        'Create one sort record
                        gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                        llDate = gDateValue(slDate)
                        gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                        gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                        slTime = Trim$(str$(llTime))
                        Do While Len(slTime) < 5
                            slTime = "0" & slTime
                        Loop
                        If tmVef.iVefCode > 0 Then
                            If tlVefL.iCode <> tmVef.iVefCode Then
                                tmVefSrchKey.iCode = tmVef.iVefCode
                                ilRet = btrGetEqual(hmVef, tlVefL, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            Else
                                ilRet = BTRV_ERR_NONE
                            End If
                            If ilRet = BTRV_ERR_NONE Then
                                tmSort(llUpper).sKey = tlVefL.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                tmSort(llUpper).lRecPos = llRecPos
                            Else
                                tmSort(llUpper).sKey = Left$(tmVef.sName, 8) & " Log Veh Missing" & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                tmSort(llUpper).lRecPos = llRecPos
                            End If
                        Else
                            tmSort(llUpper).sKey = tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                            tmSort(llUpper).lRecPos = llRecPos
                        End If
                        llUpper = llUpper + 1
                        ReDim Preserve tmSort(0 To llUpper) As TYPESORT
                    End If
                End If

                ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Loop
           Loop
        End If
    End If
    If llUpper > 0 Then
        ArraySortTyp fnAV(tmSort(), 0), llUpper, 0, LenB(tmSort(0)), 0, LenB(tmSort(0).sKey), 0 '100, 0
    End If
    Exit Sub

    ilRet = err.Number
    Resume Next
End Sub
