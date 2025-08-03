Attribute VB_Name = "RptGenIV"
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
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
'Global tmsRec As LPOPREC
'In RptGen
'Type ODFEXT
'    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sZone As String * 3
'    iEtfCode As Integer         'Event type code
'    iEnfCode As Integer         'Event name code
'    sProgCode As String * 5 'Program code #
'    iAnfCode As Integer 'Avail name code
'    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    sProduct As String * 35 'Product (either from contract or copy)
'    iMnfSubFeed As Integer
'    iBreakNo As Integer 'Reset at start of each program
'    iPositionNo As Integer
'    lCefCode As Long
'    sShortTitle As String * 15
'End Type
'In RptGen
'Type TYPESORT
'    sKey As String * 100
'    lRecPos As Long
'End Type
'In RptGen
'Type SPOTTYPESORT
'    sKey As String * 80 'Office Advertiser Contract
'    iVefCode As Integer 'line airing vehicle
'    sCostType As String * 12    'string of spot type (0,00, bonus, adu, recapturable, etc)
'    tSdf As SDF
'End Type
Type COPYCNTRSORTIV
    sKey As String * 80 'Agency, City Contrct # ID Vehicle Len
    lChfCode As Long
    iVefCode As Integer
    sVehName As String * 20
    iLen As Integer
    iNoSpots As Integer 'Number of spots that have no copy
    iNoUnAssg As Integer    'Number of spots not assigned
    iNoToReassg As Integer    'Number of spots should be reassigned
End Type
'In RptGen
'Type COPYROTNO
'    iRotNo As Integer
'    sZone As String * 3
'End Type
Type COPYSORTIV
    sKey As String * 80 'Agency, City ID
    iCopyStatus As Integer '0=No Copy; 1=Assigned; 2=Copy but not assigned; 3= Supersede; 4=Zone missing
    tSdf As SDF
    sVehName As String * 20
End Type
'In RptGenCB
'Type SPOTSALE
'    sKey As String * 100    'Vehicle|sSofName|AdvtName|Date or 99999 if total
'    iVefCode As Integer
'    sVehName As String * 20
'    sSOFName As String * 20
'    sAdvtName As String * 30
'    lCntrNo As Long
'    lDate As Long
'    sDate As String * 8
'    lCNoSpots As Integer
'    sCGross As String * 12
'    sCCommission As String * 12
'    sCNet As String * 12
'    iTNoSpots As Integer
'    sTGross As String * 12
'    sTCommission As String * 12
'    sTNet As String * 12
'End Type
'In RptGenCB
'Type CODESTNCONV
'    sName As String * 20
'    sCodeStn As String * 5
'End Type
'In RptGenCB
'Type DALLASFDSORT
'    sKey As String * 30
'    sRecord As String * 104
'End Type
'In RptGenCB
'Type VEHICLELLD
'    iVefCode As Integer             'vehicle code
'    iLLD(0 To 1) As Integer         'vehicles last log date
'End Type
Dim tmSort() As TYPESORT
Dim tmPLSdf() As SPOTTYPESORT
Dim tmSpotSOF() As SPOTTYPESORT
Dim tmCopyCntr() As COPYCNTRSORTIV
Dim tmCopy() As COPYSORTIV
Dim tmSelAdvt() As Integer
Dim tmSelChf() As Long
Dim tmSelAgf() As Integer
Dim tmSelSlf() As Integer
Dim imNoZones As Integer
Dim tmRotNo(1 To 6) As COPYROTNO
'Dim tmCodeStn() As CODESTNCONV
'Dim tmDallasFdSort() As DALLASFDSORT
Dim imSpotSaleVefCode() As Integer
Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT
Dim hmEnf As Integer            'Event name file handle
Dim tmEnf As ENF                'ENF record image
Dim tmSEnf As ENF
Dim tmEnfSrchKey As INTKEY0            'ENF record image
Dim imEnfRecLen As Integer        'ENF record length
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0            'ANF record image
Dim imAnfReclen As Integer        'ANF record length
Dim hmCef As Integer            'Event comments file handle
Dim tmCef As CEF                'CEF record image
Dim tmCefSrchKey As LONGKEY0            'CEF record image
Dim imCefRecLen As Integer        'CEF record length
Dim hmChf As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imChfRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0            'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract line flight file handle
Dim tmCffSrchKey As CFFKEY0            'CFF record image
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF
Dim hmVsf As Integer            'Vehicle combo file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmVpf As Integer            'Vehicle options file handle
Dim tmVpf As VPF                'VPF record image
Dim tmVpfSrchKey As INTKEY0     'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length
Dim hmAgf As Integer            'Agency name file handle
Dim tmAgf As AGF                'AGF record image
Dim tmAgfSrchKey As INTKEY0            'AGF record image
Dim imAgfRecLen As Integer        'AGF record length
Dim hmRcf As Integer            'Rate card file handle
Dim tmRcf As RCF                'RCF record image
Dim tmRcfSrchKey As INTKEY0            'RCF record image
Dim imRcfRecLen As Integer        'RCF record length
Dim hmRdf As Integer            'Rate card program/time file handle
Dim tmRdf As RDF                'RDF record image
Dim tmRdfSrchKey As INTKEY0     'RDF record image
Dim imRdfRecLen As Integer      'RdF record length
Dim hmSmf As Integer            'MG and outside Times file handle
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey As SMFKEY0     'SMF record image
Dim imSmfRecLen As Integer      'SMF record length
Dim hmStf As Integer            'MG and outside Times file handle
Dim tmStf As STF                'STF record image
Dim tmStfSrchKey As STFKEY0     'STF record image
Dim imStfRecLen As Integer      'STF record length
Dim tmAStf() As STF
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey As SDFKEY0     'SDF record image (key 3)
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF
'Short Title
Dim hmSif As Integer        'Short Title file handle
Dim tmSif As SIF            'SIF record image
Dim tmSifSrchKey As LONGKEY0 'SIF key record image
Dim imSifRecLen As Integer     'SIF record length
'Copy rotation
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim imCrfRecLen As Integer     'CRF record length
'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
' Copy Combo Inventory File
Dim hmCcf As Integer        'Copy Combo Inventory file handle
Dim tmCcf As CCF            'CCF record image
Dim tmCcfSrchKey As INTKEY0 'CCF key record image
Dim imCcfRecLen As Integer     'CCF record length
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfReclen As Integer     'TZF record length
'  Media code File
Dim hmMcf As Integer        'Media file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length
'  Library calendar File
Dim hmLcf As Integer        'Library calendar file handle
Dim tmLcf As LCF            'LCF record image
Dim tmLcfSrchKey As LCFKEY0 'LCF key record image
Dim imLcfRecLen As Integer     'LCF record length
Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVlf As Integer            'Vehiclee file handle
Dim tmVlf As VLF                'VEF record image
Dim tmVlfSrchKey As VLFKEY0            'VEF record image
Dim imVlfRecLen As Integer        'VEF record length
Dim hmSwf As Integer            'Spot week dump file handle
Dim tmSwf As SWF                'SWF record image
Dim tmSwfSrchKey As SWFKEY0            'SWF record image
Dim imSwfRecLen As Integer        'SWF record length
Dim hmSlf As Integer            'Salesoerson file handle
Dim tmSlf As SLF                'SLF record image
Dim tmSlfSrchKey As INTKEY0            'SLF record image
Dim imSlfReclen As Integer        'SLF record length
Dim hmMnf As Integer            'MultiName file handle
Dim tmMnf As MNF                'MNF record image
Dim tmMnfSrchKey As INTKEY0            'MNF record image
Dim imMnfRecLen As Integer        'MNF record length
Dim hmSof As Integer            'Sales Office file handle
Dim tmSof As SOF                'SOF record image
Dim tmSofSrchKey As INTKEY0            'SOF record image
Dim imSofRecLen As Integer        'SOF record length
Dim hmUrf As Integer            'User file handle
Dim tmUrf As URF                'URF record image
Dim tmUrfSrchKey As INTKEY0            'URF record image
Dim imUrfRecLen As Integer        'URF record length
Dim lmUrfRecPos As Long
Dim hmCxf As Integer            'Contract Header Comment file handle
Dim tmCxf As CXF                'CXF record image
Dim tmCxfSrchKey As LONGKEY0            'CXF record image
Dim imCxfRecLen As Integer        'CXF record length
Dim hmSpf As Integer            'Site file handle
Dim tmSpf As SPF                'SPF record image
Dim tmSpfSrchKey As INTKEY0            'SPF record image
Dim imSpfRecLen As Integer        'SPF record length
Dim imUpdateCntrNo As Integer
Dim lmSpfRecPos As Long
Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmBBSpot As BBSPOTSS
Dim tmProgTest As PROGRAMSS
Dim tmAvailTest As AVAILSS
Dim tmSpotTest As CSPOTSS
Dim tmBBSpotTest As BBSPOTSS
Dim smOrdered() As String
Dim smAired() As String
Dim lmMatchingCnts() As Long        'array of contr codes selected for contr rept
Dim tmOdf0() As ODFEXT
Dim tmOdf1() As ODFEXT
Dim tmOdf2() As ODFEXT
Dim tmOdf3() As ODFEXT
Dim tmOdf4() As ODFEXT
Dim tmOdf5() As ODFEXT
Dim tmOdf6() As ODFEXT
Dim tmVehLLD() As VEHICLELLD       'array of each vehicles last log dates
'*******************************************************
'*******************************************************
'*******************************************************

'*******************************************************
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
'*******************************************************
Function mCntrSchdSpotChk(ilDispOnly As Integer, ilClf As Integer, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'
'   ilRet = gCntrSchdSpotChk(ilClf As Integer, slStartDate)
'   Where:
'       ilDispOnly(I)- True=Discrepancy contracts only, False=All lines that fall within data span
'       ilClf(I)- Index into tgClf (which contains the line to be checked)
'                 tgCff must contain the flights for the line
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
    Dim ilSdfIndex As Integer
    Dim ilDay As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slSdfTime As String
    Dim llSdfTime As Long
    ReDim llStartTime(0 To 6) As Long
    ReDim llEndTime(0 To 6) As Long
    Dim slOrigMissedDate As String
    Dim slPrice As String
    Dim ilCVsf As Integer
    Dim ilVefFound As Integer
    Dim ilVefCode As Integer
    Dim ilTDay As Integer
    Dim ilDateFound As Integer
    Dim ilTime As Integer
    Dim ilFound As Integer
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
    ilCff = tgClf(ilClf).iFirstCff
    mCntrSchdSpotChk = True
    llLnEarliestDate = 0
    llLnLatestDate = 0
    Do While ilCff <> -1
        gUnpackDate tgCff(ilCff).CffRec.iStartDate(0), tgCff(ilCff).CffRec.iStartDate(1), slCffStartDate
        gUnpackDate tgCff(ilCff).CffRec.iEndDate(0), tgCff(ilCff).CffRec.iEndDate(1), slCffEndDate
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
            For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                If tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine Then
                    mCntrSchdSpotChk = False
                    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1  'Date invalid
                    If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(ilSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        'Obtain original dates
                        tmSmfSrchKey.lChfCode = tmSdf.lChfCode
                        tmSmfSrchKey.iLineNo = tmSdf.iLineNo
                        tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                        tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo)
                            If tmSmf.lSdfCode = tmSdf.lCode Then
                                gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                                tlSdfExt(ilSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        Loop
                    End If
                End If
            Next ilSdfIndex
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
                If (tgCff(ilCff).CffRec.iSpotsWk <> 0) Or (tgCff(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCff(ilCff).CffRec.sDyWk = "W") Then  'Weekly buy
                    ilCffSpots = tgCff(ilCff).CffRec.iSpotsWk + tgCff(ilCff).CffRec.iXSpotsWk
                    llChkStartDate = llMonDate
                    llChkEndDate = llSunDate
                    GoSub lObtainCount
                Else    'Daily buy
                    ilCffSpots = 0
                    For ilTDay = 0 To 6 Step 1
                        If (llMonDate + ilTDay >= llStartDate) And (llMonDate + ilTDay <= llEndDate) Then
                            ilCffSpots = tgCff(ilCff).CffRec.iDay(ilTDay)
                            llChkStartDate = llMonDate + ilTDay
                            llChkEndDate = llMonDate + ilTDay
                            GoSub lObtainCount
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
        ilCff = tgCff(ilCff).iNextCff
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
    For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine Then
            gUnpackDateLong tlSdfExt(ilSdfIndex).iDate(0), tlSdfExt(ilSdfIndex).iDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ilVefFound = False
                If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(ilSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    tmSmfSrchKey.lChfCode = tmClf.lChfCode
                    tmSmfSrchKey.iLineNo = tmClf.iLine
                    'slDate = Format$(llChkStartDate, "m/d/yy")
                    'gPackDate slDate, ilDate0, ilDate1
                    tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                    tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                        If tmSmf.lSdfCode = tmSdf.lCode Then
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                            tlSdfExt(ilSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
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
                            If tlSdfExt(ilSdfIndex).iVefCode = tmVsf.iFSCode(ilCVsf) Then
                                ilVefFound = True
                                Exit For
                            End If
                        End If
                    Next ilCVsf
                End If
                If Not ilVefFound Then
                'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                    If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                        tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                        mCntrSchdSpotChk = False
                    End If
                End If
                GoSub lSetStatus
                If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                    If tmSmf.lSdfCode = tmSdf.lCode Then
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                        tlSdfExt(ilSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                        'Test Date of original missed date
                        ilDateFound = False
                        ilCff = tgClf(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDateLong tgCff(ilCff).CffRec.iStartDate(0), tgCff(ilCff).CffRec.iStartDate(1), llCffStartDate
                            gUnpackDateLong tgCff(ilCff).CffRec.iEndDate(0), tgCff(ilCff).CffRec.iEndDate(1), llCffEndDate
                            If (tlSdfExt(ilSdfIndex).lMdDate >= llCffStartDate) And (tlSdfExt(ilSdfIndex).lMdDate <= llCffEndDate) Then
                                ilDay = gWeekDayLong(tlSdfExt(ilSdfIndex).lMdDate)
                                If (tgCff(ilCff).CffRec.iSpotsWk <> 0) Or (tgCff(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCff(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                    If (tgCff(ilCff).CffRec.iDay(ilDay) <> 0) Or (tgCff(ilCff).CffRec.sXDay(ilDay) = "Y") Then
                                        ilDateFound = True
                                    End If
                                Else
                                    If (tgCff(ilCff).CffRec.iDay(ilDay) <> 0) Then
                                        ilDateFound = True
                                    End If
                                End If
                                Exit Do
                            End If
                            ilCff = tgCff(ilCff).iNextCff
                        Loop
                        If Not ilDateFound Then
                            'illegal Date
                            If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                mCntrSchdSpotChk = False
                                tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                            End If
                        End If
                        mGetLegalTimes slOrigMissedDate, llStartTime(), llEndTime()
                        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slSdfTime
                        llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
                        
                        ilFound = False
                        For ilTime = 0 To 6 Step 1
                            If (llStartTime(ilTime) > 0 And llEndTime(ilTime) > 0) Then
                                'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
                                If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilTime
                        If (tlSdfExt(ilSdfIndex).sSpotType <> "X") And (Not ilFound) Then
                            mCntrSchdSpotChk = False
                            tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H2
                        End If
                    End If
                Else
                    If ilDispOnly Then
                        If (tlSdfExt(ilSdfIndex).sSchStatus <> "G") Then
                            If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                mCntrSchdSpotChk = False
                                tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1  'Date invalid
                            End If
                        End If
                    Else
                        If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                            mCntrSchdSpotChk = False
                            tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1  'Date invalid
                        End If
                    End If
                End If
                tlSdfExt(ilSdfIndex).iLineNo = -tlSdfExt(ilSdfIndex).iLineNo   'Spot not counted again
            End If
        End If
    Next ilSdfIndex
    'Remove missed date counted flag
    For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(ilSdfIndex).lMdDate < 0 Then
            tlSdfExt(ilSdfIndex).lMdDate = -tlSdfExt(ilSdfIndex).lMdDate    'Used negative to indicate missed counted
        End If
    Next ilSdfIndex
    Exit Function
lObtainCount:
    For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine Then
            ilVefFound = False
            If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(ilSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                tmSmfSrchKey.lChfCode = tmClf.lChfCode
                tmSmfSrchKey.iLineNo = tmClf.iLine
                'slDate = Format$(llChkStartDate, "m/d/yy")
                'gPackDate slDate, ilDate0, ilDate1
                tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                    If tmSmf.lSdfCode = tmSdf.lCode Then
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                        tlSdfExt(ilSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
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
                        If tlSdfExt(ilSdfIndex).iVefCode = tmVsf.iFSCode(ilCVsf) Then
                            ilVefFound = True
                            Exit For
                        End If
                    End If
                Next ilCVsf
            End If
            If Not ilVefFound Then
                If (tlSdfExt(ilSdfIndex).sSchStatus <> "G" And tlSdfExt(ilSdfIndex).sSchStatus <> "O") Then
                    If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                        tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                        mCntrSchdSpotChk = False
                    End If
                End If
                tlSdfExt(ilSdfIndex).iLineNo = -tlSdfExt(ilSdfIndex).iLineNo   'Spot not counted again
            End If
            GoSub lSetStatus
        End If
    Next ilSdfIndex
    'Scan scheduled spots- checking if from this date span
    For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        If tmVsf.iFSCode(ilCVsf) > 0 Then
            ilSdfSpots = 0
            For ilSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                If (tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine) Then
                    If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(ilSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        'Obtain original dates
                        tmSmfSrchKey.lChfCode = tmClf.lChfCode
                        tmSmfSrchKey.iLineNo = tmClf.iLine
                        'slDate = Format$(llChkStartDate, "m/d/yy")
                        'gPackDate slDate, ilDate0, ilDate1
                        tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                        tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                            If tmSmf.lSdfCode = tmSdf.lCode Then
                                ilVefCode = tmSmf.iOrigSchVef
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        Loop
                    Else
                        ilVefCode = tlSdfExt(ilSdfIndex).iVefCode
                    End If
                End If
                If (tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine) And (tmVsf.iFSCode(ilCVsf) = ilVefCode) Then
                    If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                        gUnpackDate tlSdfExt(ilSdfIndex).iDate(0), tlSdfExt(ilSdfIndex).iDate(1), slSchDate
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slSdfDate
                        If ((gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate)) Or ((gDateValue(slSchDate) >= llChkStartDate) And (gDateValue(slSchDate) <= llChkEndDate)) Then
                            'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                            '    mCntrSchdSpotChk = False
                            'End If
                            tlSdfExt(ilSdfIndex).iLineNo = -tlSdfExt(ilSdfIndex).iLineNo   'Spot not counted again
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slOrigMissedDate
                            tlSdfExt(ilSdfIndex).lMdDate = gDateValue(slOrigMissedDate)
                            If ((gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate)) Then
                                If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                    ilSdfSpots = ilSdfSpots + 1
                                End If
                                tlSdfExt(ilSdfIndex).lMdDate = -tlSdfExt(ilSdfIndex).lMdDate    'Use negative to indicate missed counted
                            End If
                            mGetLegalTimes slOrigMissedDate, llStartTime(), llEndTime()
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
                            If (tgCff(ilCff).CffRec.iSpotsWk <> 0) Or (tgCff(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCff(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) And (tgCff(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                    'illegal Date
                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                    '    mCntrSchdSpotChk = False
                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    'End If
                                End If
                            Else
                                If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) Then
                                    'illegal Date
                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                    '    mCntrSchdSpotChk = False
                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    'End If
                                End If
                            End If
                        End If
                    Else
                        gUnpackDate tlSdfExt(ilSdfIndex).iDate(0), tlSdfExt(ilSdfIndex).iDate(1), slSdfDate
                        If (gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate) Then
                            'If tlSdfExt(ilSdfIndex).iVefCode <> tmClf.iVefCode Then
                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H4  'Vehicle invalid
                            '    mCntrSchdSpotChk = False
                            'End If
                            If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                ilSdfSpots = ilSdfSpots + 1
                            End If
                            tlSdfExt(ilSdfIndex).iLineNo = -tlSdfExt(ilSdfIndex).iLineNo   'Spot not counted again
                            'If scheduled or missed spot, check its time
                            mGetLegalTimes slSdfDate, llStartTime(), llEndTime()
                            gUnpackTime tlSdfExt(ilSdfIndex).iTime(0), tlSdfExt(ilSdfIndex).iTime(1), "A", "1", slSdfTime
                            llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
                        
                            ilFound = False
                            For ilTime = 0 To 6 Step 1
                                If (llStartTime(ilTime) > 0 And llEndTime(ilTime) > 0) Then
                                    'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
                                    If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                            Next ilTime
                            If (tlSdfExt(ilSdfIndex).sSpotType <> "X") And (Not ilFound) Then
                                mCntrSchdSpotChk = False
                                tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H2
                            End If
                            ilDay = gWeekDayStr(slSdfDate)
                            If (tgCff(ilCff).CffRec.iSpotsWk <> 0) Or (tgCff(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCff(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) And (tgCff(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                    'illegal Date
                                    If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                        mCntrSchdSpotChk = False
                                        tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    End If
                                End If
                            Else
                                If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) Then
                                    'illegal Date
                                    If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                        mCntrSchdSpotChk = False
                                        tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    'Process MG that are scheduled prior to missed and are in different weeks
                    If (tlSdfExt(ilSdfIndex).sSchStatus = "O") Or (tlSdfExt(ilSdfIndex).sSchStatus = "G") Then
                        If (-tlSdfExt(ilSdfIndex).iLineNo = tmClf.iLine) And (tlSdfExt(ilSdfIndex).lMdDate > 0) Then
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(ilSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            'Obtain original dates
                            tmSmfSrchKey.lChfCode = tmClf.lChfCode
                            tmSmfSrchKey.iLineNo = tmClf.iLine
                            'slDate = Format$(llChkStartDate, "m/d/yy")
                            'gPackDate slDate, ilDate0, ilDate1
                            tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
                            tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
                            ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmClf.lChfCode) And (tmSmf.iLineNo = tmClf.iLine)
                                If tmSmf.lSdfCode = tmSdf.lCode Then
                                    ilVefCode = tmSmf.iOrigSchVef
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            Loop
                            If (tmVsf.iFSCode(ilCVsf) = ilVefCode) Then
                                If ((tlSdfExt(ilSdfIndex).lMdDate >= llChkStartDate) And (tlSdfExt(ilSdfIndex).lMdDate <= llChkEndDate)) Then
                                    If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                        ilSdfSpots = ilSdfSpots + 1
                                    End If
                                    ilDay = gWeekDayLong(tlSdfExt(ilSdfIndex).lMdDate)
                                    If (tgCff(ilCff).CffRec.iSpotsWk <> 0) Or (tgCff(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCff(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
                                        If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) And (tgCff(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
                                            'illegal Date
                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                            '    mCntrSchdSpotChk = False
                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                            'End If
                                        End If
                                    Else
                                        If (tgCff(ilCff).CffRec.iDay(ilDay) = 0) Then
                                            'illegal Date
                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
                                            '    mCntrSchdSpotChk = False
                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                                            'End If
                                        End If
                                    End If
                                    tlSdfExt(ilSdfIndex).lMdDate = -tlSdfExt(ilSdfIndex).lMdDate    'Use negative to indicate missed counted
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilSdfIndex
            If tmVsf.iNoSpots(ilCVsf) * ilCffSpots <> ilSdfSpots Then
                mCntrSchdSpotChk = False
            End If
        End If
    Next ilCVsf
    Return
lSetStatus:
    If tlSdfExt(ilSdfIndex).iLen <> tmClf.iLen Then
        If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
            tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H8  'Length invalid
            mCntrSchdSpotChk = False
        End If
    End If
    'If tmClf.sPriceType = "T" Then
    '    If tlSdfExt(ilSdfIndex).sPriceType = "N" Then
    '        gPDNToStr tmClf.sActPrice, 2, slPrice
    '        If gCompNumberStr(slPrice, "0.00") <> 0 Then
    '            tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H10  'Price discrepancy
    '            mCntrSchdSpotChk = False
    '        End If
    '    ElseIf tlSdfExt(ilSdfIndex).sPriceType = "P" Then
    '        tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H10  'Price discrepancy
    '        mCntrSchdSpotChk = False
    '    End If
    'End If
    Return
End Function
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
Function mDetermineStfIndex(ilStartIndex As Integer, ilEndIndex As Integer) As Integer
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilLoop1 As Integer
    Dim ilLoop2 As Integer
    Dim ilIndex As Integer
    Dim ilAnyRemoved As Integer
    Dim llTime0 As Long
    Dim llTime1 As Long
    Dim slVeh0 As String
    Dim slVeh1 As String
    Dim slTime As String
    Do
        ReDim tmAStf(0 To 0) As STF
        If ilStartIndex >= UBound(tmSort) Then
            mDetermineStfIndex = False
            Exit Function
        End If
        ilRet = btrGetDirect(hmStf, tmAStf(0), imStfRecLen, tmSort(ilStartIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        tmAStf(0).iLineNo = ilStartIndex 'Save index
        ilRet = gParseItem(tmSort(ilStartIndex).sKey, 1, "|", slVeh0)
        ilRet = gParseItem(tmSort(ilStartIndex).sKey, 3, "|", slTime)
        slVeh0 = Trim$(slVeh0)
        slTime = Trim$(slTime)
        llTime0 = Val(slTime)
        ilUpper = 1
        ReDim Preserve tmAStf(0 To ilUpper) As STF
        ilEndIndex = ilStartIndex + 1
        Do While ilEndIndex < UBound(tmSort)
            ilRet = btrGetDirect(hmStf, tmAStf(ilUpper), imStfRecLen, tmSort(ilEndIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            tmAStf(ilUpper).iLineNo = ilEndIndex 'Save index
            ilRet = gParseItem(tmSort(ilEndIndex).sKey, 1, "|", slVeh1)
            ilRet = gParseItem(tmSort(ilEndIndex).sKey, 3, "|", slTime)
            slVeh1 = Trim$(slVeh1)
            slTime = Trim$(slTime)
            llTime1 = Val(slTime)
            If (StrComp(slVeh0, slVeh1, 0) <> 0) Or (tmAStf(0).iLogDate(0) <> tmAStf(ilUpper).iLogDate(0)) Or (tmAStf(0).iLogDate(1) <> tmAStf(ilUpper).iLogDate(1)) Or (llTime0 <> llTime1) Then
                Exit Do
            End If
            ilUpper = ilUpper + 1
            ilEndIndex = ilEndIndex + 1
            ReDim Preserve tmAStf(0 To ilUpper) As STF
        Loop
        ilEndIndex = ilEndIndex - 1
        'Remove pars -Any matching Contract # and length and Add with remove
        ilAnyRemoved = False
        Do
            ilAnyRemoved = False
            For ilLoop1 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                For ilLoop2 = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                    If ilLoop1 <> ilLoop2 Then
                        If (tmAStf(ilLoop1).lChfCode = tmAStf(ilLoop2).lChfCode) And (tmAStf(ilLoop1).iLen = tmAStf(ilLoop2).iLen) And (tmAStf(ilLoop1).sAction <> tmAStf(ilLoop2).sAction) Then
                            ilAnyRemoved = True
                            ilUpper = 0
                            For ilIndex = LBound(tmAStf) To UBound(tmAStf) - 1 Step 1
                                If (ilIndex <> ilLoop1) And (ilIndex <> ilLoop2) Then
                                    tmAStf(ilUpper) = tmAStf(ilIndex)
                                    ilUpper = ilUpper + 1
                                End If
                            Next ilIndex
                            ReDim Preserve tmAStf(0 To ilUpper) As STF
                            Exit For
                        End If
                    End If
                Next ilLoop2
                If ilAnyRemoved Then
                    Exit For
                End If
            Next ilLoop1
        Loop While ilAnyRemoved
        If ilUpper <= 0 Then
            ilStartIndex = ilEndIndex + 1
        End If
    Loop While ilUpper <= 0
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
    slMsg = "Report Error #" & Trim$(Str$(ilError)) & ": " & slMsg
    MsgBox slMsg, vbOkOnly, "Report Error"
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFormatPhoneNo                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove _ characters from number *
'*                                                     *
'*******************************************************
Function mFormatPhoneNo(slPhone As String) As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slTmp As String
    gSetPhoneNo slPhone, RptSelIv!mkcPhone
    If RptSelIv!mkcPhone.Text = sgPhoneImage Then
        slStr = ""
    Else
        slStr = RptSelIv!mkcPhone.Text
        If InStr(slStr, "(____)") <> 0 Then 'Test for missing extension
            ilPos = InStr(slStr, "Ext(")
            slStr = Left$(slStr, ilPos - 1)
        End If
        If InStr(slStr, "(___)") <> 0 Then  'Test for missing area code
            slStr = right$(slStr, Len(slStr) - 5)
        End If
        ilPos = InStr(slStr, "_")
        Do While ilPos > 0
            slTmp = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
            slStr = slTmp
            ilPos = InStr(slStr, "_")
        Loop
    End If
    mFormatPhoneNo = slStr
End Function
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
Sub mGetLegalTimes(slDate As String, llStartTime() As Long, llEndTime() As Long)
    Dim ilTimeAdded As Integer
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
            If (tgSsf(ilDay).sType = "O") And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iStartTime(0) = 0) And (tgSsf(ilDay).iStartTime(1) = 0) Then
                ilSsfInMem = True
                ilRet = BTRV_ERR_NONE
                llRecPos = lgSsfRecPos(ilDay)
            End If
        End If
        If Not ilSsfInMem Then
            imSsfRecLen = Len(tgSsf(ilDay)) 'Max size of variable length record
            tgSsfSrchKey.sType = "O" 'slType
            tgSsfSrchKey.iVefCode = tmClf.iVefCode
            tgSsfSrchKey.iDate(0) = ilDate0
            tgSsfSrchKey.iDate(1) = ilDate1
            tgSsfSrchKey.iStartTime(0) = 0
            tgSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tgSsf(ilDay), imSsfRecLen, tgSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            ilRPRet = gSSFGetPosition(hmSsf, llRecPos)
        End If
        Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilDay).sType = "O") And (tgSsf(ilDay).iVefCode = tmClf.iVefCode) And (tgSsf(ilDay).iDate(0) = ilDate0) And (tgSsf(ilDay).iDate(1) = ilDate1)
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
            If (tgSsf(ilDay).iNextTime(0) = 1) And (tgSsf(ilDay).iNextTime(1) = 0) Then
                Exit Do
            Else
                imSsfRecLen = Len(tgSsf(ilDay)) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tgSsf(ilDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                ilRPRet = gSSFGetPosition(hmSsf, llRecPos)
            End If
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
    Dim slSsfType As String
    Dim ilRet As Integer
    slSsfType = "O" 'On Air
    ilSpotSeqNo = 0
    If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefCode <> tlSdf.iVefCode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
            tmSsfSrchKey.sType = slSsfType
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
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
            ilEvtIndex = 1
            Do
                If ilEvtIndex > tmSsf.iCount Then
                    imSsfRecLen = Len(tmSsf)
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, INDEXKEY0, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
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
    Dim ilCifFound As Integer
    slProduct = ""
    slZone = ""
    slCart = ""
    slISCI = ""
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmCif.lCpfCode > 0 Then
                tmCpfSrchKey.lCode = tmCif.lCpfCode
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmCpf.sISCI = ""
                    tmCpf.sName = ""
                End If
                slISCI = Trim$(tmCpf.sISCI)
                slProduct = Trim$(tmCpf.sName)
            End If
            If tgSpf.sUseCartNo <> "N" Then
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
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfReclen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ' Look for the first positive lZone value
            For ilIndex = 1 To 6 Step 1
                If (tmTzf.lCifZone(ilIndex) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex)), "Oth", 1) <> 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If slZone = "" Then
                            slZone = Trim$(tmTzf.sZone(ilIndex))
                        Else
                            slZone = slZone & Chr$(10) & Trim$(tmTzf.sZone(ilIndex))
                        End If
                        If tmCif.lCpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lCpfCode
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
                        If tgSpf.sUseCartNo <> "N" Then
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
                If (tmTzf.lCifZone(ilIndex) > 0) And (StrComp(Trim$(tmTzf.sZone(ilIndex)), "Oth", 1) = 0) Then ' Process just the first positive Zone
                    ' Read CIF using lCopyCode from SDF
                    If slZone = "" Then
                        slZone = Trim$(tmTzf.sZone(ilIndex))
                    Else
                        slZone = slZone & Chr$(10) & "Other"
                    End If
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmCif.lCpfCode > 0 Then
                            tmCpfSrchKey.lCode = tmCif.lCpfCode
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
                        If tgSpf.sUseCartNo <> "N" Then
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
'*      Procedure Name:mObtainMissedBySOF              *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the missed Sdf records   *
'*                     to be reported for Order only   *
'*                                                     *
'*******************************************************
Sub mObtainMissed(ilSortType As Integer, ilVefCode As Integer, slStartDate As String, slEndDate As String, ilByOrderOrAir As Integer, ilCostType As Integer)
'
'   Where:
'       ilSortType(I)- 0- For SalesBy Advt; 1=Sales by Vehicle (tmVef must contain vehicle name)
'       ilVefCode(I)- Vehicle Code
'       slStartDate(I)- Start Date
'       slEndDate(I)- End Date
'       ilByOrderOrAir(I)- 0=Order; 1=Aired
'       ilCostType(I) - bit map of spot costs to include (n/c, adu, bonus, etc)
'
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llTime As Long
    Dim slMktRank As String
    Dim slRecord As String
    Dim ilUpper As Integer
    Dim ilOk As Integer
    Dim slStr As String
    Dim slPrice As String
    Dim ilSpotSeqNo As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    
    If ilByOrderOrAir <> 0 Then
        Exit Sub
    End If
    If ilSortType = 0 Then
        ilUpper = UBound(tmSpotSOF)
    Else
        ilUpper = UBound(tmPLSdf)
    End If
    btrExtClear hmSmf   'Clear any previous extend operation
    ilExtLen = Len(tmSmf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSmf   'Clear any previous extend operation
    tmSmfSrchKey.lChfCode = 0
    tmSmfSrchKey.iLineNo = 0
    gPackDate slStartDate, tmSmfSrchKey.iMissedDate(0), tmSmfSrchKey.iMissedDate(1)
    ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSmf, llNoRec, -1, "UC") 'Set extract limits (all records)
        'tlIntTypeBuff.iType = ilVefCode
        'ilOffset = gFieldOffset("Smf", "SmfOrigSchVEF")
        'ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Smf", "SmfMissedDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Smf", "SmfMissedDate")
            ilRet = btrExtAddLogicConst(hmSmf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
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
                    If tgSpf.sInvAirOrder = "O" Then    'Test Airing Vehicle
                        If tmSdf.iVefCode <> ilVefCode Then
                            ilOk = False
                        End If
                    Else    'Test Order Vehicle
                        If tmSmf.iOrigSchVef <> ilVefCode Then
                            ilOk = False
                        End If
                    End If
                Else
                    ilOk = False
                End If
                If ilOk Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then          'always ignore psas & promos
                            ilOk = False
                    End If
                        
                    If tmChf.sType = "C" And Not RptSelIv!ckcSelC6(0).value Then  'include std cntrs?
                        ilOk = False
                    End If
                    If tmChf.sType = "V" And Not RptSelIv!ckcSelC6(1).value Then   'include reserves?
                        ilOk = False
                    End If
                    If tmChf.sType = "T" And Not RptSelIv!ckcSelC6(2).value Then   'include remnants?
                        ilOk = False
                    End If
                    If tmChf.sType = "R" And Not RptSelIv!ckcSelC6(3).value Then   'direct response?
                        ilOk = False
                    End If
                    If tmChf.sType = "Q" And Not RptSelIv!ckcSelC6(4).value Then   'per inquiry?
                        ilOk = False
                    End If
                End If
                If ilOk Then                'filter out spots
                    'get line first, to send to filter routine
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                        tmSdf.iVefCode = tmClf.iVefCode
                        If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                            mTestCostType ilOk, ilCostType, slPrice
                        End If
                    Else
                        ilOk = False
                    End If
                End If
                If ilOk Then
                    tmSdf.sSchStatus = "M"
                    If tgSpf.sInvAirOrder <> "O" Then    'Test Airing Vehicle
                        tmSdf.iVefCode = tmSmf.iOrigSchVef
                    End If
                    tmSdf.iDate(0) = tmSmf.iMissedDate(0)
                    tmSdf.iDate(1) = tmSmf.iMissedDate(1)
                    tmSdf.iTime(0) = tmSmf.iMissedTime(0)
                    tmSdf.iTime(1) = tmSmf.iMissedTime(1)
                    If ilSortType = 0 Then
                        'Build Key
                        tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfReclen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
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
                        tmSpotSOF(ilUpper).tSdf = tmSdf
                        slMktRank = Trim$(Str$(tmSof.iMktRank))
                        Do While Len(slMktRank) < 4
                            slMktRank = "0" & slMktRank
                        Loop
                        tmSpotSOF(ilUpper).sKey = slMktRank & "|" & tmAdf.sName & "|" & Trim$(Str$(tmChf.lCntrNo)) & "|" & tmSof.sName
                        ReDim Preserve tmSpotSOF(0 To ilUpper + 1) As SPOTTYPESORT
                        ilUpper = ilUpper + 1
                    Else
                        tmPLSdf(ilUpper).tSdf = tmSdf
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
                        slStr = Trim$(Str$(llTime))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop
                        If ilSpotSeqNo < 10 Then
                            slStr = slStr & "0" & Trim$(Str$(ilSpotSeqNo))
                        Else
                            slStr = slStr & Trim$(Str$(ilSpotSeqNo))
                        End If
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slStr
                        ReDim Preserve tmPLSdf(0 To ilUpper + 1) As SPOTTYPESORT
                        ilUpper = ilUpper + 1
                    End If
                    'End If
                End If
                ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSmf, tmSmf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
mObtainMissedErr:
    ilRet = Err
    Resume Next
End Sub
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
'*******************************************************
Sub mObtainSdf(ilVefCode As Integer, slStartDate As String, slEndDate As String, ilSpotType As Integer, ilBillType As Integer, ilIncludePSA As Integer, ilMissedType As Integer, ilISCIOnly As Integer, ilCostType As Integer, ilByOrderOrAir As Integer, ilIncludeType As Integer, llContrCode As Long)
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
'       ilIncludeType - true to test contract type inclusions, else false to ignore test
'       llContrCode - if selective contract, code # (else 0 for all)
'
    Dim slDate As String
    Dim llDate As Long
    Dim llTime As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilOk As Integer
    Dim ilTestISCI As Integer
    Dim ilSpotSeqNo As Integer
    Dim slProduct As String
    Dim slZone As String
    Dim slCart As String
    Dim slISCI As String
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    
    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
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
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        ilOffset = gFieldOffset("Sdf", "SdfVefCode")
        If (slStartDate <> "") Or (slEndDate <> "") Then
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        Else
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        End If
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
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
                        Else
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
                        Else
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
                        End If
                    End If
                End If
                If (ilOk) And (Not ilIncludePSA) Then
                    tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                        ilOk = False
                    End If
                End If
                If (ilOk And ilIncludeType) Then
                    If tmChf.sType = "C" And Not RptSelIv!ckcSelC6(0).value Then  'include std cntrs?
                        ilOk = False
                    End If
                    If tmChf.sType = "V" And Not RptSelIv!ckcSelC6(1).value Then   'include reserves?
                        ilOk = False
                    End If
                    If tmChf.sType = "T" And Not RptSelIv!ckcSelC6(2).value Then   'include remnants?
                        ilOk = False
                    End If
                    If tmChf.sType = "R" And Not RptSelIv!ckcSelC6(3).value Then   'direct response?
                        ilOk = False
                    End If
                    If tmChf.sType = "Q" And Not RptSelIv!ckcSelC6(4).value Then   'per inquiry?
                        ilOk = False
                    End If
                End If
                If (ilOk) And (ilISCIOnly) Then
                    tmSdf = tmPLSdf(ilUpper).tSdf
                    mObtainCopy slProduct, slZone, slCart, slISCI
                    If Len(slISCI) <= 0 Then
                        tmAdfSrchKey.iCode = tmPLSdf(ilUpper).tSdf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If tmAdf.sShowISCI <> "Y" Then
                            tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                            ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If tmChf.iagfCode > 0 Then     'agency exists
                                tmAgfSrchKey.iCode = tmChf.iagfCode
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
                If ilOk Then            'check for selective contract (zero if all contracts)
                    If llContrCode <> 0 Then
                        If llContrCode <> tmPLSdf(ilUpper).tSdf.lChfCode Then
                            ilOk = False
                        End If
                    End If
                End If
                If ilOk Then
                    'get line first, to send to filter routine
                    tmClfSrchKey.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode
                    tmClfSrchKey.iLine = tmPLSdf(ilUpper).tSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(ilUpper).tSdf.iLineNo) And (tmClf.sSchStatus = "A")
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
                    slStr = Trim$(Str$(llTime))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    If ilSpotSeqNo < 10 Then
                        slStr = slStr & "0" & Trim$(Str$(ilSpotSeqNo))
                    Else
                        slStr = slStr & Trim$(Str$(ilSpotSeqNo))
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
mObtainSdfErr:
    ilRet = Err
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
'*                     reported                        *
'*                                                     *
'*******************************************************
Sub mObtainSdfBySOF(ilVefCode As Integer, slStartDate As String, slEndDate As String, ilMissedType As Integer, ilCostType As Integer, ilByOrderOrAir As Integer)
'
'   Where:
'       ilVefCode(I)- Vehicle Code
'       slStartDate(I)- Start Date
'       slEndDate(I)- End Date
'       ilMissedType(I)-
'       ilCostType(I) - bit map -inclusion of different spot spot types (n/c, adu, bonus, fill, etc)
'       ilByOrderOrAir(I)- 0=Order; 1=Aired
'
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slMktRank As String
    Dim slPrice As String
    Dim ilUpper As Integer
    Dim ilOk As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    ReDim tmSpotSOF(0 To 0) As SPOTTYPESORT
    
    ilUpper = LBound(tmSpotSOF)
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
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        ilOffset = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
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
                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    If ilByOrderOrAir = 0 Then
                        If (tmSdf.sSchStatus = "S") Then
                            ilOk = True
                        Else
                            ilOk = False
                        End If
                    Else
                        ilOk = True
                    End If
                Else
                    'If ilByOrderOrAir = 0 Then
                    '    If (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "C") Then
                    '        ilOk = False
                    '    Else
                    '        ilOk = True
                    '    End If
                    'Else
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
                    'End If
                End If
                If ilOk Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    If (tmChf.sType = "S") Or (tmChf.sType = "M") Then          'always ignore psas & promos
                        ilOk = False
                    End If
                    
                    If tmChf.sType = "C" And Not RptSelIv!ckcSelC6(0).value Then  'include std cntrs?
                        ilOk = False
                    End If
                    If tmChf.sType = "V" And Not RptSelIv!ckcSelC6(1).value Then   'include reserves?
                        ilOk = False
                    End If
                    If tmChf.sType = "T" And Not RptSelIv!ckcSelC6(2).value Then   'include remnants?
                        ilOk = False
                    End If
                    If tmChf.sType = "R" And Not RptSelIv!ckcSelC6(3).value Then   'direct response?
                        ilOk = False
                    End If
                    If tmChf.sType = "Q" And Not RptSelIv!ckcSelC6(4).value Then   'per inquiry?
                        ilOk = False
                    End If
                        
                    If ilOk Then                'filter out spots
                        'get line first, to send to filter routine
                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                        tmClfSrchKey.iLine = tmSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
                            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                            ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                            tmSdf.iVefCode = tmClf.iVefCode
                            If ilCostType >= 0 Then                 'if negative, no testing on spot type selectivity
                                mTestCostType ilOk, ilCostType, slPrice
                            End If
                        Else
                            ilOk = False
                        End If
                    End If
                    If ilOk Then
                        'Build Key
                        tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfReclen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
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
                        tmSpotSOF(ilUpper).tSdf = tmSdf
                        slMktRank = Trim$(Str$(tmSof.iMktRank))
                        Do While Len(slMktRank) < 4
                            slMktRank = "0" & slMktRank
                        Loop
                        tmSpotSOF(ilUpper).sKey = slMktRank & "|" & tmAdf.sName & "|" & Trim$(Str$(tmChf.lCntrNo)) & "|" & tmSof.sName
                        ReDim Preserve tmSpotSOF(0 To ilUpper + 1) As SPOTTYPESORT
                        ilUpper = ilUpper + 1
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
mObtainSdfBySOFErr:
    ilRet = Err
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSelSdf                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Sdf records to be    *
'*                     reported                        *
'*                                                     *
'*******************************************************
Sub mObtainSelSdf(ilVefCode As Integer, slStartDate As String, slEndDate As String, slStartEDate As String, slEndEDate As String, ilSelType As Integer, ilCostType As Integer)
'
'
'   where:
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
    Dim slDate As String
    Dim llDate As Long
    Dim llTime As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilOk As Integer
    Dim llStartEDate As Long
    Dim llEndEDate As Long
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    
    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
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
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        ilOffset = gFieldOffset("Sdf", "SdfVefCode")
        If (slStartDate <> "") Or (slEndDate <> "") Then
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        Else
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        End If
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
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
                ilOk = False
                If ilSelType = 0 Then   'Advertiser
                    For ilLoop = 0 To UBound(tmSelChf) - 1 Step 1
                        If tmSelChf(ilLoop) = tmPLSdf(ilUpper).tSdf.lChfCode Then
                            ilOk = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf ilSelType = 1 Then   'Agency
                    tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    For ilLoop = 0 To UBound(tmSelAgf) - 1 Step 1
                        If tmSelAgf(ilLoop) = tmChf.iagfCode Then
                            ilOk = True
                            Exit For
                        End If
                    Next ilLoop
                ElseIf ilSelType = 2 Then    'Salesperson
                    tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    For ilLoop = 0 To UBound(tmSelSlf) - 1 Step 1
                        If tmSelSlf(ilLoop) = tmChf.iSlfCode(0) Then
                            ilOk = True
                            Exit For
                        End If
                    Next ilLoop
                Else
                    ilOk = True
                End If
                If ilOk Then            'test for spot cost inclusion
                    'get line first, to send to filter routine
                    tmClfSrchKey.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode
                    tmClfSrchKey.iLine = tmPLSdf(ilUpper).tSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(ilUpper).tSdf.iLineNo) And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmPLSdf(ilUpper).tSdf.lChfCode) And (tmClf.iLine = tmPLSdf(ilUpper).tSdf.iLineNo) Then
                        ilRet = gGetSpotPrice(tmPLSdf(ilUpper).tSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, tmPLSdf(ilUpper).sCostType)
                        tmPLSdf(ilUpper).iVefCode = tmClf.iVefCode
                        mTestCostType ilOk, ilCostType, tmPLSdf(ilUpper).sCostType
                        'If Not ilOk Then
                            'ilOk = False
                        'End If
                    Else
                        ilOk = False
                    End If
                End If
                If ilOk Then
                    If (ilSelType = 0) Or (ilSelType = 3) Then   'Advertiser
                        tmChfSrchKey.lCode = tmPLSdf(ilUpper).tSdf.lChfCode
                        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    End If
                    If llStartEDate <> 0 Then
                        gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slDate
                        If gDateValue(slDate) < llStartEDate Then
                            ilOk = False
                        End If
                    End If
                    If llEndEDate <> 0 Then
                        gUnpackDate tmChf.iPropDate(0), tmChf.iPropDate(1), slDate
                        If gDateValue(slDate) > llEndEDate Then
                            ilOk = False
                        End If
                    End If
                    If ilOk Then
                        If tmAdf.iCode <> tmPLSdf(ilUpper).tSdf.iAdfCode Then
                            tmAdfSrchKey.iCode = tmPLSdf(ilUpper).tSdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        End If
                        tmPLSdf(ilUpper).sKey = tmAdf.sName
                        slStr = Trim$(Str$(tmChf.lCntrNo))
                        Do While Len(slStr) < 8
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slStr
                        slStr = Trim$(Str$(tmPLSdf(ilUpper).tSdf.iLineNo))
                        Do While Len(slStr) < 4
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slStr
                        gUnpackDateForSort tmPLSdf(ilUpper).tSdf.iDate(0), tmPLSdf(ilUpper).tSdf.iDate(1), slDate
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slDate
                        If (tmPLSdf(ilUpper).tSdf.sSchStatus = "S") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "G") Or (tmPLSdf(ilUpper).tSdf.sSchStatus = "O") Then
                            tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|A"
                        Else
                            tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|Z"
                        End If
                        gUnpackTimeLong tmPLSdf(ilUpper).tSdf.iTime(0), tmPLSdf(ilUpper).tSdf.iTime(1), False, llTime
                        slStr = Trim$(Str$(llTime))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop
                        tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & "|" & slStr
                        ReDim Preserve tmPLSdf(0 To ilUpper + 1) As SPOTTYPESORT
                        ilUpper = ilUpper + 1
                    End If
                End If
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
mObtainSelSdfErr:
    ilRet = Err
    Resume Next
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
Sub mObtainStf(ilAllRec As Integer, slCreateStartdate As String, slCreateEndDate As String, slAirStartDate As String, slAirEndDate As String)
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
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slRecord As String
    Dim ilUpper As Integer
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
    tlVef.iCode = 0
    ilUpper = LBound(tmSort)
    btrExtClear hmStf   'Clear any previous extend operation
    ilExtLen = Len(tmStf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmStf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmStf, tmStf, imStfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmStf, llNoRec, -1, "UC") 'Set extract limits (all records)
        If ilAllRec <> 7 Then
            If (ilAllRec And 1) <> 0 Then
                tlCharTypeBuff.sType = "R"    'Extract all matching records
                ilOffset = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartdate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") And ((ilAllRec And 2) = 0) And ((ilAllRec And 4) = 0) Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
            If (ilAllRec And 2) <> 0 Then
                tlCharTypeBuff.sType = "P"    'Extract all matching records
                ilOffset = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartdate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") And ((ilAllRec And 4) = 0) Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
            If (ilAllRec And 4) <> 0 Then
                tlCharTypeBuff.sType = "D"    'Extract all matching records
                ilOffset = gFieldOffset("Stf", "StfPrint")
                If (slCreateStartdate = "") And (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                End If
            End If
        End If
        If slCreateStartdate <> "" Then
            gPackDate slCreateStartdate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Stf", "StfCreateDate")
            If (slCreateEndDate = "") And (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slCreateEndDate <> "" Then
            gPackDate slCreateEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Stf", "StfCreateDate")
            If (slAirStartDate = "") And (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirStartDate <> "" Then
            gPackDate slAirStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Stf", "StfLogDate")
            If (slAirEndDate = "") Then
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            End If
        End If
        If slAirEndDate <> "" Then
            gPackDate slAirEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffset = gFieldOffset("Stf", "StfLogDate")
            ilRet = btrExtAddLogicConst(hmStf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
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
                slActionTime = Trim$(Str$(llTime))
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
                    If tmVef.sType = "S" Then
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
                        ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
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
                        ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
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
                                                slTime = Trim$(Str$(llTime))
                                                Do While Len(slTime) < 5
                                                    slTime = "0" & slTime
                                                Loop
                                                tmSort(ilUpper).sKey = tlVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                                tmSort(ilUpper).lRecPos = llRecPos
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmSort(0 To ilUpper) As TYPESORT
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If Not ilVlfFd Then
                            gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                            llDate = gDateValue(slDate)
                            gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                            gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                            slTime = Trim$(Str$(llTime))
                            Do While Len(slTime) < 5
                                slTime = "0" & slTime
                            Loop
                            tmSort(ilUpper).sKey = "~" & tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                            tmSort(ilUpper).lRecPos = llRecPos
                            ilUpper = ilUpper + 1
                            ReDim Preserve tmSort(0 To ilUpper) As TYPESORT
                        End If
                    Else
                        'Create one sort record
                        gUnpackDateForSort tmStf.iLogDate(0), tmStf.iLogDate(1), slDate
                        llDate = gDateValue(slDate)
                        gUnpackTime tmStf.iLogTime(0), tmStf.iLogTime(1), "A", "1", slAirTime
                        gUnpackTimeLong tmStf.iLogTime(0), tmStf.iLogTime(1), False, llTime
                        slTime = Trim$(Str$(llTime))
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
                                tmSort(ilUpper).sKey = tlVefL.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                tmSort(ilUpper).lRecPos = llRecPos
                            Else
                                tmSort(ilUpper).sKey = Left$(tmVef.sName, 8) & " Log Veh Missing" & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                                tmSort(ilUpper).lRecPos = llRecPos
                            End If
                        Else
                            tmSort(ilUpper).sKey = tmVef.sName & "|" & slDate & "|" & slTime & "|" & slActionDate & "|" & slActionTime & "|" & slActionType & "|" & slAirTime
                            tmSort(ilUpper).lRecPos = llRecPos
                        End If
                        ilUpper = ilUpper + 1
                        ReDim Preserve tmSort(0 To ilUpper) As TYPESORT
                    End If
                End If
                ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmStf, tmStf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tmSort(), 0), ilUpper, 0, LenB(tmSort(0)), 0, Len(tmSort(0).sKey), 0 '100, 0
    End If
    Exit Sub
mObtainStfErr:
    ilRet = Err
    Resume Next
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
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
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
    Dim ilOffset As Integer
    ilUpperBound = UBound(tgCff)
    ilFirst = True
    RptSelIv!lbcLnCode.Clear
    btrExtClear hmCff   'Clear any previous extend operation
    ilExtLen = Len(tlCffExt)  'Extract operation record size
    tmCffSrchKey.lChfCode = tgChf.lCode
    tmCffSrchKey.iClfLine = tgClf(ilClfIndex).ClfRec.iLine
    tmCffSrchKey.iCntRevNo = tgClf(ilClfIndex).ClfRec.iCntRevNo
    tmCffSrchKey.iPropVer = tgClf(ilClfIndex).ClfRec.iPropVer
    tmCffSrchKey.iStartDate(0) = 0
    tmCffSrchKey.iStartDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlCff.lChfCode = tgChf.lCode) And (tlCff.iClfLine = tgClf(ilClfIndex).ClfRec.iLine) Then
        'If (tlCff.iClfVersion = tgClf(ilClfIndex).ClfRec.iVersion) And (tlCff.sDelete <> "Y") Then
        '    gUnpackDateForSort tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        '    ilRet = btrGetPosition(hmCff, llRecPos)
        '    slStr = slStr & "\" & Trim$(Str$(llRecPos))
        '    RptSelIv!lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
        'End If
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCff, llNoRec, -1, "UC") 'Set extract limits (all records)
        ilOffset = gFieldOffset("Cff", "CffChfCode")
        ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChf.lCode, 4)
        If ilRet <> BTRV_ERR_NONE Then
            mReadCffRec = False
            Exit Function
        End If
        ilOffset = gFieldOffset("Cff", "CffClfLine")
        ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClf(ilClfIndex).ClfRec.iLine, 2)
        If ilRet <> BTRV_ERR_NONE Then
            mReadCffRec = False
            Exit Function
        End If
        If ilAllVersions Then
            ilOffset = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tgClf(ilClfIndex).ClfRec.iCntRevNo, 2)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
        Else
            ilOffset = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgClf(ilClfIndex).ClfRec.iCntRevNo, 2)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
            ilOffset = gFieldOffset("Cff", "CffDelete")
            ilRet = btrExtAddLogicConst(hmCff, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
            If ilRet <> BTRV_ERR_NONE Then
                mReadCffRec = False
                Exit Function
            End If
        End If
        ilOffset = gFieldOffset("Cff", "CffStartDate")
        ilRet = btrExtAddField(hmCff, ilOffset, ilExtLen)  'Extract start date
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
                slStr = slStr & "\" & Trim$(Str$(llRecPos))
                RptSelIv!lbcLnCode.AddItem slStr    'Add ID (retain matching sorted order) and Code number to list box
                ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCff, tlCffExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmCff   'Clear any previous extend operation
            For ilLoop = 0 To RptSelIv!lbcLnCode.ListCount - 1 Step 1
                slNameCode = RptSelIv!lbcLnCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmCff, tgCff(ilUpperBound).CffRec, imCffRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadCffRec = False
                    Exit Function
                End If
                If tgClf(ilClfIndex).iFirstCff = -1 Then
                    tgClf(ilClfIndex).iFirstCff = ilUpperBound
                Else
                    tgCff(ilUpperBound - 1).iNextCff = ilUpperBound
                End If
                tgCff(ilUpperBound).iNextCff = -1
                tgCff(ilUpperBound).lRecPos = llRecPos
                tgCff(ilUpperBound).iStatus = 1 'Old and retain
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgCff(0 To ilUpperBound) As CFFLIST
                tgCff(ilUpperBound).iStatus = -1 'Not Used
                tgCff(ilUpperBound).iNextCff = -1
                tgCff(ilUpperBound).lRecPos = 0
            Next ilLoop
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
Function mReadChfRec(ilCntrIndex As Integer, ilDispOnly As Integer, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'
'   iRet = mReadChfRec(ilFirstTime)
'   Where:
'       ilFirstTime (I) - True=First time getting contract
'       iRet (O)- True if record read,
'                 False if not read
'
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
    Dim ilLoop As Integer
    ReDim tlSdfExt(1 To 1) As SDFEXT
    Dim ilVersions As Integer
    ilFound = False
    Do
        '
        'If ilFirstTime Then
        '    ilRet = btrGetFirst(hmChf, tgChf, imChfRecLen, INDEXKEY0, BTRV_LOCK_NONE)
        '    ilFirstTime = False
        'Else
        '    ilRet = btrGetNext(hmChf, tgChf, imChfRecLen, BTRV_LOCK_NONE)
        'End If
        If RptSelIv!ckcAll.value Then
            If ilCntrIndex = 0 Then
                If (lgStartingCntrNo > 0) And ilDispOnly Then
                    tmChfSrchKey1.lCntrNo = lgStartingCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmChf, tgChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    lgStartingCntrNo = 0
                Else
                    ilRet = btrGetFirst(hmChf, tgChf, imChfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                End If
            Else
                ilRet = btrGetNext(hmChf, tgChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfRec = False
                Exit Function
            End If
            If ((tgChf.sSchStatus = "F") Or (tgChf.sSchStatus = "I")) And (tgChf.sDelete <> "Y") Then
                If imUpdateCntrNo Then
                    Do
                        ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, lmSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        tmSRec = tmSpf
                        ilRet = gGetByKeyForUpdate("Spf", hmSpf, tmSRec)
                        tmSpf = tmSRec
                        tmSpf.lDiscCurrCntrNo = tgChf.lCntrNo
                        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
                ilCheckCntr = True
            Else
                ilCheckCntr = False
            End If
        Else
            If ilCntrIndex > RptSelIv!lbcSelection(0).ListCount - 1 Then
                mReadChfRec = False
                Exit Function
            End If
            ilCheckCntr = False
            If RptSelIv!lbcSelection(0).Selected(ilCntrIndex) Then
                ilCheckCntr = True
                slNameCode = RptSelIv!lbcCntrCode.List(ilCntrIndex)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmChfSrchKey.lCode = Val(slCode)
                ilRet = btrGetEqual(hmChf, tgChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    mReadChfRec = False
                    Exit Function
                End If
            End If
        End If
        If ilCheckCntr Then
            'If (Not RptSelIv!ckcAll.Value) Or ((tgChf.sStatus <> "M") And (tgChf.sStatus <> "P") And (tgChf.sStatus <> "N")) Then
            'If discrepancies only and manually moved, proposal or new contract- bypass
            '(sType is also tested to cover errors when sStatus was not set correctly)
            If (tgChf.sSchStatus = "A") Or ((ilDispOnly) And ((tgChf.sSchStatus = "M") Or (tgChf.sSchStatus = "P") Or (tgChf.sSchStatus = "N") Or (tgChf.sType = "T") Or (tgChf.sType = "Q") Or (tgChf.sType = "M") Or (tgChf.sType = "S"))) Then
                ilFound = False
            Else
                'Get all spots
                'RptSelIv!lbcSort.Clear
                ReDim tlSdfExtSort(0 To 0) As SDFEXTSORT
                If tgChf.lVefCode > 0 Then                      'all same veh on this order
                    ilRet = gObtainCntrSpot(tgChf.lVefCode, False, tgChf.lCode, -1, "", "", tlSdfExtSort(), tlSdfExt())
                Else                                            'possibly multiple vehicles on order
                    If ((tgChf.sStatus = "M") Or (tgChf.sStatus = "P") Or (tgChf.sStatus = "N") Or (tgChf.sType = "T") Or (tgChf.sType = "Q") Or (tgChf.sType = "M") Or (tgChf.sType = "S")) Then
                        'PSA/ Promo,.. can't have MG so only get spots for specified dates to reduce number of spots
                        slStartDate = Format$(llStartDate, "m/d/yy")
                        slEndDate = Format$(llEndDate, "m/d/yy")
                        'ilRet = gObtainCntrSpot(-1, False, tgChf.lCode, -1, slStartDate, slEndDate, RptSelIv!lbcSort, tlSdfExt())
                        ilRet = gObtainCntrSpot(tgChf.lVefCode, False, tgChf.lCode, -1, slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt())
                    Else
                        'ilRet = gObtainCntrSpot(-1, False, tgChf.lCode, -1, "", "", RptSelIv!lbcSort, tlSdfExt())
                        ilRet = gObtainCntrSpot(tgChf.lVefCode, False, tgChf.lCode, -1, "", "", tlSdfExtSort(), tlSdfExt())
                   End If
                End If
                gUnpackDate tgChf.iStartDate(0), tgChf.iStartDate(1), slDate
                llSDate = gDateValue(slDate)
                gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slDate
                llEDate = gDateValue(slDate)
                If (llEDate >= llStartDate) And (llSDate <= llEndDate) Then
                    ilFound = True
                Else
                    'Determine if any spots are within date span as contract is not
                    For ilLoop = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                        gUnpackDateLong tlSdfExt(ilLoop).iDate(0), tlSdfExt(ilLoop).iDate(1), llDate
                        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                End If
            End If
        End If
        ilCntrIndex = ilCntrIndex + 1
    Loop While Not ilFound
    mReadChfRec = True
    ReDim tgClf(0 To 0) As CLFLIST
    tgClf(0).iStatus = -1 'Not Used
    tgClf(0).lRecPos = 0
    tgClf(0).iFirstCff = -1
    ReDim tgCff(0 To 0) As CFFLIST
    tgCff(0).iStatus = -1 'Not Used
    tgCff(0).lRecPos = 0
    tgCff(0).iNextCff = -1
    ilVersions = 2
    If mReadClfRec(ilVersions) Then
        For ilClfIndex = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            If Not mReadCffRec(ilClfIndex, False) Then
                mReadChfRec = False
                Exit Function
            End If
        Next ilClfIndex
        ''Get all spots
        'RptSelIv!lbcSort.Clear
        'If tgChf.iVefCode > 0 Then
        '    ilRet = gObtainCntrSpot(tgChf.iVefCode, False, tgChf.lCode, -1, "", "", RptSelIv!lbcSort, tlSdfExt())
        'Else
        '    If ((tgChf.sStatus = "M") Or (tgChf.sStatus = "P") Or (tgChf.sStatus = "N") Or (tgChf.sType = "T") Or (tgChf.sType = "Q") Or (tgChf.sType = "M") Or (tgChf.sType = "S")) Then
        '        'PSA/ Promo,.. can't have MG so only get spots for specified dates to reduce number of spots
        '        slStartDate = Format$(llStartDate, "m/d/yy")
        '        slEndDate = Format$(llEndDate, "m/d/yy")
        '        ilRet = gObtainCntrSpot(-1, False, tgChf.lCode, -1, slStartDate, slEndDate, RptSelIv!lbcSort, tlSdfExt())
        '    Else
        '        ilRet = gObtainCntrSpot(-1, False, tgChf.lCode, -1, "", "", RptSelIv!lbcSort, tlSdfExt())
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
'
'   iRet = mReadClfRec(ilVersions)
'   Where:
'       illVersions(I)- 0=All versions; 1=Latest version only; 2=Latest Fully Schedule Versions
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim tlClf As CLF
    Dim tlClfExt As CLFEXT    'Contract line extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim slLine As String
    Dim ilVersion As Integer
    Dim slVersion As String
    Dim ilAddLine As Integer
    Dim ilIndex As Integer
    Dim ilOffset As Integer
    
    ReDim tgClf(0 To 0) As CLFLIST
    ilUpperBound = UBound(tgClf)
    tgClf(ilUpperBound).iStatus = -1 'Not Used
    tgClf(ilUpperBound).lRecPos = 0
    tgClf(ilUpperBound).iFirstCff = -1
    RptSelIv!lbcLnCode.Clear
    btrExtClear hmClf   'Clear any previous extend operation
    ilExtLen = Len(tlClfExt)  'Extract operation record size
    tmClfSrchKey.lChfCode = tgChf.lCode
    tmClfSrchKey.iLine = 0
    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlClf.lChfCode = tgChf.lCode) Then 'And ((ilVersions = 0) Or (ilVersions = 2) Or (tlClf.sDelete <> "Y")) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmClf, llNoRec, -1, "UC") 'Set extract limits (all records)
        If (ilVersions = 0) Then 'Or (ilVersion = 2) Then
            ilOffset = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tgChf.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
        Else
            ilOffset = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tgChf.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
            ilOffset = gFieldOffset("Clf", "ClfDelete")
            ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
            If ilRet <> BTRV_ERR_NONE Then
                mReadClfRec = False
                Exit Function
            End If
        End If
        ilOffset = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddField(hmClf, ilOffset, ilExtLen - 3) 'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            mReadClfRec = False
            Exit Function
        End If
        ilOffset = gFieldOffset("Clf", "ClfSchStatus")
        ilRet = btrExtAddField(hmClf, ilOffset, 1) 'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            mReadClfRec = False
            Exit Function
        End If
        ilOffset = gFieldOffset("Clf", "ClfPropVer")
        ilRet = btrExtAddField(hmClf, ilOffset, 2) 'Extract start/end time, and days
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
                    For ilLoop = 0 To RptSelIv!lbcLnCode.ListCount - 1 Step 1
                        slNameCode = RptSelIv!lbcLnCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slLine)
                        ilRet = gParseItem(slNameCode, 2, "\", slVersion)
                        If tlClfExt.iLine = Val(slCode) Then
                            If tlClfExt.iCntRevNo > Val(slVersion) Then
                                RptSelIv!lbcLnCode.RemoveItem ilLoop
                            Else
                                ilAddLine = False
                            End If
                            Exit For
                        End If
                    Next ilLoop
                ElseIf ilVersions = 2 Then
                    'Manually schedule (M) are only shown when running spot placement
                    If (tlClfExt.sSchStatus = "F") Or (tlClfExt.sSchStatus = "M") Then
                        For ilLoop = 0 To RptSelIv!lbcLnCode.ListCount - 1 Step 1
                            slNameCode = RptSelIv!lbcLnCode.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slLine)
                            ilRet = gParseItem(slNameCode, 2, "\", slVersion)
                            If tlClfExt.iLine = Val(slCode) Then
                                If tlClfExt.iCntRevNo > Val(slVersion) Then
                                    RptSelIv!lbcLnCode.RemoveItem ilLoop
                                Else
                                    ilAddLine = False
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        ilAddLine = False
                    End If
                End If
                If ilAddLine Then
                    slStr = Trim$(Str$(tlClfExt.iLine))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    If ilVersions = 0 Then
                        slVersion = Trim$(Str$(999 - tlClfExt.iCntRevNo))
                        Do While Len(slVersion) < 3
                            slVersion = "0" & slVersion
                        Loop
                    Else
                        slVersion = Trim$(Str$(tlClfExt.iCntRevNo))
                    End If
                    slStr = slStr & "\" & slVersion
                    slStr = slStr & "\" & Trim$(Str$(llRecPos))
                    RptSelIv!lbcLnCode.AddItem slStr
                End If
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmClf   'Clear any previous extend operation
            For ilLoop = 0 To RptSelIv!lbcLnCode.ListCount - 1 Step 1
                slNameCode = RptSelIv!lbcLnCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 3, "\", slCode)
                slCode = Trim$(slCode)
                llRecPos = CLng(slCode)
                ilRet = btrGetDirect(hmClf, tgClf(ilUpperBound).ClfRec, imClfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    mReadClfRec = False
                    Exit Function
                End If
                tgClf(ilUpperBound).iFirstCff = -1
                tgClf(ilUpperBound).lRecPos = llRecPos
                tgClf(ilUpperBound).iStatus = 1 'Old line
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgClf(0 To ilUpperBound) As CLFLIST
                tgClf(ilUpperBound).iStatus = -1 'Not Used
                tgClf(ilUpperBound).iFirstCff = -1
                tgClf(ilUpperBound).lRecPos = 0
            Next ilLoop
        End If
    End If
    mReadClfRec = True
    Exit Function
End Function
'
'                   mSetCostType - set up bit string in ilCostType for
'                   the types of spots to include
'                   Used in spots by ADvt and Spots by Date & Time
'                   Include: Charged, 0.00, ADU, Bonus, Extra, Fill
'                   No Charge, Nc MG, recaptureable, & spinoff
'                   <output> ilCosttype - as bit string
Sub mSetCostType(ilCostType As Integer)
    'ilCostType = 0
    'If RptSelIv!ckcSelC5(0).Value Then
    '    ilCostType = ilCostType Or SPOT_CHARGE          'bit 0
    'End If
    'If RptSelIv!ckcSelC5(1).Value Then
    '    ilCostType = ilCostType Or SPOT_00
    'End If
    'If RptSelIv!ckcSelC5(2).Value Then
    '    ilCostType = ilCostType Or SPOT_ADU
    'End If
    'If RptSelIv!ckcSelC5(3).Value Then
    '    ilCostType = ilCostType Or SPOT_BONUS
    'End If
    'If RptSelIv!ckcSelC5(4).Value Then
    '    ilCostType = ilCostType Or SPOT_EXTRA
    'End If
    'If RptSelIv!ckcSelC5(5).Value Then
    '    ilCostType = ilCostType Or SPOT_FILL
    'End If
    'If RptSelIv!ckcSelC5(6).Value Then
    '    ilCostType = ilCostType Or SPOT_NC
    'End If
    'If RptSelIv!ckcSelC5(7).Value Then
    '    ilCostType = ilCostType Or SPOT_MG
    'End If
    'If RptSelIv!ckcSelC5(8).Value Then
    '    ilCostType = ilCostType Or SPOT_RECAP
    'End If
    'If RptSelIv!ckcSelC5(9).Value Then                    'bit 9
    '    ilCostType = ilCostType Or SPOT_SPINOFF
    'End If
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
'
Sub mSpotSalesTitle(ilMissedType As Integer, slInclude As String, slExclude As String)
    If RptSelIv!rbcSelC7(0).value Then                'ordered
        slInclude = "Include- Ordered"
    Else
        slInclude = "Include- Aired"
    End If
    If RptSelIv!ckcSelC3(0).value Then
        ilMissedType = 1
    End If
    If RptSelIv!ckcSelC3(1).value Then
        ilMissedType = ilMissedType + 2
    End If
    If RptSelIv!ckcSelC3(2).value Then
        ilMissedType = ilMissedType + 4
    End If
    mIncludeExclude RptSelIv!rbcSelC4(0), slInclude, slExclude, "Spots"
    mIncludeExclude RptSelIv!rbcSelC4(1), slInclude, slExclude, "Units"
    mIncludeExclude RptSelIv!ckcSelC3(0), slInclude, slExclude, "Missed"
    mIncludeExclude RptSelIv!ckcSelC3(1), slInclude, slExclude, "Cancel"
    mIncludeExclude RptSelIv!ckcSelC3(2), slInclude, slExclude, "Hidden"
    mIncludeExclude RptSelIv!ckcSelC6(0), slInclude, slExclude, "Std"
    mIncludeExclude RptSelIv!ckcSelC6(1), slInclude, slExclude, "Resv"
    mIncludeExclude RptSelIv!ckcSelC6(2), slInclude, slExclude, "Rem"
    mIncludeExclude RptSelIv!ckcSelC6(3), slInclude, slExclude, "DR"
    mIncludeExclude RptSelIv!ckcSelC6(4), slInclude, slExclude, "PI"
    mIncludeExclude RptSelIv!ckcSelC5(0), slInclude, slExclude, "Charge"
    mIncludeExclude RptSelIv!ckcSelC5(1), slInclude, slExclude, "0.00"
    mIncludeExclude RptSelIv!ckcSelC5(2), slInclude, slExclude, "ADU"
    mIncludeExclude RptSelIv!ckcSelC5(3), slInclude, slExclude, "Bonus"
    mIncludeExclude RptSelIv!ckcSelC5(4), slInclude, slExclude, "Extra"
    mIncludeExclude RptSelIv!ckcSelC5(5), slInclude, slExclude, "Fill"
    mIncludeExclude RptSelIv!ckcSelC5(6), slInclude, slExclude, "NC"
    mIncludeExclude RptSelIv!ckcSelC5(8), slInclude, slExclude, "Recap"
    mIncludeExclude RptSelIv!ckcSelC5(9), slInclude, slExclude, "Spinoff"
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
    'If (InStr(slStrCost, ".") <> 0) Then        'found spot cost
    '    'is it a .00?
    '    If gCompNumberStr(slStrCost, "0.00") = 0 Then       'its a .00 spot
    '        If (ilCostType And SPOT_00) <> SPOT_00 Then      'include .00?
    '            ilOk = False
    '        End If
    '    Else
    '        If (ilCostType And SPOT_CHARGE) <> SPOT_CHARGE Then    'include charged spots?
    '            ilOk = False                                            'no
    '        End If
    '    End If
    'ElseIf Trim$(slStrCost) = "ADU" And (ilCostType And SPOT_ADU) <> SPOT_ADU Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "Bonus" And (ilCostType And SPOT_BONUS) <> SPOT_BONUS Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "Extra" And (ilCostType And SPOT_EXTRA) <> SPOT_EXTRA Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "Fill" And (ilCostType And SPOT_FILL) <> SPOT_FILL Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "N/C" And (ilCostType And SPOT_NC) <> SPOT_NC Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "MG" And (ilCostType And SPOT_MG) <> SPOT_MG Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "Recapturable" And (ilCostType And SPOT_RECAP) <> SPOT_RECAP Then
    '        ilOk = False
    'ElseIf Trim$(slStrCost) = "Spinoff" And (ilCostType And SPOT_SPINOFF) <> SPOT_SPINOFF Then
    '        ilOk = False
    'End If
End Sub
