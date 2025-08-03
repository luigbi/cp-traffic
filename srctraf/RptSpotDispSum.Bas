Attribute VB_Name = "RptSpotDispSum"

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

Dim imOrderWeekCount(0 To 4) As Integer
Dim imAirWeekCount(0 To 4) As Integer

Type VEHICLELLD
    iVefCode As Integer             'vehicle code
    iLLD(0 To 1) As Integer         'vehicles last log date
End Type
Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbf As CBF

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
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1    'SDF record image (key 3)
Dim tmSdfSrchKey3 As LONGKEY0    'SDF record image (key 3)
Dim tmSdfSrchKey4 As SDFKEY4
Dim tmSdfSrchKey7 As SDFKEY7
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF

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


Dim tmChfAdvtExt() As CHFADVTEXT

'**************************************************************************************
'*                                                                                    *
'*      Procedure Name:gSpotDispSumRpt                                                *
'*                                                                                    *
'*            Created:6/16/93       By:D. LeVine                                      *
'*            Modified:             copy of gcntrDispRpt but stripped                 *
'*                                  down to create a 1 line summary per sched         *
'*                                                                                    *
'*            Comments: Generate Contract discrepancy Summary report                  *
'*                                                                                    *

'**************************************************************************************
'
Sub gSpotDispSumRpt()
    Dim llRecsRemaining As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim llCntrIndex As Long
    Dim llStartIndex As Long
    Dim llLbcIndex As Long
    Dim llSdfIndex As Long
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim slInvalid As String
    Dim tlVef As VEF
    ReDim tmSdfExtSort(0 To 0) As SDFEXTSORT
    'ReDim tmSdfExt(1 To 1) As SDFEXT
    ReDim tmSdfExt(0 To 0) As SDFEXT
    ReDim tmVehLLD(0 To 0) As VEHICLELLD
    Dim ilUpperVehLLD As Integer            'total vehicles built into tmVehLLD array
    Dim llVehLatestDate As Long                'all vehicles, latest log date
    Dim ilSlfCode As Integer                'slsp code is running report (only looks at own stuff)
    Dim ilPkgSpot As Integer            '12-20-00 true if spot found that is tied to package line (normally wouldnt occur,
                                        'except for MAI import of M for N spots
    Dim slVefType As String
    Dim ilRepSpot As Integer
    Dim slLLDate As String          'last log as vehicle processed
    Dim llLLDate As Long
    Dim ilMonthSelected As Integer
    Dim ilYear As Integer
    Dim llLoopOnSpots As Long

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
    ilMonthSelected = RptSelCb!cbcSet1.ListIndex
    ilMonthSelected = ilMonthSelected + 1
    
    slStr = RptSelCb!edcSelCTo.Text   'Year
    slDate = Trim$(str(ilMonthSelected)) & "/15/" & Trim$(slStr)
    slDate = gObtainStartStd(slDate)
    llStartDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    llEndDate = gDateValue(slDate)
 
    tmAdf.iCode = 0
    tmVef.iCode = 0
    
    'Darlene: I did not look to see if this can be removed
    gAnyClustersDef 'any markets defined as clusters?

    If lgStartingCntrNo = 0 Then
        ilRet = mObtainCntrForDates(llStartDate, llEndDate)
        llRecsRemaining = UBound(tmChfAdvtExt)
    Else
        llRecsRemaining = 1
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lgStartingCntrNo
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tgChfCB, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
            Exit Sub
        End If
        tmChfAdvtExt(0).lCode = tgChfCB.lCode
    End If
    
    llCntrIndex = LBound(tmChfAdvtExt)
    
    Do
        Screen.MousePointer = vbHourglass
        llStartIndex = llCntrIndex
        ilRet = mReadChfRec(llCntrIndex, llStartDate, llEndDate, tmSdfExtSort(), tmSdfExt())
        If Not ilRet Then
            Exit Do
        End If
        tmChf = tgChfCB
            'Number of lines in the contract
            'One Loop for each contract line
            tmCbf.iExtra2Byte = 0       '9-9-15 init pkg vehicle references  (pkg vehicle code)
            tmCbf.lExtra4Byte = 0       'init pkg vehicle line reference

            
            For ilClf = LBound(tgClfCB) To UBound(tgClfCB) - 1 Step 1
                Screen.MousePointer = vbHourglass
                tmClf = tgClfCB(ilClf).ClfRec

                'ignore overides for m for n customers like MAI
                'For ilLoop = 0 To UBound(tgVpf) Step 1
                '    If tmClf.iVefCode = tgVpf(ilLoop).iVefKCode Then
                    slVefType = ""
                    ilLoop = gBinarySearchVpf(tmClf.iVefCode)
                    If ilLoop <> -1 Then
                       If tgVpf(ilLoop).sGMedium = "S" Then
                           tmClf.iStartTime(0) = 1
                           tmClf.iStartTime(1) = 0
                        End If
                        
                        'need to ignore  BB from days in future, so get the last log date for this vehicle
                        gUnpackDate tgVpf(ilLoop).iLLD(0), tgVpf(ilLoop).iLLD(1), slLLDate
                        If slLLDate = "" Then
                            slLLDate = Format(Now, "m/d/yy")
                        Else
                            If gDateValue(slLLDate) < gDateValue(Format(Now, "m/d/yy")) Then
                                slLLDate = Format(Now, "m/d/yy")
                            End If
                        End If
                        slLLDate = gIncOneDay(slLLDate)
                        llLLDate = gDateValue(slLLDate)   'last log date or todays date +1, whichever is greater

                        ilLoop = gBinarySearchVef(tmClf.iVefCode)       '12-01-06, determine the vehicle type
                        If ilLoop <> -1 Then
                            slVefType = tgMVef(ilLoop).sType
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
                    For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                        llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex
                        If Abs(tmSdfExt(llSdfIndex).iLineNo) = tmClf.iLine Then
                            ilPkgSpot = True
                            Exit For
                        End If
                    Next llLbcIndex
                End If
                ilRepSpot = False

                If slVefType = "R" Then              'ignore Rep lines
                    For llLbcIndex = 0 To UBound(tmSdfExtSort) - 1
                        llSdfIndex = tmSdfExtSort(llLbcIndex).lSdfExtIndex
                        If Abs(tmSdfExt(llSdfIndex).iLineNo) = tmClf.iLine Then
                            ilRepSpot = True
                            Exit For
                        End If
                    Next llLbcIndex
                End If
                If (slVefType <> "R") And (tmClf.sType <> "O" And tmClf.sType <> "A" And tmClf.sType <> "E") Or (ilPkgSpot) Or (ilRepSpot) Then    '12-20-00 ignore package lines only if there wasnt a line tied to a package
                    'determine last log date of the vehicle it's working on
                    If ilSlfCode > 0 Then
                        For ilLoop = 0 To UBound(tmVehLLD) - 1 Step 1
                            If tmVehLLD(ilLoop).iVefCode = tmClf.iVefCode Then
                                gUnpackDate tmVehLLD(ilLoop).iLLD(0), tmVehLLD(ilLoop).iLLD(1), slStr
                                llEndDate = gDateValue(slStr)
                                If slStr = "" Then
                                    llEndDate = llVehLatestDate
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    
                    DoEvents
 
                    ilRet = mCntrSchdSpotChk(ilClf, llStartDate, llEndDate, tmSdfExtSort(), tmSdfExt())
                   
                    DoEvents

                    'ilret = false for 1 of 2 reasons:  1) spot countdiscrepancy, 2) something illegal (invalid time, vehicle)
                    If ilRet = False Then
                        ilRet = True
                        For ilLoop = 0 To 4
                            'tmCbf.lWkCntGrp(ilLoop + 1) = 0   'ordered weekly spot bucket counts
                            'tmCbf.lWkVehGrp(ilLoop + 1) = 0   'aired (found) weekly spot bucket counts
                            tmCbf.lWkCntGrp(ilLoop) = 0   'ordered weekly spot bucket counts
                            tmCbf.lWkVehGrp(ilLoop) = 0   'aired (found) weekly spot bucket counts
                        Next ilLoop
                        'tmCbf.lMonth(1) = 0                     'month ordered count
                        'tmCbf.lMonth(2) = 0                     'month aired count
                        tmCbf.lMonth(0) = 0                     'month ordered count
                        tmCbf.lMonth(1) = 0                     'month aired count
                        For ilLoop = 0 To 4
                            'tmCbf.lWkCntGrp(ilLoop + 1) = imOrderWeekCount(ilLoop)
                            'tmCbf.lWkVehGrp(ilLoop + 1) = imAirWeekCount(ilLoop)   'aired (found) weekly spot bucket counts
                            tmCbf.lWkCntGrp(ilLoop) = imOrderWeekCount(ilLoop)
                            tmCbf.lWkVehGrp(ilLoop) = imAirWeekCount(ilLoop)   'aired (found) weekly spot bucket counts
                            'tmCbf.lMonth(1) = tmCbf.lMonth(1) + imOrderWeekCount(ilLoop)        'accum total ordered for month from weekly buckets
                            'tmCbf.lMonth(2) = tmCbf.lMonth(2) + imAirWeekCount(ilLoop)          'accum total aired (found) for month from weekly buckets
                            tmCbf.lMonth(0) = tmCbf.lMonth(0) + imOrderWeekCount(ilLoop)        'accum total ordered for month from weekly buckets
                            tmCbf.lMonth(1) = tmCbf.lMonth(1) + imAirWeekCount(ilLoop)          'accum total aired (found) for month from weekly buckets
                        Next ilLoop
'                        For ilLoop = 0 To 4
'                            If imOrderWeekCount(ilLoop) <> imAirWeekCount(ilLoop) Then
'                                ilRet = False
'                                Exit For
'                            End If
'                        Next ilLoop
                        'If tmCbf.lMonth(1) <> tmCbf.lMonth(2) Then
                        If tmCbf.lMonth(0) <> tmCbf.lMonth(1) Then
                            ilRet = False
                        End If
                        
                        tmCbf.sLineType = ""
'  this code is to show the error message for Other.  Removed for now.  If needs to be shown, daily Invalid Day/Date code isnt complete.  It needs to maintain the # of spots per day to
'   compare against for flag tht error type
'                        For llLoopOnSpots = LBound(tmSdfExt) To UBound(tmSdfExt) - 1
'                            If Abs(tmSdfExt(llLoopOnSpots).iLineNo) = tmClf.iLine Then
'
'                                If (tmSdfExt(llLoopOnSpots).iStatus And &H1) = &H1 Then
'                                    tmCbf.sLineType = "D"           'invalid day/date
'                                    ilRet = False
'                                    Exit For
'                                ElseIf (tmSdfExt(llLoopOnSpots).iStatus And &H2) = &H2 Then
'                                    tmCbf.sLineType = "T"           'invalid time (this flag isnt set in check routine)
'                                    ilRet = False
'                                    Exit For
'                                ElseIf (tmSdfExt(llLoopOnSpots).iStatus And &H4) = &H4 Then
'                                    tmCbf.sLineType = "V"           'invalid vehicle
'                                    ilRet = False
'                                    Exit For
'                                ElseIf (tmSdfExt(llLoopOnSpots).iStatus And &H8) = &H8 Then
'                                    tmCbf.sLineType = "L"           'invalid spot length
'                                    ilRet = False
'                                    Exit For
'                                End If
'                            End If
'                        Next llLoopOnSpots
                    End If
                    
                    'darlene: ask Jason:  if thereis something else wrong (invalid time, vehicle) and the spot counts match, should we show an error or not?
                    
                    If ilRet = False Then

                        
                        tmCbf.lLineNo = tmClf.iLine
                        'vehicle Name
                        If tmClf.iVefCode <> tmVef.iCode Then
                            tmVefSrchKey.iCode = tmClf.iVefCode
                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                                tmVef.sName = "Missing"
                            End If
                        End If
                        'Vehicle code
                        tmCbf.iVefCode = tmVef.iCode
                        tmCbf.sLineSurvey = Trim$(tmVef.sName)
                        tmCbf.iLen = tmClf.iLen
                        tmCbf.lContrNo = tgChfCB.lCntrNo
                        tmCbf.iAdfCode = tgChfCB.iAdfCode
                        tmCbf.lGenTime = lgNowTime
                        tmCbf.iGenDate(0) = igNowDate(0)
                        tmCbf.iGenDate(1) = igNowDate(1)
                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                        
                    End If
                End If                      'tmclf.stype <> "A" and tmclf.stype <> "O"
            Next ilClf
    Loop

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
    
'   Looks for spot without a contract.  If found, add to array
'    llPrevChfCode = -1
'    btrExtClear hmSdf   'Clear any previous extend operation
'    ilExtLen = Len(tmSdf)
'    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
'    btrExtClear hmSdf   'Clear any previous extend operation
'    gPackDate slStartDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
'    tmSdfSrchKey4.lChfCode = 0
'    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    If ilRet <> BTRV_ERR_END_OF_FILE Then
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Function
'        End If
'        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
'        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'        ilOffset = gFieldOffset("Sdf", "SdfDate")
'        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
'        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'        ilOffset = gFieldOffset("Sdf", "SdfDate")
'        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
'
'
'        ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Function
'        End If
'        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
'                Exit Function
'            End If
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'            Loop
'            Do While ilRet = BTRV_ERR_NONE
'                If llPrevChfCode <> tmSdf.lChfCode Then
'                    ilFound = False
'                    For llChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
'                        If tmSdf.lChfCode = tmChfAdvtExt(llChf).lCode Then
'                            ilFound = True
'                            Exit For
'                        End If
'                    Next llChf
'                    If Not ilFound Then
'                        tmChfAdvtExt(UBound(tmChfAdvtExt)).lCode = tmSdf.lChfCode
'                        ReDim Preserve tmChfAdvtExt(LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) + 1) As CHFADVTEXT
'                    End If
'                End If
'                llPrevChfCode = tmSdf.lChfCode
'                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                Do While ilRet = BTRV_ERR_REJECT_COUNT
'                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                Loop
'            Loop
'        End If
'    End If

    mObtainCntrForDates = True
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
Function mReadChfRec(llCntrIndex As Long, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

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
    Dim llLoop As Long
    'ReDim tlSdfExt(1 To 1) As SDFEXT
    ReDim tlSdfExt(0 To 0) As SDFEXT
    Dim ilVersions As Integer
    Dim llVefCode As Long
    Dim ilIncludeType As Integer        '1-11-03

    ilFound = False
    Do
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
            
            ilCheckCntr = True
        Else
            ilCheckCntr = False
        End If
        If ilCheckCntr Then
            ilIncludeType = False
            If (tgChfCB.sType = "T" And tgSpf.sSchdRemnant = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            ElseIf (tgChfCB.sType = "M" And tgSpf.sSchdPromo = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            ElseIf (tgChfCB.sType = "S" And tgSpf.sSchdPSA = "Y") Or tgChfCB.sSchStatus = "F" Or tgChfCB.sSchStatus = "I" Then
                ilIncludeType = True
            End If

            If (tgChfCB.sSchStatus = "A") Or ((Not ilIncludeType)) Or (tgChfCB.sType = "Q") Or (gIsCntrRemote(tgChfCB, hmVsf)) Then
                ilFound = False
            Else
                'Get all spots
                slStartDate = Format$(llStartDate, "m/d/yy")
                slEndDate = Format$(llEndDate, "m/d/yy")
                llVefCode = 0   'Use key by Contract instead of Vehicle, Contract
                ReDim tlSdfExtSort(0 To 0) As SDFEXTSORT
                If tgChfCB.lVefCode > 0 Then                      'all same veh on this order
                    '5-9-11 Remove all the invalid bb spots that doesnt belong
                    'SDF opened in calling rtn
'                    ilRet = gRemoveBBSpots(hmSdf, CInt(tgChfCB.lVefCode), 0, slStartDate, slEndDate, tgChfCB.lCode, 0)
                    ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                Else                                            'possibly multiple vehicles on order
                    If ((tgChfCB.sStatus = "M") Or (tgChfCB.sStatus = "P") Or (tgChfCB.sStatus = "N") Or (tgChfCB.sType = "T") Or (tgChfCB.sType = "Q") Or (tgChfCB.sType = "M") Or (tgChfCB.sType = "S")) Then
                        ''PSA/ Promo,.. can't have MG so only get spots for specified dates to reduce number of spots
                        'slStartDate = Format$(llStartDate, "m/d/yy")
                        'slEndDate = Format$(llEndDate, "m/d/yy")
                        ''ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, slStartDate, slEndDate, RptSelCb!lbcSort, tlSdfExt())
'                        mRemoveBBSpotSetup slStartDate, slEndDate, tgChfCB
                        ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                    Else
                        'ilRet = gObtainCntrSpot(-1, False, tgChfCB.lCode, -1, "", "", RptSelCb!lbcSort, tlSdfExt())
'                        mRemoveBBSpotSetup slStartDate, slEndDate, tgChfCB
                        ilRet = gObtainCntrSpot(llVefCode, False, tgChfCB.lCode, -1, "S", slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt(), 0, False)
                   End If
                End If
                gUnpackDate tgChfCB.iStartDate(0), tgChfCB.iStartDate(1), slDate

                llSDate = gDateValue(slDate)
                gUnpackDate tgChfCB.iEndDate(0), tgChfCB.iEndDate(1), slDate
                llEDate = gDateValue(slDate)
                If (llEDate >= llStartDate) And (llSDate <= llEndDate) Then
                    ilFound = True
                Else            'we dont include spots without contract so this code should never be executed; removed getting those spots in mObtainCntrForDAte
                    'Determine if any spots are within date span as contract is not
                    'For ilLoop = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                    '    gUnpackDateLong tlSdfExt(ilLoop).iDate(0), tlSdfExt(ilLoop).iDate(1), llDate
                    If UBound(tlSdfExt) > LBound(tlSdfExt) Then
                        For llLoop = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                            gUnpackDateLong tlSdfExt(llLoop).iDate(0), tlSdfExt(llLoop).iDate(1), llDate
                            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                ilFound = True
                                Exit For
                            End If
                        'Next ilLoop
                        Next llLoop
                    Else
                        ilFound = True
                    End If
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
    ReDim llLinePos(0 To 0) As Long

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
    
                llLinePos(UBound(llLinePos)) = llRecPos
                ReDim Preserve llLinePos(0 To UBound(llLinePos) + 1) As Long
                
'                ilAddLine = True
'                If ilVersions = 1 Then
'                    For ilLoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
'                        slNameCode = RptSelCb!lbcLnCode.List(ilLoop)
'                        ilRet = gParseItem(slNameCode, 1, "\", slLine)
'                        ilRet = gParseItem(slNameCode, 2, "\", slVersion)
'                        If tlClfExt.iLine = Val(slCode) Then
'                            If tlClfExt.iCntRevNo > Val(slVersion) Then
'                                RptSelCb!lbcLnCode.RemoveItem ilLoop
'                            Else
'                                ilAddLine = False
'                            End If
'                            Exit For
'                        End If
'                    Next ilLoop
'                ElseIf ilVersions = 2 Then
'                    'Manually schedule (M) are only shown when running spot placement
'                    If (tlClfExt.sSchStatus = "F") Or (tlClfExt.sSchStatus = "M") Then
'                        For ilLoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
'                            slNameCode = RptSelCb!lbcLnCode.List(ilLoop)
'                            ilRet = gParseItem(slNameCode, 1, "\", slLine)
'                            ilRet = gParseItem(slNameCode, 2, "\", slVersion)
'                            If tlClfExt.iLine = Val(slCode) Then
'                                If tlClfExt.iCntRevNo > Val(slVersion) Then
'                                    RptSelCb!lbcLnCode.RemoveItem ilLoop
'                                Else
'                                    ilAddLine = False
'                                End If
'                                Exit For
'                            End If
'                        Next ilLoop
'                    Else
'                        ilAddLine = False
'                    End If
'                End If
'                If ilAddLine Then
'                    slStr = Trim$(str$(tlClfExt.iLine))
'                    Do While Len(slStr) < 4
'                        slStr = "0" & slStr
'                    Loop
'                    If ilVersions = 0 Then
'                        slVersion = Trim$(str$(999 - tlClfExt.iCntRevNo))
'                        Do While Len(slVersion) < 3
'                            slVersion = "0" & slVersion
'                        Loop
'                    Else
'                        slVersion = Trim$(str$(tlClfExt.iCntRevNo))
'                    End If
'                    slStr = slStr & "\" & slVersion
'                    slStr = slStr & "\" & Trim$(str$(llRecPos))
'                    RptSelCb!lbcLnCode.AddItem slStr
'                End If
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                End If
            Loop
            btrExtClear hmClf   'Clear any previous extend operation
            For ilLoop = 0 To UBound(llLinePos) - 1 Step 1
                llRecPos = llLinePos(ilLoop)
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
            Next ilLoop
        End If
    End If
    mReadClfRec = True
    Exit Function
End Function

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
            For ilLoop = 0 To RptSelCb!lbcLnCode.ListCount - 1 Step 1
                slNameCode = RptSelCb!lbcLnCode.List(ilLoop)
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
            Next ilLoop
        End If
    End If
    mReadCffRec = True
    Exit Function
End Function

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
Function mCntrSchdSpotChk(ilClf As Integer, llStartDate As Long, llEndDate As Long, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT) As Integer
'
'   ilRet = gCntrSchdSpotChk(ilClf As Integer, slStartDate)
'   Where:
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
    Dim ilWeekNo As Integer
    
    For ilWeekNo = 0 To 4 Step 1
        imOrderWeekCount(ilWeekNo) = 0
        imAirWeekCount(ilWeekNo) = 0
    Next ilWeekNo
If tmClf.iLine = 7 Then
ilRet = ilRet
End If
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
    
    For ilCVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
        tmVsf.iFSCode(ilCVsf) = 0
    Next ilCVsf
    tmVsf.iFSCode(LBound(tmVsf.iFSCode)) = tmClf.iVefCode
    tmVsf.iNoSpots(LBound(tmVsf.iFSCode)) = 1
    
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
                    gUnpackDateLong tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), llChkStartDate
                    If (llChkStartDate >= llStartDate) And (llChkStartDate <= llEndDate) Then
                        ilWeekNo = (llChkStartDate - llStartDate) \ 7
                        imAirWeekCount(ilWeekNo) = imAirWeekCount(ilWeekNo) + 1
                        mCntrSchdSpotChk = False
                    End If
                    tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1  'Date invalid
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
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
    
    'this handles remaining discrepancies not handled by the count
    For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
        If tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
            gUnpackDateLong tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), llDate
            If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                ilVefFound = False
                If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
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
'                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
'                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                    If tmSdf.sTracer = "*" Then
'                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
'                        ilVefFound = True
'                    End If
'                End If
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
'                                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
'                                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                                    If tmSdf.sTracer = "*" Then
'                                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
'                                        ilDateFound = True
'                                    End If
'                                End If
'                                'mCntrSchdSpotChk = False
'                                'tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
                            End If
                        End If
'                        mGetLegalTimes slOrigMissedDate, tmSmf.iGameNo, llStartTime(), llEndTime()
'                        gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slSdfTime
'                        llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
'
'                        ilFound = False
'                        For ilTime = 0 To 6 Step 1
'                            If (llStartTime(ilTime) >= 0 And llEndTime(ilTime) > 0) Then
'                                'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
'                                If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
'                                    ilFound = True
'                                    Exit For
'                                End If
'                            End If
'                        Next ilTime
'                        If (tlSdfExt(llSdfIndex).sSpotType <> "X") And (Not ilFound) Then
'                            mCntrSchdSpotChk = False
'                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H2
'                        End If
                    End If
                Else
                    ilFlag = False  'Set a flag variable
'                    If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
'                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                        If tmSdf.sTracer = "*" Then
'                            tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
'                            ilFlag = True
'                        End If
'                    End If
                    If Not ilFlag Then
                        If (tlSdfExt(llSdfIndex).sSchStatus <> "G") Then
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                mCntrSchdSpotChk = False
                                tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H1  'Date invalid
                            End If
                        End If
                    End If
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
    Dim ilWeekNo As Integer
    
    mObtainCount = True
    If UBound(tlSdfExt) <= LBound(tlSdfExt) Then
        If ilCffSpots > 0 Then
            ilWeekNo = (llChkStartDate - llStartDate) \ 7
            imOrderWeekCount(ilWeekNo) = imOrderWeekCount(ilWeekNo) + ilCffSpots
            mObtainCount = False
        End If
        Exit Function
    End If
    
    For llSdfIndex = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1

        If tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine Then
            ilVefFound = False
            If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
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
'                If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
'                    ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                    If tmSdf.sTracer = "*" Then
'                        tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
'                        ilVefFound = True
'                    End If
'                End If
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
'                            mGetLegalTimes slOrigMissedDate, tmSmf.iGameNo, llStartTime(), llEndTime()
'                            gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "1", slSdfTime
'                            llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
'                            'If (llSdfTime < llStartTime) Or (llSdfTime > llEndTime) Then
'                                'illegal time
'                                'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
'                                '    mCntrSchdSpotChk = False
'                                '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H2
'                                'End If
'                            'End If
                            'Check Day
'                            ilDay = gWeekDayStr(slSdfDate)
'                            If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
'                                If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) And (tgCffCB(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
'                                    'illegal Date
'                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
'                                    '    mCntrSchdSpotChk = False
'                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
'                                    'End If
'                                End If
'                            Else
'                                If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) Then
'                                    'illegal Date
'                                    'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
'                                    '    mCntrSchdSpotChk = False
'                                    '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
'                                    'End If
'                                End If
'                            End If
                        End If
                    Else
                        gUnpackDate tlSdfExt(llSdfIndex).iDate(0), tlSdfExt(llSdfIndex).iDate(1), slSdfDate
                        If (gDateValue(slSdfDate) >= llChkStartDate) And (gDateValue(slSdfDate) <= llChkEndDate) Then
                            If tlSdfExt(llSdfIndex).sSpotType <> "X" Then
                                ilSdfSpots = ilSdfSpots + 1
                            End If
                            tlSdfExt(llSdfIndex).iLineNo = -tlSdfExt(llSdfIndex).iLineNo   'Spot not counted again
                            'If scheduled or missed spot, check its time
'                            mGetLegalTimes slSdfDate, tlSdfExt(llSdfIndex).iGameNo, llStartTime(), llEndTime()
'                            gUnpackTime tlSdfExt(llSdfIndex).iTime(0), tlSdfExt(llSdfIndex).iTime(1), "A", "1", slSdfTime
'                            llSdfTime = CLng(gTimeToCurrency(slSdfTime, False))
'
'                            ilFound = False
'                            For ilTime = 0 To 6 Step 1
'                                If (llStartTime(ilTime) >= 0 And llEndTime(ilTime) > 0) Then
'                                    If (llSdfTime >= llStartTime(ilTime)) And (llSdfTime <= llEndTime(ilTime)) Then
'                                        ilFound = True
'                                        Exit For
'                                    End If
'                                End If
'                            Next ilTime
'                            If (tlSdfExt(llSdfIndex).sSpotType <> "X") And (Not ilFound) Then
'                                mObtainCount = False
'                                tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H2
'                            End If
                            ilFlag = False  'Set a flag variable
'                            If (tlSdfExt(llSdfIndex).sSchStatus = "M") And (lgMtfNoRecs > 0) Then
'                                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                                If tmSdf.sTracer = "*" Then
'                                    tlSdfExt(llSdfIndex).iStatus = tlSdfExt(llSdfIndex).iStatus Or &H20  'Vehicle invalid
'                                    ilFlag = True
'                                End If
'                            End If
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
                        End If
                    End If
                Else
                    'Process MG that are scheduled prior to missed and are in different weeks
                    If (tlSdfExt(llSdfIndex).sSchStatus = "O") Or (tlSdfExt(llSdfIndex).sSchStatus = "G") Then
                        If (-tlSdfExt(llSdfIndex).iLineNo = tmClf.iLine) And (tlSdfExt(llSdfIndex).lMdDate > 0) Then
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, tlSdfExt(llSdfIndex).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
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
'                                    ilDay = gWeekDayLong(tlSdfExt(llSdfIndex).lMdDate)
'                                    If (tgCffCB(ilCff).CffRec.iSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.iXSpotsWk <> 0) Or (tgCffCB(ilCff).CffRec.sDyWk = "W") Then 'Weekly buy
'                                        If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) And (tgCffCB(ilCff).CffRec.sXDay(ilDay) <> "Y") Then
'                                            'illegal Date
'                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
'                                            '    mCntrSchdSpotChk = False
'                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
'                                            'End If
'                                        End If
'                                    Else
'                                        If (tgCffCB(ilCff).CffRec.iDay(ilDay) = 0) Then
'                                            'illegal Date
'                                            'If tlSdfExt(ilSdfIndex).sSpotType <> "X" Then
'                                            '    mCntrSchdSpotChk = False
'                                            '    tlSdfExt(ilSdfIndex).iStatus = tlSdfExt(ilSdfIndex).iStatus Or &H1
'                                            'End If
'                                        End If
'                                    End If
                                    tlSdfExt(llSdfIndex).lMdDate = -tlSdfExt(llSdfIndex).lMdDate    'Use negative to indicate missed counted
                                End If
                            End If
                        End If
                    End If
                End If
            Next llSdfIndex
            ilWeekNo = (llChkStartDate - llStartDate) \ 7
            imOrderWeekCount(ilWeekNo) = imOrderWeekCount(ilWeekNo) + tmVsf.iNoSpots(ilCVsf) * ilCffSpots
            imAirWeekCount(ilWeekNo) = imAirWeekCount(ilWeekNo) + ilSdfSpots
            
            If tmVsf.iNoSpots(ilCVsf) * ilCffSpots <> ilSdfSpots Then
                mObtainCount = False
            End If
        End If
    Next ilCVsf
End Function

'
'               Cycle thru the VSF for all vehicles used on the contract
'               For each vehicle, remove all bb in future
'
Public Sub mRemoveBBSpotSetup(slStartDate As String, slEndDate As String, tlChf As CHF)
'    Dim ilVefCode As Integer
'    Dim ilVsf As Integer
'    Dim ilRet As Integer
'
'    If tlChf.lVefCode < 0 Then
'        tmVsfSrchKey.lCode = -tlChf.lVefCode
'        ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        Do While ilRet = BTRV_ERR_NONE
'            For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
'                If tmVsf.iFSCode(ilVsf) > 0 Then
'                    ilVefCode = tmVsf.iFSCode(ilVsf)
'                    'do not game #s or line ids
'                    ilRet = gRemoveBBSpots(hmSdf, ilVefCode, 0, slStartDate, slEndDate, tlChf.lCode, 0)
'                End If
'            Next ilVsf
'
'            If tmVsf.lLkVsfCode <= 0 Then
'                Exit Do
'            End If
'            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
'            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    End If
End Sub
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
