Attribute VB_Name = "RPTGENLG"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptgenlg.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  CODESTNCONV                   DALLASFDSORT                                            *
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
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
Type ODFEXT
    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sZone As String * 3
    iEtfCode As Integer         'Event type code
    iEnfCode As Integer         'Event name code
    sProgCode As String * 5 'Program code #
    ianfCode As Integer 'Avail name code
    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sProduct As String * 35 'Product (either from contract or copy)
    iMnfSubFeed As Integer
    iBreakNo As Integer 'Reset at start of each program
    iPositionNo As Integer
    lCefCode As Long
    sShortTitle As String * 15
    iAdfCode As Integer     'advt code
    lEvtIDCefCode As Long       'Event ID from Program Library
    sDupeAvailID As String * 5  'Duplicated Avail ID from Delivery Links (this will be appended to EvtIDCefCode )
    imnfSeg As Integer          '6-19-01 mnf Segment code from chf
    lCifCode As Long            '8-1-16 Copy code
End Type
Public Const ODFEXTPK As String = "IIB3IIB5IIIB35IIILB15ILB5IL"     'add long for cifcode

Type CMMLSUM
    sKey As String * 100
    iVefCode As Integer
    sVehicle As String
    sZone As String * 3  '0=EST; 1=CST; 2=MST; 3=PST
    sAdvt As String * 30
    iAdfCode As Integer
    sProduct As String * 35     'changed from 20 to 35 2/27/98
    sShortTitle As String * 15
    iLen As Integer
    iMFEarliest As Integer  'added 8/30/99
    iMFEarly As Integer     'added 2/1/98
    iMFAM As Integer
    iMFMid As Integer
    iMFPM As Integer
    iMFEve As Integer
    iSaEarliest As Integer  'added 8/30/99
    iSaEarly As Integer     'added 2/1/98
    iSaAM As Integer
    iSaMid As Integer
    iSaPM As Integer
    iSaEve As Integer
    iSuEarliest As Integer  'added 8/30/99
    iSuEarly As Integer     'added 2/1/98
    iSuAM As Integer
    iSuMid As Integer
    iSuPM As Integer
    iSuEve As Integer
    iTotal As Integer
    iDay(0 To 6)  As Integer
    iAirDate(0 To 1) As Integer 'added 10-23-00
    iHourOfDay As Integer       '12-18-02
End Type
Type CODESTNCONV 'VBC NR
    sName As String * 20 'VBC NR
    sCodeStn As String * 5 'VBC NR
End Type 'VBC NR
Type DALLASFDSORT 'VBC NR
    sKey As String * 30 'VBC NR
    sRecord As String * 104 'VBC NR
End Type 'VBC NR
Dim hmEnf As Integer            'Event name file handle
Dim tmEnf As ENF                'ENF record image
Dim tmSEnf As ENF
Dim tmEnfSrchKey As INTKEY0            'ENF record image
Dim imEnfRecLen As Integer        'ENF record length
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0            'ANF record image
Dim imAnfRecLen As Integer        'ANF record length
Dim hmCef As Integer            'Event comments file handle
Dim tmCef As CEF                'CEF record image
Dim tmCefSrchKey As LONGKEY0            'CEF record image
Dim imCefRecLen As Integer        'CEF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmAdf As Integer            'Advertsier name file handle
Dim tmAdf As ADF                'ADF record image
Dim tmAdfSrchKey As INTKEY0            'ADF record image
Dim imAdfRecLen As Integer        'ADF record length

Dim hmCif As Integer            'Copy name file handle
Dim tmCif As CIF                'Copy record image
Dim tmCifSrchKey As LONGKEY0            'Copy key field
Dim imCifRecLen As Integer        'CIF record length

'Short Title
'Copy rotation
'Copy inventory
' Copy Combo Inventory File
'  Copy Product/Agency File
' Time Zone Copy FIle
'  Media code File
'  Library calendar File
Dim hmVef As Integer            'Vehiclee file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim tmMnfSegs() As MNF              '6-19-01
Dim hmGrf As Integer            'Comml Summary report file handle  (prepass in GRF file)
Dim tmGrf As GRF                'GRF record image
Dim tmGrfSrchKey As GRFKEY0     'GRF record key (date & time)
Dim imGrfRecLen As Integer      'GRF record length
Dim hmSvr As Integer            'Seven Day report  file handle
Dim tmSvr As SVR                'SVR record image
Dim tmSvrSrchKey As GRFKEY0     'SVR record key (date & time)
Dim imSvrRecLen As Integer      'SVR record length
Dim hmSsf As Integer
Dim tmSsf As SSF
Dim imSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim smLogName As String
Dim tmOdf0() As ODFEXT
Dim tmOdf1() As ODFEXT
Dim tmOdf2() As ODFEXT
Dim tmOdf3() As ODFEXT
Dim tmOdf4() As ODFEXT
Dim tmOdf5() As ODFEXT
Dim tmOdf6() As ODFEXT
'
'
'                   gClearGrf - clear prepass file GRF - change to use common clear routine gCrGRFClear
'
'           Created:  2/1/98   D.Hosaka
''
'Sub gClearGrf()
'Dim ilRet As Integer
'    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'    ilRet = btrClose(hmGrf)
'    btrDestroy hmGrf
'    Exit Sub
'    End If
'    imGrfRecLen = Len(tmGrf)
'    tmGrfSrchKey.iGenDate(0) = igNowDate(0)
'    tmGrfSrchKey.iGenDate(1) = igNowDate(1)
'    tmGrfSrchKey.lGenTime = lgNowTime       '10-01-01
'     ilRet = btrGetGreaterOrEqual(hmGrf, tmGrf, imGrfRecLen, tmGrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'    Do While (ilRet = BTRV_ERR_NONE) And (tmGrf.iGenDate(0) = igNowDate(0)) And (tmGrf.iGenDate(1) = igNowDate(1)) And (tmGrf.lGenTime = lgNowTime)
'    ilRet = btrDelete(hmGrf)
'    ilRet = btrGetNext(hmGrf, tmGrf, imGrfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'    Loop
'    ilRet = btrClose(hmGrf)
'    btrDestroy hmGrf
'End Sub
'
'
'               Clear 7Day Log file (SVR)
'
'               Created: 1/29/98  D.Hosaka
Sub gClearSvr()
 Dim ilRet As Integer
    hmSvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSvr, "", sgDBPath & "Svr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSvr)
    btrDestroy hmSvr
    Exit Sub
    End If
    imSvrRecLen = Len(tmSvr)
    tmSvrSrchKey.iGenDate(0) = igNowDate(0)
    tmSvrSrchKey.iGenDate(1) = igNowDate(1)
    '10-10-01
    tmSvrSrchKey.lGenTime = lgNowTime
    'tmSvrSrchKey.iGenTime(0) = igNowTime(0)
    'tmSvrSrchKey.iGenTime(1) = igNowTime(1)
    ilRet = btrGetGreaterOrEqual(hmSvr, tmSvr, imSvrRecLen, tmSvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    '10-10-01
    Do While (ilRet = BTRV_ERR_NONE) And (tmSvr.iGenDate(0) = igNowDate(0)) And (tmSvr.iGenDate(1) = igNowDate(1)) And (tmSvr.lGenTime = lgNowTime)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmSvr.iGenDate(0) = igNowDate(0)) And (tmSvr.iGenDate(1) = igNowDate(1)) And (tmSvr.iGenTime(0) = igNowTime(0)) And (tmSvr.iGenTime(1) = igNowTime(1))
    ilRet = btrDelete(hmSvr)
    ilRet = btrGetNext(hmSvr, tmSvr, imSvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmSvr)
    btrDestroy hmSvr
End Sub
'
'
'                   gCmlSum15DP - Create Commercial Summary prepass
'                   1 line per advt/product.  Each advt showing the
'                   following:  5 DP with spot counts for M-f, SA, SU.
'                   Also, for each advt, show days of week that advt airs.
'
'                   This code duplicated from gCmmlSumm and all bridge
'                   stuff omitted.
'                   8/30/99 L35 (AMFM) has 18 DP instead of 15
'                   Created:  02/01/98        D.Hosaka
'                             10/20/99 D Hosaka: take out filter of "N" avails
'                                      if L32
'                   12-13-99 Do not test for CHf if no contract # exits in ODF.
'                   If contr # exists, it bypasses psas and promo spots.
'                   6-20-00 Insert comment pointers for header and footer notations
'                   from VOF table
'
'               dh 10-23-00 multiple weeks did not work for C72 due to date not
'                   stored in record.  Added date to type CMMLSUM
'               dh 8-8-01 prevent same advt & product from printing multiple times depending on day of week
'               dh 8-27-01 Add customized header & footer notes to L32 Comml Summary. chg GRF for additional long field
'               dh 12-18-02 For C73,make it sort advt within hour of day if customized table "Show Hour = Y"
'                   otherwise it is sorted by advertiser
'               dh 5-4-04 Add c84 to this routine, which was cloned from C73
'               dh 7-30-04 match ODF generate date and time when gathering ODF records for prepass
'                           Any old records previously not cleared out where gathered resulting in errorneous log/cp
Sub gCmlSum15DP()
    Dim ilRet As Integer
    Dim ilDBRet As Integer
    Dim hlODF As Integer            'One day log file handle
    Dim tlOdf As ODF
    Dim tlOdfSrchKey As ODFKEY0        'ODF record image
    Dim ilOdfRecLen As Integer        'ODF record length
    Dim ilZone As Integer
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim slDates As String
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llTime As Long
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer
    Dim slLen As String
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilVehicle As Integer
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilVpfIndex As Integer
    Dim ilStartUp As Integer
    Dim slAvailFirstLetter As String
    Dim ilUseZone As Integer
    Dim ilLoopZone As Integer
    Dim llGenlEndTime As Long
    Dim ilHour As Integer
    ReDim llDPStartTime(0 To 4) As Long 'Daypart Start Time for zones
    ReDim llDPEndTime(0 To 4) As Long   'Daypart End Time for zones
    ReDim tlSort(0 To 0) As CMMLSUM
    ilUpper = 0
    slDate = RptSelLg!edcSelCFrom.Text   'Start date
    ilStartUp = True
    If (slDate = "") Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCFrom.SetFocus
        Exit Sub
    End If
    If Not gValidDate(slDate) Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCFrom.SetFocus
        Exit Sub
    End If
    llStartDate = gDateValue(slDate)
    llEndDate = llStartDate + Val(RptSelLg!edcSelCFrom1.Text) - 1
    slDates = Format$(llStartDate, "dddd" & ", " & "m/d/yy") & " To " & Format$(llEndDate, "dddd" & ", " & "m/d/yy")
    slTime = RptSelLg!edcSelCTo.Text   'Start Time
    If (slTime = "") Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCTo.SetFocus
        Exit Sub
    End If
    If Not gValidTime(slTime) Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCTo.SetFocus
        Exit Sub
    End If
    llStartTime = CLng(gTimeToCurrency(slTime, False))
    slTime = RptSelLg!edcSelCTo1.Text   'End Time
    If (slTime = "") Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCTo1.SetFocus
        Exit Sub
    End If
    If Not gValidTime(slTime) Then
        igGenRpt = False
        RptSelLg!frcOutput.Enabled = igOutput
        RptSelLg!frcCopies.Enabled = igCopies
        'RptSelLg!frcWhen.Enabled = igWhen
        RptSelLg!frcFile.Enabled = igFile
        RptSelLg!frcOption.Enabled = igOption
        'RptSelLg!frcRptType.Enabled = igReportType
        Beep
        RptSelLg!edcSelCTo1.SetFocus
        Exit Sub
    End If
    llEndTime = CLng(gTimeToCurrency(slTime, True)) - 1
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlODF)
        btrDestroy hlODF
        Exit Sub
    End If
    ilOdfRecLen = Len(tlOdf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hlODF)
        btrDestroy hmAdf
        btrDestroy hlODF
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hlODF)
        btrDestroy hmAnf
        btrDestroy hmAdf
        btrDestroy hlODF
        Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hlODF)
        btrDestroy hmVef
        btrDestroy hmAnf
        btrDestroy hmAdf
        btrDestroy hlODF
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hlODF)
        btrDestroy hmCHF
        btrDestroy hmVef
        btrDestroy hmAnf
        btrDestroy hmAdf
        btrDestroy hlODF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hlODF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        btrDestroy hmVef
        btrDestroy hmAnf
        btrDestroy hmAdf
        btrDestroy hlODF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    Screen.MousePointer = vbHourglass
    ilZone = 0
    For ilLoop = 0 To 3 Step 1
        If RptSelLg!ckcSelC3(ilLoop).Value = vbChecked Then
            ilZone = ilZone + 1
        End If
    Next ilLoop
    tmVef.iCode = 0
    tmAnf.iCode = 0
    For ilLoopZone = 0 To 3
        If RptSelLg!ckcSelC3(ilLoopZone).Value = vbChecked Then
        For ilVehicle = 0 To UBound(igcodes) - 1 Step 1
            ReDim tlSort(0 To 0) As CMMLSUM
            ilUpper = 0
            ilDBRet = 0
            ilVefCode = igcodes(ilVehicle)
            tmVefSrchKey.iCode = ilVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tmVef.sName = "Missing"
            End If
            ilVpfIndex = -1
            ilUseZone = False
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(ilVefCode)
                If ilLoop <> -1 Then
                    ilVpfIndex = ilLoop
                    'If tgVpf(ilVpfIndex).sGZone(1) = "   " Then
                    If tgVpf(ilVpfIndex).sGZone(0) = "   " Then
                        ilZone = 1                          '1 zone to process
            '            Exit For                            'not using zones
                    Else
                        ilUseZone = True                    'something defined for zones
            '            Exit For
                    End If
                End If
            'Next ilLoop

            'Set titles for week
            For llDate = llStartDate To llEndDate Step 1
                ilDay = gWeekDayLong(llDate)
                'tlOdfSrchKey.iUrfCode = ilUrfCode
                tlOdfSrchKey.iVefCode = ilVefCode
                slDate = Format$(llDate, "m/d/yy")
                gPackDate slDate, tlOdfSrchKey.iAirDate(0), tlOdfSrchKey.iAirDate(1)
                gPackDate slDate, ilAirDate0, ilAirDate1
                slTime = gCurrencyToTime(CCur(llStartTime))
                gPackTime slTime, tlOdfSrchKey.iLocalTime(0), tlOdfSrchKey.iLocalTime(1)
                tlOdfSrchKey.sZone = " "
                tlOdfSrchKey.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlOdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                'Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = ilUrfCode) And (tlOdf.iVefCode = ilVefCode) And (tlOdf.iAirDate(0) = ilAirDate0) And (tlOdf.iAirDate(0) = ilAirDate0)
                Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iVefCode = ilVefCode) And (tlOdf.iAirDate(0) = ilAirDate0) And (tlOdf.iAirDate(1) = ilAirDate1)
                    'the ODF generation date & time must match; otherwise retriving old records for same vehicle which causes too many spots
                    If (tlOdf.iMnfSubFeed = 0) And tlOdf.lGenTime = lgGenTime And tlOdf.iGenDate(0) = igODFGenDate(0) And tlOdf.iGenDate(1) = igODFGenDate(1) Then    'bypass records with subfeed
                        If (tlOdf.iEtfCode = 0) Then
                            slAvailFirstLetter = "L"
                            If tmAnf.iCode <> tlOdf.ianfCode Then
                                If tlOdf.ianfCode <> 0 Then
                                    tmAnfSrchKey.iCode = tlOdf.ianfCode
                                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) Then
                                        slAvailFirstLetter = Left$(Trim$(tmAnf.sName), 1)
                                    End If
                                End If
                            Else
                                slAvailFirstLetter = Left$(Trim$(tmAnf.sName), 1)
                            End If
                            '9-11-00 If (slAvailFirstLetter = "N") Or (sgRnfRptName = "L32") Or sgRnfRptName = "C22" Or sgRnfRptName = "C23" Then   'use avail if its Network avail or any avail for L32
                            If (slAvailFirstLetter = "N") Or (sgRnfRptName = "L32") Or sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C84" Then   '5-4-04 use avail if its Network avail or any avail for L32
                                'Bypass PSA and Promo contracts
                                If tlOdf.lCntrNo = 0 Then
                                    tmChf.lCntrNo = 0
                                    tmChf.sType = "O"           'set to order
                                    tmChf.sSchStatus = "F"      'fully scheduled
                                    ilRet = 0
                                Else
                                    tmChfSrchKey1.lCntrNo = tlOdf.lCntrNo
                                    tmChfSrchKey1.iCntRevNo = 32000
                                    tmChfSrchKey1.iPropVer = 32000
                                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                                End If
                                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlOdf.lCntrNo) And (tmChf.sSchStatus = "A")
                                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                If (ilRet = BTRV_ERR_NONE) And (tmChf.sType <> "S") And (tmChf.sType <> "M") And (tmChf.lCntrNo = tlOdf.lCntrNo) Then
                                    ilZone = -1
                                    If Not ilUseZone Then
                                        tlOdf.sZone = "EST"
                                    End If
                                    Select Case Trim$(tlOdf.sZone)
                                        Case "EST"
                                            ilZone = 0
                                        Case "CST"
                                            ilZone = 1
                                        Case "MST"
                                            ilZone = 2
                                        Case "PST"
                                            ilZone = 3
                                        Case ""
                                            ilZone = ilLoopZone         'fake it out and force to EST
                                            If ilLoopZone = 0 Then
                                                tlOdf.sZone = "EST"
                                            ElseIf ilLoopZone = 1 Then
                                                tlOdf.sZone = "CST"
                                            ElseIf ilLoopZone = 2 Then
                                                tlOdf.sZone = "MST"
                                            Else
                                                tlOdf.sZone = "PST"
                                            End If
                                    End Select
                                    If ilZone >= 0 And ilLoopZone = ilZone Then
                                        If ilZone = 0 Then
                                            llDPStartTime(0) = 0
                                            'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) - 1 '21599    '6am
                                            'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1))
                                            'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) - 1 '35999
                                            'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) '36000  '10Am
                                            'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) - 1 '53999
                                            'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) '54000  '3pm
                                            'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) - 1 '68399
                                            'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) '68400  '7pm
                                            'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(5)) - 1 '86399
                                        
                                            llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(0)) - 1 '21599    '6am
                                            llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1))
                                            llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) - 1 '35999
                                            llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) '36000  '10Am
                                            llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) - 1 '53999
                                            llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) '54000  '3pm
                                            llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) - 1 '68399
                                            llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) '68400  '7pm
                                            llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) - 1 '86399
                                        ElseIf ilZone = 1 Then
                                            llDPStartTime(0) = 0
                                            'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) - 1 '21599    '6am
                                            'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1))
                                            'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) - 1 '35999
                                            'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) '36000  '10Am
                                            'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) - 1 '53999
                                            'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) '54000  '3pm
                                            'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) - 1 '68399
                                            'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) '68400  '7pm
                                            'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(5)) - 1 '86399
                                        
                                            llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(0)) - 1 '21599    '6am
                                            llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(0))
                                            llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) - 1 '35999
                                            llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) '36000  '10Am
                                            llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) - 1 '53999
                                            llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) '54000  '3pm
                                            llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) - 1 '68399
                                            llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) '68400  '7pm
                                            llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) - 1 '86399
                                        ElseIf ilZone = 2 Then
                                            llDPStartTime(0) = 0
                                            'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) - 1 '21599    '6am
                                            'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1))
                                            'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) - 1 '35999
                                            'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) '36000  '10Am
                                            'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) - 1 '53999
                                            'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) '54000  '3pm
                                            'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) - 1 '68399
                                            'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) '68400  '7pm
                                            'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(5)) - 1 '86399
                                            
                                            llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(0)) - 1 '21599    '6am
                                            llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(0))
                                            llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) - 1 '35999
                                            llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) '36000  '10Am
                                            llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) - 1 '53999
                                            llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) '54000  '3pm
                                            llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) - 1 '68399
                                            llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) '68400  '7pm
                                            llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) - 1 '86399
                                        Else
                                            llDPStartTime(0) = 0
                                            'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) - 1 '21599    '6am
                                            'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1))
                                            'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) - 1 '35999
                                            'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) '36000  '10Am
                                            'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) - 1 '53999
                                            'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) '54000  '3pm
                                            'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) - 1 '68399
                                            'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) '68400  '7pm
                                            'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(5)) - 1 '86399
                                        
                                            llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(0)) - 1 '21599    '6am
                                            llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(0))
                                            llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) - 1 '35999
                                            llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) '36000  '10Am
                                            llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) - 1 '53999
                                            llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) '54000  '3pm
                                            llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) - 1 '68399
                                            llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) '68400  '7pm
                                            llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) - 1 '86399
                                        End If
                                        For ilTest = 1 To 4 Step 1
                                            If (StrComp(tgVpf(ilVpfIndex).sMapZone(ilTest - 1), Trim$(tlOdf.sZone), 1) = 0) And (StrComp(tgVpf(ilVpfIndex).sMapProgCode(ilTest - 1), tlOdf.sProgCode, 1) = 0) Then
                                                gUnpackTime tlOdf.iLocalTime(0), tlOdf.iLocalTime(1), "A", "1", slTime
                                                llTime = CLng(gTimeToCurrency(slTime, False))
                                                llDPEndTime(tgVpf(ilVpfIndex).iMapDPNo(ilTest - 1) - 2) = llTime - 1
                                                llDPStartTime(tgVpf(ilVpfIndex).iMapDPNo(ilTest - 1) - 1) = llTime '9Am
                                            End If
                                        Next ilTest
                                        If RptSelLg!ckcSelC3(ilZone).Value = vbChecked Then
                                            gUnpackTime tlOdf.iLocalTime(0), tlOdf.iLocalTime(1), "A", "1", slTime
                                            llTime = CLng(gTimeToCurrency(slTime, False))
                                            'mod 2/2/98 chg from lldpstarttime(1) to lldpstarttime(0)
                                            llGenlEndTime = llDPEndTime(4)
                                            If sgRnfRptName = "L35" Then
                                                llGenlEndTime = 86400           '12m
                                            End If
                                            'If (llTime >= llDPStartTime(0)) And (llTime <= llDPEndTime(4)) Then  'after 6am and before 12M
                                            If (llTime >= llDPStartTime(0)) And (llTime <= llGenlEndTime) Then  'testing for correct time 8/30/99
                                                'Determine hour of this spot.  For C73 or C84, check to see if by hour.  If so, the breakout is Advt within Hour of day, so build
                                                'table with hour as part of the keyfield.  For all other cp/logs, make the hour the same
                                                If tgVof.sShowHour = "Y" And (sgRnfRptName = "C73" Or sgRnfRptName = "C84") Then
                                                    ilHour = llTime \ 3600          'obtain the hour of day
                                                Else
                                                    ilHour = 0
                                                End If

                                                gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "1", True, slLen
                                                ilLen = Val(slLen)
                                                Do While Len(slLen) < 3
                                                    slLen = "0" & slLen
                                                Loop
                                                ilFound = -1
                                                For ilTest = 0 To ilUpper - 1 Step 1
                                                    If tgSpf.sUseProdSptScr = "P" Then
                                                        'using short titles , not advt/prod names
                                                        If (tlSort(ilTest).iVefCode = tlOdf.iVefCode) And (StrComp(Trim$(tlSort(ilTest).sZone), Trim$(tlOdf.sZone), 1) = 0) And (tlSort(ilTest).iAdfCode = tlOdf.iAdfCode) And (Trim$(tlSort(ilTest).sShortTitle) = Trim$(tlOdf.sShortTitle)) And (tlSort(ilTest).iLen = ilLen) Then
                                                            If (tlSort(ilTest).iHourOfDay = ilHour) Then
                                                                ilFound = ilTest
                                                                Exit For
                                                            End If
                                                        End If
                                                    Else
                                                        'use advt/prod names, not short title
                                                        If (tlSort(ilTest).iVefCode = tlOdf.iVefCode) And (StrComp(Trim$(tlSort(ilTest).sZone), Trim$(tlOdf.sZone), 1) = 0) And (tlSort(ilTest).iAdfCode = tlOdf.iAdfCode) And (Trim$(tlSort(ilTest).sProduct) = Trim$(tlOdf.sProduct)) And (tlSort(ilTest).iLen = ilLen) Then
                                                            '8-8-01 removed: If tlSort(ilTest).iAirDate(0) = ilAirDate0 And tlSort(ilTest).iAirDate(1) = ilAirDate1 Then
                                                            If (tlSort(ilTest).iHourOfDay = ilHour) Then
                                                                ilFound = ilTest
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next ilTest
                                                If ilFound = -1 Then
                                                    ilFound = ilUpper
                                                    ilUpper = ilUpper + 1
                                                    ReDim Preserve tlSort(0 To ilUpper) As CMMLSUM
                                                    If tlOdf.iAdfCode <> tmAdf.iCode Then
                                                        tmAdfSrchKey.iCode = tlOdf.iAdfCode
                                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tmAdf.sName = "Missing"
                                                        End If
                                                    End If
                                                    slStr = Trim$(str$(ilZone))
                                                    If slStr = "" Then
                                                        slStr = "0"
                                                    End If
                                                    'tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tlOdf.sProduct & slLen
                                                    tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tlOdf.sShortTitle & slLen
                                                    tlSort(ilFound).iVefCode = tlOdf.iVefCode
                                                    tlSort(ilFound).sVehicle = Trim$(tmVef.sName)
                                                    tlSort(ilFound).sZone = tlOdf.sZone
                                                    tlSort(ilFound).sAdvt = Trim$(tmAdf.sName)
                                                    tlSort(ilFound).iAdfCode = tlOdf.iAdfCode
                                                    tlSort(ilFound).sProduct = Trim$(tlOdf.sProduct)
                                                    tlSort(ilFound).sShortTitle = Trim$(tlOdf.sShortTitle)   'Trim$(tlOdf.sProduct)
                                                    tlSort(ilFound).iLen = ilLen
                                                    tlSort(ilFound).iMFEarliest = 0
                                                    tlSort(ilFound).iSaEarliest = 0
                                                    tlSort(ilFound).iSuEarliest = 0
                                                    tlSort(ilFound).iMFEarly = 0
                                                    tlSort(ilFound).iSaEarly = 0
                                                    tlSort(ilFound).iSuEarly = 0
                                                    tlSort(ilFound).iMFAM = 0
                                                    tlSort(ilFound).iSaAM = 0
                                                    tlSort(ilFound).iSuAM = 0
                                                    tlSort(ilFound).iMFMid = 0
                                                    tlSort(ilFound).iSaMid = 0
                                                    tlSort(ilFound).iSuMid = 0
                                                    tlSort(ilFound).iMFPM = 0
                                                    tlSort(ilFound).iSaPM = 0
                                                    tlSort(ilFound).iSuPM = 0
                                                    tlSort(ilFound).iMFEve = 0
                                                    tlSort(ilFound).iSaEve = 0
                                                    tlSort(ilFound).iSuEve = 0

                                                    tlSort(ilFound).iTotal = 0
                                                    tlSort(ilFound).iAirDate(0) = ilAirDate0
                                                    tlSort(ilFound).iAirDate(1) = ilAirDate1
                                                    tlSort(ilFound).iHourOfDay = ilHour
                                                End If
                                                If sgRnfRptName = "L35" Then
                                                    If llTime >= llDPStartTime(1) And llTime <= llDPEndTime(1) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFEarly = tlSort(ilFound).iMFEarly + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaEarly = tlSort(ilFound).iSaEarly + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuEarly = tlSort(ilFound).iSuEarly + 1
                                                        End If
                                                    ElseIf llTime >= llDPStartTime(2) And llTime <= llDPEndTime(2) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFAM = tlSort(ilFound).iMFAM + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaAM = tlSort(ilFound).iSaAM + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuAM = tlSort(ilFound).iSuAM + 1
                                                        End If
                                                    ElseIf llTime >= llDPStartTime(3) And llTime <= llDPEndTime(3) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFMid = tlSort(ilFound).iMFMid + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaMid = tlSort(ilFound).iSaMid + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuMid = tlSort(ilFound).iSuMid + 1
                                                        End If
                                                    ElseIf llTime >= llDPStartTime(4) And llTime <= llDPEndTime(4) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFPM = tlSort(ilFound).iMFPM + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaPM = tlSort(ilFound).iSaPM + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuPM = tlSort(ilFound).iSuPM + 1
                                                        End If
                                                    Else
                                                        If llTime >= llDPStartTime(0) And llTime <= llDPEndTime(0) Then   'is log event time within first DP
                                                            If ilDay <= 4 Then  'M-F
                                                                tlSort(ilFound).iMFEarliest = tlSort(ilFound).iMFEarliest + 1
                                                            ElseIf ilDay = 5 Then   'Sa
                                                                tlSort(ilFound).iSaEarliest = tlSort(ilFound).iSaEarliest + 1
                                                            Else    'Sun
                                                                tlSort(ilFound).iSuEarliest = tlSort(ilFound).iSuEarliest + 1
                                                            End If
                                                        Else
                                                            If ilDay <= 4 Then  'M-F
                                                                tlSort(ilFound).iMFEve = tlSort(ilFound).iMFEve + 1
                                                            ElseIf ilDay = 5 Then   'Sa
                                                                tlSort(ilFound).iSaEve = tlSort(ilFound).iSaEve + 1
                                                            Else    'Sun
                                                                tlSort(ilFound).iSuEve = tlSort(ilFound).iSuEve + 1
                                                            End If

                                                        End If
                                                    End If
                                                Else
                                                    If llTime < llDPStartTime(1) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFEarly = tlSort(ilFound).iMFEarly + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaEarly = tlSort(ilFound).iSaEarly + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuEarly = tlSort(ilFound).iSuEarly + 1
                                                        End If
                                                    ElseIf llTime < llDPStartTime(2) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFAM = tlSort(ilFound).iMFAM + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaAM = tlSort(ilFound).iSaAM + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuAM = tlSort(ilFound).iSuAM + 1
                                                        End If
                                                    ElseIf llTime < llDPStartTime(3) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFMid = tlSort(ilFound).iMFMid + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaMid = tlSort(ilFound).iSaMid + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuMid = tlSort(ilFound).iSuMid + 1
                                                        End If
                                                    ElseIf llTime < llDPStartTime(4) Then
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFPM = tlSort(ilFound).iMFPM + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaPM = tlSort(ilFound).iSaPM + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuPM = tlSort(ilFound).iSuPM + 1
                                                        End If
                                                    Else
                                                        If ilDay <= 4 Then  'M-F
                                                            tlSort(ilFound).iMFEve = tlSort(ilFound).iMFEve + 1
                                                        ElseIf ilDay = 5 Then   'Sa
                                                            tlSort(ilFound).iSaEve = tlSort(ilFound).iSaEve + 1
                                                        Else    'Sun
                                                            tlSort(ilFound).iSuEve = tlSort(ilFound).iSuEve + 1
                                                        End If
                                                    End If
                                                End If
                                                tlSort(ilFound).iTotal = tlSort(ilFound).iTotal + 1
                                                'tlSort(ilFound).iDay(ilDay) = 1
                                                tlSort(ilFound).iDay(ilDay) = tlSort(ilFound).iDay(ilDay) + 1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrGetNext(hlODF, tlOdf, ilOdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            Next llDate
            'Loop thru tlsort and create one disk record for entry
            For ilIndex = LBound(tlSort) To UBound(tlSort) - 1 Step 1
                Select Case Trim$(tlSort(ilIndex).sZone)
                    Case "EST"
                        If ilUseZone Then
                            tmGrf.sBktType = "E"
                        Else
                            tmGrf.sBktType = " "
                        End If
                    Case "CST"
                        tmGrf.sBktType = "C"
                    Case "MST"
                        tmGrf.sBktType = "M"
                    Case "PST"
                        tmGrf.sBktType = "P"
                End Select
                tmGrf.iGenDate(0) = igNowDate(0)
                tmGrf.iGenDate(1) = igNowDate(1)
                '10-10-01
                tmGrf.lGenTime = lgNowTime
                'tmGrf.iGenTime(0) = igNowTime(0)
                'tmGrf.iGenTime(1) = igNowTime(1)
                tmGrf.iVefCode = tlSort(ilIndex).iVefCode
                tmGrf.iAdfCode = tlSort(ilIndex).iAdfCode
                If tgSpf.sUseProdSptScr = "P" Then
                    tmGrf.sGenDesc = Trim$(tlSort(ilIndex).sShortTitle)
                Else
                    tmGrf.sGenDesc = Trim$(tlSort(ilIndex).sProduct)
                End If
                tmGrf.iCode2 = tlSort(ilIndex).iLen
                'loop to setup the spot counts per day and daypart
                'not used for C22 & C23
                '9-11-00 If sgRnfRptName = "C22" Or sgRnfRptName = "C23" Then        '6-20-00
                If sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C84" Then            '5-4-04
                    'tmGrf.lDollars(1) = tgVof.lHd1CefCode                   '6-20-00 header comment
                    'tmGrf.lDollars(2) = tgVof.lFt1CefCode                   '6-20-00 footer #1 comment
                    'tmGrf.lDollars(3) = tgVof.lFt2CefCode                   '6-20-00 footer #2 comment
                    tmGrf.lDollars(0) = tgVof.lHd1CefCode                   '6-20-00 header comment
                    tmGrf.lDollars(1) = tgVof.lFt1CefCode                   '6-20-00 footer #1 comment
                    tmGrf.lDollars(2) = tgVof.lFt2CefCode                   '6-20-00 footer #2 comment
                    tmGrf.iStartDate(0) = tlSort(ilIndex).iAirDate(0)
                    tmGrf.iStartDate(1) = tlSort(ilIndex).iAirDate(1)
                Else
                    tmGrf.lCode4 = tgVof.lHd1CefCode                   '8-27-01 header comment
                    tmGrf.lChfCode = tgVof.lFt1CefCode                 '8-27-01 footer #1 comment
                    tmGrf.lLong = tgVof.lFt2CefCode                 '8-27-01 footer #2 comment
'                    tmGrf.lDollars(1) = tlSort(ilIndex).iMFEarly
'                    tmGrf.lDollars(2) = tlSort(ilIndex).iMFAM
'                    tmGrf.lDollars(3) = tlSort(ilIndex).iMFMid
'                    tmGrf.lDollars(4) = tlSort(ilIndex).iMFPM
'                    tmGrf.lDollars(5) = tlSort(ilIndex).iMFEve
'                    tmGrf.lDollars(6) = tlSort(ilIndex).iSaEarly
'                    tmGrf.lDollars(7) = tlSort(ilIndex).iSaAM
'                    tmGrf.lDollars(8) = tlSort(ilIndex).iSaMid
'                    tmGrf.lDollars(9) = tlSort(ilIndex).iSaPM
'                    tmGrf.lDollars(10) = tlSort(ilIndex).iSaEve
'                    tmGrf.lDollars(11) = tlSort(ilIndex).iSuEarly
'                    tmGrf.lDollars(12) = tlSort(ilIndex).iSuAM
'                    tmGrf.lDollars(13) = tlSort(ilIndex).iSuMid
'                    tmGrf.lDollars(14) = tlSort(ilIndex).iSuPM
'                    tmGrf.lDollars(15) = tlSort(ilIndex).iSuEve
'                    tmGrf.lDollars(16) = tlSort(ilIndex).iMFEarliest        '8/30/99
'                    tmGrf.lDollars(17) = tlSort(ilIndex).iSaEarliest        '8/30/99
'                    tmGrf.lDollars(18) = tlSort(ilIndex).iSuEarliest        '8/30/99
                    tmGrf.lDollars(0) = tlSort(ilIndex).iMFEarly
                    tmGrf.lDollars(1) = tlSort(ilIndex).iMFAM
                    tmGrf.lDollars(2) = tlSort(ilIndex).iMFMid
                    tmGrf.lDollars(3) = tlSort(ilIndex).iMFPM
                    tmGrf.lDollars(4) = tlSort(ilIndex).iMFEve
                    tmGrf.lDollars(5) = tlSort(ilIndex).iSaEarly
                    tmGrf.lDollars(6) = tlSort(ilIndex).iSaAM
                    tmGrf.lDollars(7) = tlSort(ilIndex).iSaMid
                    tmGrf.lDollars(8) = tlSort(ilIndex).iSaPM
                    tmGrf.lDollars(9) = tlSort(ilIndex).iSaEve
                    tmGrf.lDollars(10) = tlSort(ilIndex).iSuEarly
                    tmGrf.lDollars(11) = tlSort(ilIndex).iSuAM
                    tmGrf.lDollars(12) = tlSort(ilIndex).iSuMid
                    tmGrf.lDollars(13) = tlSort(ilIndex).iSuPM
                    tmGrf.lDollars(14) = tlSort(ilIndex).iSuEve
                    tmGrf.lDollars(15) = tlSort(ilIndex).iMFEarliest        '8/30/99
                    tmGrf.lDollars(16) = tlSort(ilIndex).iSaEarliest        '8/30/99
                    tmGrf.lDollars(17) = tlSort(ilIndex).iSuEarliest        '8/30/99
                End If
                For ilLoop = 0 To 6
                    'tmGrf.iPerGenl(ilLoop + 1) = tlSort(ilIndex).iDay(ilLoop)
                    tmGrf.iPerGenl(ilLoop) = tlSort(ilIndex).iDay(ilLoop)
                '    If tlSort(ilIndex).iDay(ilLoop) > 0 Then
                '        tmGrf.iPerGenl(ilLoop + 1) = 1             'show X for spots allowed on this day on report
                '    Else
                '        tmGrf.iPerGenl(ilLoop + 1) = 0             'show nospots allowed for this day on report
                '    End If
                Next ilLoop
                tmGrf.iYear = tlSort(ilIndex).iHourOfDay            '12-18-02

                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            Next ilIndex
        Next ilVehicle
        End If
    Next ilLoopZone

    Erase llDPStartTime
    Erase llDPEndTime
    Erase tlSort
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hmGrf)
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    btrDestroy hmGrf
    Exit Sub
    Return
End Sub






'
'
'               Create 7 day format in prepass SVR  (Seven Day Report File)
'               Loop on time zones, then 7 days within zone, gathering
'               from ODF all entries.  From that, extract the spot types
'               and write one record to disk for each unique zone, vehicle,
'               event type(spots only), avail time, and position.
'
'               To use this format, enter the end times for each of the dayparts
'               in the Vehicle Options Interface table.
'               For example, AMFM has DP 5-6a, 6-10a,10a-3p,3p-7p, 7p-12m
'               Enter end times of 6A, 10A, 3P, 7P, 12M
'               No short title option is tested.
'               Only spots are obtained.
'
'               8/30/99 dh L34 - AMFM needs additional DP from 12M-5A.  In keeping
'               with the end times defined in the Interface Vehicle options table,
'               assume the last DP ends at 12m.  This allows the user to have:
'               ie.  5a, 6a, 10a, 3p, & 7p--giving
'               12m-5a, 5a-6a, 6a-10a, 10a-3p, 3p-7p, & 7p-12m.
'
'               1/3/00 Add code to place daypart text into description field (svr.sprograminfo)
'               for AMFM L40 (the end times are the DP end times:  assume 0 to be a time of 6a for start of day)
'               9-12-00 Update VOF comment codes into SVR
'
'               8-22-01 When a program (avails) do not exist on Monday, the M-F spot ID doesnt appear.
'               Before creating the prepass record (SVR), loop and see if theres a spot on Tuesday-Fri and
'               pick up the spot id from that day.
'               Created : 1/29/98  D. Hosaka
'
Sub gCreate7Day()
Dim hlODF As Integer
Dim hlSvr As Integer
Dim ilRet As Integer
Dim slDate As String
Dim slLen As String
Dim llOdfTime As Long
Dim llSvrTime As Long
Dim ilUseZone As Integer            'true if at least one zone defined
Dim slZone As String * 3            'EST, CST, MST, PST or blank
Dim ilZoneLoop As Integer           'time zone to process in loop
Dim ilDay As Integer                'day of week to process
Dim ilDayOfWeek As Integer
Dim llStartOfWk As Long             'start date of week (temp)
ReDim ilStartofWk(0 To 1) As Integer  'start date of week btrieve format to store in svr
Dim ilRec As Integer                'event within week to write to disk
Dim ilFoundUnique As Integer        'found unique SVR entry (same zone, veh, event type,time, position) for week
Dim ilOdf As Integer                'event to process for day from ODF table
Dim ilUpper As Integer
Dim ilVpfIndex As Integer           'vehicle options index
Dim ilZones As Integer              'zones requested by user
Dim ilLoZone As Integer             'lo limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilHizone.  if all Zones, this will be 1, and ilHIzones will be 4
Dim ilHiZone As Integer             'Hi limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilLozone.  If all zones, this will be a 4.
'ReDim llZoneEndTimes(1 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime)
ReDim llZoneEndTimes(0 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime). Index zero ignored
Dim ilDPIndex As Integer                   'DP index used to separate the DP on crystal report (stored in svr)
Dim ilLoop As Integer               '8-22-01
'**** parameters passed from Log program
Dim llStartTime As Long             'start time of log gen
Dim llEndTime As Long               'end time of log gen
Dim llStartDate As Long             'start date of log gen
Dim llEndDate As Long               'end date of log gen
Dim ilUrfCode As Integer           'user code requesting log
Dim ilVefCode As Integer            ' vehicle to process
Dim llCpfCode As Long               '8-1-16 product/isci code to show isci on L10
Dim llTempTime As Long
Dim ilIsItPolitical As Integer
Dim blInclOnC89 As Boolean          '9-20-16 if Political and C89, ignore the spot to print on log
Dim llWeek As Long

'**** end parameters passed from Log program
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlODF)
        btrDestroy hlODF
        Exit Sub
    End If
    hlSvr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlSvr, "", sgDBPath & "Svr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSvr)
        ilRet = btrClose(hlODF)
        btrDestroy hlSvr
        btrDestroy hlODF
        Exit Sub
    End If
    imSvrRecLen = Len(tmSvr)
    
    hmCif = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hlSvr)
        ilRet = btrClose(hlODF)
        btrDestroy hmCif
        btrDestroy hlSvr
        btrDestroy hlODF
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)

    llStartDate = gDateValue(sgLogStartDate)
    llEndDate = llStartDate + Val(sgLogNoDays) - 1
    llStartOfWk = llStartDate                       'get to Monday start week
    Do While (gWeekDayLong(llStartOfWk)) <> 0
        llStartOfWk = llStartOfWk - 1
    Loop
    'convert to btrieve format
    gPackDateLong llStartOfWk, ilStartofWk(0), ilStartofWk(1)
    llStartTime = CLng(gTimeToCurrency(sgLogStartTime, False))
    llEndTime = CLng(gTimeToCurrency(sgLogEndTime, True)) - 1
    ilUrfCode = Val(sgLogUserCode)                         'user code requesting log
    ilVefCode = igcodes(0)                                  'passed for log function
    ilVpfIndex = -1
    ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)              'determine vehicle options index

    ilUseZone = False                                       'assume not using zones until one is found in the vehicle options table
    ilZones = igZones                                       'save zones requested by user : 0=none, 1 =est, 2= cst, 3 =mst, 4 = pst
    ilLoZone = 1                                            'low loop factor to process zones
    ilHiZone = 4                                            'hi loop factor to process zones
    If ilZones <> 0 Then                                    'user has requested one zone in particular
        ilLoZone = ilZones
        ilHiZone = ilZones
    End If
    If ilVpfIndex >= 0 Then                                  'associated vehicle options record exists
        'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
        If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
            ilUseZone = True
        Else
            'Zones not used, fake out flag to do 1 zone (EST)
            ilZones = 1
            ilLoZone = 1
            ilHiZone = 1
        End If
    Else
        'no vehicle options table
    End If
    For ilZoneLoop = ilLoZone To ilHiZone               'loop on all time zones (or just the selective one,  variety of zones are not allowed)
        ReDim tmTsvr(0 To 0) As SVR                      'prepass image built into memory

        Select Case ilZoneLoop
            Case 1  'Eastern
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(ilRec - 1))
                Next ilRec
                If ilUseZone Then
                    slZone = "EST"
                Else
                    slZone = "   "
                End If
            Case 2  'Central
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "CST"
            Case 3  'Mountain
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "MST"
            Case 4  'Pacific
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "PST"
        End Select
        For llWeek = llStartDate To llEndDate Step 7
            For ilDay = 0 To 6                                'loop on all days of the week
'                slDate = Format$(llStartDate + ilDay, "m/d/yy")
                slDate = Format$(llWeek + ilDay, "m/d/yy")
'                ilDayOfWeek = gWeekDayLong(llStartDate + ilDay)
                ilDayOfWeek = gWeekDayLong(llWeek + ilDay)
                Select Case ilDay
                    Case 0
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 1
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 2
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 3
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 4
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 5
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                    Case 6
                        mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                End Select
                'Loop thru each event and build a unique record based on zone, vehicle, event type, time & position
                'for the entire week
                For ilOdf = LBound(tmOdf0) To UBound(tmOdf0) - 1 Step 1
                    'the ODF generation date & time must match; otherwise retriving old records for same vehicle which causes too many spots
                    '4-8-05 filter out non-matching zones and spots only
                    ilIsItPolitical = gIsItPolitical(tmOdf0(ilOdf).iAdfCode)           '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
                    blInclOnC89 = True
                    If sgRnfRptName = "C89" And ilIsItPolitical = True Then
                        blInclOnC89 = False             'ignore politicals spots on c89 printout
                    End If
                    
                    If tmOdf0(ilOdf).iEtfCode = 0 And (Trim$(slZone) = "" Or Left$(slZone, 1) = Left$(tmOdf0(ilOdf).sZone, 1)) And (blInclOnC89 = True) Then
                    'If tmOdf0(ilOdf).iEtfCode = 0 Then                       'look for only spots
                        gUnpackTimeLong tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1), False, llOdfTime   'dont convert 12m to end of day
                        'Determine daypart index so they will be separated on Crystal report
                        'L89 (copy of L10 with columns removed), added 2/22/19
                        If sgRnfRptName = "L10" Or sgRnfRptName = "L40" Or sgRnfRptName = "C89" Or sgRnfRptName = "L89" Then     '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
                            For ilDPIndex = 1 To 5
                                If llOdfTime >= llZoneEndTimes(ilDPIndex) And llOdfTime < llZoneEndTimes(ilDPIndex + 1) Then
                                    Exit For
                                End If
                            Next ilDPIndex
                        Else
                            For ilDPIndex = 1 To 5
                                If llOdfTime >= llZoneEndTimes(ilDPIndex) And llOdfTime < llZoneEndTimes(ilDPIndex + 1) Then
                                    Exit For
                                End If
                            Next ilDPIndex
                            If llOdfTime >= llZoneEndTimes(6) And llOdfTime <= 86400 Then   'test for the last DP with 12m as end of day
                                ilDPIndex = 6
                            End If
                        End If
                        ilFoundUnique = False
                        ilUpper = UBound(tmTsvr)
                        llCpfCode = mGetPrfFromCif(tmOdf0(ilOdf).lCifCode)
    
                        For ilRec = 0 To ilUpper - 1 Step 1
                            'look for unique entry, if none found create one
                            gUnpackTimeLong tmTsvr(ilRec).iAirTime(0), tmTsvr(ilRec).iAirTime(1), False, llSvrTime   'dont convert 12m to end of day
                            If (llOdfTime = llSvrTime) And (tmOdf0(ilOdf).iPositionNo = tmTsvr(ilRec).iPosition) Then
                                ilFoundUnique = True
                                gPackTimeLong llZoneEndTimes(ilDPIndex), tmTsvr(ilRec).iDPStartTime(0), tmTsvr(ilRec).iDPStartTime(1)
                                tmTsvr(ilRec).sSpotID(ilDayOfWeek) = tmOdf0(ilOdf).sProgCode
                                tmTsvr(ilRec).iBreak(ilDayOfWeek) = tmOdf0(ilOdf).iBreakNo
                                gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "1", True, slLen
                                tmTsvr(ilRec).iLen(ilDayOfWeek) = Val(slLen)
                                tmTsvr(ilRec).iAdfCode(ilDayOfWeek) = tmOdf0(ilOdf).iAdfCode
                                tmTsvr(ilRec).sProduct(ilDayOfWeek) = tmOdf0(ilOdf).sProduct
                                'tmTSvr(ilRec).lRefCode(ilDayOfWeek + 1) = llCpfCode        '8-1-16
                                tmTsvr(ilRec).lRefCode(ilDayOfWeek) = llCpfCode        '8-1-16
                                Exit For
                            End If
                            
                        Next ilRec
                        If Not ilFoundUnique Then
                            ilUpper = UBound(tmTsvr)
                            tmTsvr(ilUpper).iGenDate(0) = igNowDate(0)
                            tmTsvr(ilUpper).iGenDate(1) = igNowDate(1)
                            '10-10-01
                            tmTsvr(ilUpper).lGenTime = lgNowTime
                            'tmTSvr(ilUpper).iGenTime(0) = igNowTime(0)
                            'tmTSvr(ilUpper).iGenTime(1) = igNowTime(1)
                            For ilLoop = 0 To 6             '8-22-01 set fields to blank, otherwise the field is zero
                                tmTsvr(ilUpper).sSpotID(ilLoop) = ""
                            Next ilLoop
                            gPackTimeLong llZoneEndTimes(ilDPIndex), tmTsvr(ilRec).iDPStartTime(0), tmTsvr(ilRec).iDPStartTime(1)    'keep DP separted on Crystal report
                            tmTsvr(ilUpper).iStartofWk(0) = ilStartofWk(0)
                            tmTsvr(ilUpper).iStartofWk(1) = ilStartofWk(1)
                            tmTsvr(ilUpper).iVefCode = ilVefCode
                            tmTsvr(ilUpper).sZone = slZone
                            tmTsvr(ilUpper).iPosition = tmOdf0(ilOdf).iPositionNo
                            tmTsvr(ilUpper).iAirTime(0) = tmOdf0(ilOdf).iLocalTime(0)       'this is the air time
                            tmTsvr(ilUpper).iAirTime(1) = tmOdf0(ilOdf).iLocalTime(1)
                            
                                                   
                            tmTsvr(ilUpper).sSpotID(ilDayOfWeek) = tmOdf0(ilOdf).sProgCode
                            tmTsvr(ilUpper).iBreak(ilDayOfWeek) = tmOdf0(ilOdf).iBreakNo
                            gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "1", True, slLen
                            tmTsvr(ilUpper).iLen(ilDayOfWeek) = Val(slLen)
                            tmTsvr(ilUpper).iAdfCode(ilDayOfWeek) = tmOdf0(ilOdf).iAdfCode
                            tmTsvr(ilUpper).sProduct(ilDayOfWeek) = tmOdf0(ilOdf).sProduct
                            'tmTSvr(ilUpper).lRefCode(ilDayOfWeek + 1) = llCpfCode        '8-1-16
                            tmTsvr(ilUpper).lRefCode(ilDayOfWeek) = llCpfCode        '8-1-16
                            gUnpackTimeLong tmTsvr(ilRec).iAirTime(0), tmTsvr(ilRec).iAirTime(1), False, llTempTime   'convert time for debugging
    
                            '9-12-00
                            tmTsvr(ilUpper).lHd1CefCode = tgVof.lHd1CefCode 'log header comment code
                            tmTsvr(ilUpper).lFt1CefCode = tgVof.lFt1CefCode 'log footer #1 comment code
                            tmTsvr(ilUpper).lFt2CefCode = tgVof.lFt2CefCode 'log footer #2 comment code
                            'setup the DP text
    
                            tmTsvr(ilUpper).sProgramInfo = ""
                            If llZoneEndTimes(ilDPIndex) = 0 Then
                                tmTsvr(ilUpper).sProgramInfo = "6A"
                            ElseIf llZoneEndTimes(ilDPIndex) = 86400 Then
                                tmTsvr(ilUpper).sProgramInfo = ""       'found last entry whose end time is already 12m
                            Else
                                tmTsvr(ilUpper).sProgramInfo = gFormatTimeLong(llZoneEndTimes(ilDPIndex), "A", "1")
                            End If
                            If ilDPIndex <= 5 Then          'valid entry to determine end time since there is only 6 entries in options table
                                tmTsvr(ilUpper).sProgramInfo = Trim$(tmTsvr(ilUpper).sProgramInfo) & "-" & gFormatTimeLong(llZoneEndTimes(ilDPIndex + 1), "A", "1")
                            End If
                            ReDim Preserve tmTsvr(0 To UBound(tmTsvr) + 1) As SVR
                        End If
                    End If                          'tmOdf0(ilOdf).itype = 4
                Next ilOdf
                ReDim tmOdf0(0 To 0) As ODFEXT
            Next ilDay
            'Loop thru all event times built in memory and write one record to disk for each
            'unique zone, veh, event type, time, & position
            For ilRec = 0 To UBound(tmTsvr) - 1 Step 1
                gPackDateLong llWeek, tmTsvr(ilRec).iStartofWk(0), tmTsvr(ilRec).iStartofWk(1)
            
                '8-22-01 adjust for Mon-Fri Program code if program (avail and no spot) doesnt exist on Monday.
                'ie - client has program on Thursday only, and M-F spot ID doesnt showing
                ilFoundUnique = False
                For ilOdf = 0 To 4
                    If Trim$(tmTsvr(ilRec).sSpotID(ilOdf)) <> "" Then
                        ilFoundUnique = True
                        Exit For
                    End If
                Next ilOdf
                If ilFoundUnique And ilOdf > 0 Then
                    tmTsvr(ilRec).sSpotID(0) = tmTsvr(ilRec).sSpotID(ilOdf)
                End If
    
                ilRet = btrInsert(hlSvr, tmTsvr(ilRec), imSvrRecLen, INDEXKEY0)
            Next ilRec
            ReDim tmTsvr(0 To 0) As SVR
        Next llWeek
    Next ilZoneLoop                                         'for ilzone
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hlSvr)
    btrDestroy hmCif
    btrDestroy hlODF
    btrDestroy hlSvr
End Sub
'***********************************************************************************************
'*
'*      Procedure Name:gL36ComlSmry
'*
'*             Created:11/2/99       By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: Create Commerical Summary
'*            into Crystal report (converted from
'*            Bridge)
'
'*      12-13-99 Do not test Chf for Psa or Promo if
'*              contract # missing from ODF
'*
'               dh 7-30-04 match ODF generate date and time when gathering ODF records for prepass
'                           Any old records previously not cleared out where gathered resulting in errorneous log/cp
'*************************************************************************************************
Sub gL36ComlSmry()
    Dim ilRet As Integer
    Dim hlODF As Integer            'One day log file handle
    Dim tlOdf As ODF
    Dim tlOdfSrchKey As ODFKEY0            'ODF record image
    Dim ilOdfRecLen As Integer        'ODF record length
    Dim ilZone As Integer
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim slDates As String
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llTime As Long
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer
    Dim slLen As String
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilVehicle As Integer
    Dim ilVefCode As Integer
    Dim ilShowTotals As Integer
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilVpfIndex As Integer
    Dim ilStartUp As Integer
    Dim slAvailFirstLetter As String
    Dim ilUseZone As Integer
    ReDim llDPStartTime(0 To 4) As Long 'Daypart Start Time for zones
    ReDim llDPEndTime(0 To 4) As Long   'Daypart End Time for zones
    ReDim slValues(0 To 15) As String
    ReDim ilTotals(0 To 12) As Integer
    ReDim tlSort(0 To 0) As CMMLSUM
    ilUpper = 0
    slDate = RptSelLg!edcSelCFrom.Text   'Start date
    ilStartUp = True
    If (slDate = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCFrom.SetFocus
    Exit Sub
    End If
    If Not gValidDate(slDate) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCFrom.SetFocus
    Exit Sub
    End If
    llStartDate = gDateValue(slDate)
    llEndDate = llStartDate + Val(RptSelLg!edcSelCFrom1.Text) - 1
    slDates = Format$(llStartDate, "dddd" & ", " & "m/d/yy") & " To " & Format$(llEndDate, "dddd" & ", " & "m/d/yy")
    slTime = RptSelLg!edcSelCTo.Text   'Start Time
    If (slTime = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo.SetFocus
    Exit Sub
    End If
    If Not gValidTime(slTime) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo.SetFocus
    Exit Sub
    End If
    llStartTime = CLng(gTimeToCurrency(slTime, False))
    slTime = RptSelLg!edcSelCTo1.Text   'End Time
    If (slTime = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo1.SetFocus
    Exit Sub
    End If
    If Not gValidTime(slTime) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo1.SetFocus
    Exit Sub
    End If
    llEndTime = CLng(gTimeToCurrency(slTime, True)) - 1
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hlODF)
    btrDestroy hlODF
    Exit Sub
    End If
    ilOdfRecLen = Len(tlOdf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    btrDestroy hmAdf
    btrDestroy hlODF
    Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    btrDestroy hmVef
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    Screen.MousePointer = vbHourglass
    ilZone = 0
    For ilLoop = 0 To 3 Step 1
    If RptSelLg!ckcSelC3(ilLoop).Value = vbChecked Then
        ilZone = ilZone + 1
    End If
    Next ilLoop
    ilShowTotals = False
    tmVef.iCode = 0
    tmAnf.iCode = 0
    For ilVehicle = 0 To UBound(igcodes) - 1 Step 1
    'If RptSelLg!lbcSelection(0).Selected(ilVehicle) Then
    ReDim tlSort(0 To 0) As CMMLSUM
    ilUpper = 0
    ilVefCode = igcodes(ilVehicle)
    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        tmVef.sName = "Missing"
    End If
    ilVpfIndex = -1
    ilUseZone = False
    'For ilLoop = 0 To UBound(tgVpf) Step 1
    '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
        ilLoop = gBinarySearchVpf(ilVefCode)
        If ilLoop <> -1 Then
            ilVpfIndex = ilLoop
            'If tgVpf(ilVpfIndex).sGZone(1) = "   " Then
            If tgVpf(ilVpfIndex).sGZone(1) = "   " Then
                ilZone = 1                          '1 zone to process
        '        Exit For                            'not using zones
            Else
                ilUseZone = True                    'something defined for zones
        '        Exit For
            End If
        End If
    'Next ilLoop
    'Set titles for week
    For llDate = llStartDate To llEndDate Step 1
        ilDay = gWeekDayLong(llDate)
        'tlOdfSrchKey.iUrfCode = ilUrfCode
        tlOdfSrchKey.iVefCode = ilVefCode
        slDate = Format$(llDate, "m/d/yy")
        gPackDate slDate, tlOdfSrchKey.iAirDate(0), tlOdfSrchKey.iAirDate(1)
        gPackDate slDate, ilAirDate0, ilAirDate1
        slTime = gCurrencyToTime(CCur(llStartTime))
        gPackTime slTime, tlOdfSrchKey.iLocalTime(0), tlOdfSrchKey.iLocalTime(1)
        tlOdfSrchKey.sZone = " "
        tlOdfSrchKey.iSeqNo = 0
        ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlOdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
        'Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = ilUrfCode) And (tlOdf.iVefCode = ilVefCode) And (tlOdf.iAirDate(0) = ilAirDate0) And (tlOdf.iAirDate(0) = ilAirDate0)
        Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iVefCode = ilVefCode) And (tlOdf.iAirDate(0) = ilAirDate0) And (tlOdf.iAirDate(1) = ilAirDate1)
        'the ODF generation date & time must match; otherwise retriving old records for same vehicle which causes too many spots
        If (tlOdf.iMnfSubFeed = 0) And tlOdf.lGenTime = lgGenTime And tlOdf.iGenDate(0) = igODFGenDate(0) And tlOdf.iGenDate(1) = igODFGenDate(1) Then   'bypass records with subfeed
            If (tlOdf.iEtfCode = 0) Then
            slAvailFirstLetter = "L"
            If tmAnf.iCode <> tlOdf.ianfCode Then
                If tlOdf.ianfCode <> 0 Then
                tmAnfSrchKey.iCode = tlOdf.ianfCode
                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    slAvailFirstLetter = Left$(Trim$(tmAnf.sName), 1)
                End If
                End If
            Else
                slAvailFirstLetter = Left$(Trim$(tmAnf.sName), 1)
            End If
            If (slAvailFirstLetter = "N") Or (sgRnfRptName = "L88") Then   '8-10-16 bypass all avails except Network avails unless its L88, it needs to show PSAs
                'Bypass PSA and Promo contracts
                If tlOdf.lCntrNo = 0 Then
                    tmChf.lCntrNo = 0
                    tmChf.sType = "O"           'set to order
                    tmChf.sSchStatus = "F"      'fully scheduled
                    ilRet = 0
                Else
                    tmChfSrchKey1.lCntrNo = tlOdf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                End If
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlOdf.lCntrNo) And (tmChf.sSchStatus = "A")
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
               ' If (ilRet = BTRV_ERR_NONE) And (tmChf.sType <> "S") And (tmChf.sType <> "M") And (tmChf.lCntrNo = tlOdf.lCntrNo) Then
               '8-10-16 for L88, allow psas
                If (ilRet = BTRV_ERR_NONE) And (((tmChf.sType <> "S") And (tmChf.sType <> "M") And (tmChf.lCntrNo = tlOdf.lCntrNo) And (sgRnfRptName <> "L88")) Or ((tmChf.sType <> "M") And (tmChf.lCntrNo = tlOdf.lCntrNo) And (sgRnfRptName = "L88"))) Then

                ilZone = -1
                If Not ilUseZone Then
                    tlOdf.sZone = "EST"
                End If
                Select Case Trim$(tlOdf.sZone)
                    Case "EST"
                    ilZone = 0
                    llDPStartTime(0) = 0
                    'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) - 1 '21599    '6am
                    'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1))
                    'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) - 1 '35999
                    'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) '36000  '10Am
                    'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) - 1 '53999
                    'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) '54000  '3pm
                    'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) - 1 '68399
                    'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) '68400  '7pm
                    'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(5)) - 1 '86399
                    
                    llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(0)) - 1 '21599    '6am
                    llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(0))
                    llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) - 1 '35999
                    llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(1)) '36000  '10Am
                    llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) - 1 '53999
                    llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(2)) '54000  '3pm
                    llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) - 1 '68399
                    llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(3)) '68400  '7pm
                    llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(4)) - 1 '86399
                    
                    Case "CST"
                    ilZone = 1
                    llDPStartTime(0) = 0
                    'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) - 1 '21599    '6am
                    'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1))
                    'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) - 1 '35999
                    'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) '36000  '10Am
                    'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) - 1 '53999
                    'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) '54000  '3pm
                    'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) - 1 '68399
                    'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) '68400  '7pm
                    'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(5)) - 1 '86399
                    
                    llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(0)) - 1 '21599    '6am
                    llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(0))
                    llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) - 1 '35999
                    llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(1)) '36000  '10Am
                    llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) - 1 '53999
                    llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(2)) '54000  '3pm
                    llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) - 1 '68399
                    llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(3)) '68400  '7pm
                    llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(4)) - 1 '86399
                    
                    Case "MST"
                    ilZone = 2
                    llDPStartTime(0) = 0
                    'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) - 1 '21599    '6am
                    'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1))
                    'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) - 1 '35999
                    'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) '36000  '10Am
                    'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) - 1 '53999
                    'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) '54000  '3pm
                    'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) - 1 '68399
                    'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) '68400  '7pm
                    'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(5)) - 1 '86399
                    
                    llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(0)) - 1 '21599    '6am
                    llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(0))
                    llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) - 1 '35999
                    llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(1)) '36000  '10Am
                    llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) - 1 '53999
                    llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(2)) '54000  '3pm
                    llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) - 1 '68399
                    llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(3)) '68400  '7pm
                    llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(4)) - 1 '86399
                    
                    Case "PST"
                    ilZone = 3
                    llDPStartTime(0) = 0
                    'llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) - 1 '21599    '6am
                    'llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1))
                    'llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) - 1 '35999
                    'llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) '36000  '10Am
                    'llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) - 1 '53999
                    'llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) '54000  '3pm
                    'llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) - 1 '68399
                    'llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) '68400  '7pm
                    'llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(5)) - 1 '86399
                
                    llDPEndTime(0) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(0)) - 1 '21599    '6am
                    llDPStartTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(0))
                    llDPEndTime(1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) - 1 '35999
                    llDPStartTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(1)) '36000  '10Am
                    llDPEndTime(2) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) - 1 '53999
                    llDPStartTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(2)) '54000  '3pm
                    llDPEndTime(3) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) - 1 '68399
                    llDPStartTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(3)) '68400  '7pm
                    llDPEndTime(4) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(4)) - 1 '86399
                
                End Select
                If ilZone >= 0 Then
                    For ilTest = 1 To 4 Step 1
                    If (StrComp(tgVpf(ilVpfIndex).sMapZone(ilTest - 1), Trim$(tlOdf.sZone), 1) = 0) And (StrComp(tgVpf(ilVpfIndex).sMapProgCode(ilTest - 1), tlOdf.sProgCode, 1) = 0) Then
                        gUnpackTime tlOdf.iLocalTime(0), tlOdf.iLocalTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, False))
                        llDPEndTime(tgVpf(ilVpfIndex).iMapDPNo(ilTest - 1) - 2) = llTime - 1
                        llDPStartTime(tgVpf(ilVpfIndex).iMapDPNo(ilTest - 1) - 1) = llTime '9Am
                    End If
                    Next ilTest
                    If RptSelLg!ckcSelC3(ilZone).Value = vbChecked Then
                    gUnpackTime tlOdf.iLocalTime(0), tlOdf.iLocalTime(1), "A", "1", slTime
                    llTime = CLng(gTimeToCurrency(slTime, False))
                    If (llTime >= llDPStartTime(1)) And (llTime <= llDPEndTime(4)) Then  'after 6am and before 12M
                        gUnpackLength tlOdf.iLen(0), tlOdf.iLen(1), "1", True, slLen
                        ilLen = Val(slLen)
                        Do While Len(slLen) < 3
                        slLen = "0" & slLen
                        Loop
                        ilFound = -1
                        For ilTest = 0 To ilUpper - 1 Step 1
'                            If tgSpf.sUseProdSptScr = "P" Then      '4-18-17 use advt product (vs short title
'                                If (tlSort(ilTest).iVefCode = tlOdf.iVefCode) And (StrComp(Trim$(tlSort(ilTest).sZone), Trim$(tlOdf.sZone), 1) = 0) And (tlSort(ilTest).iAdfCode = tlOdf.iAdfCode) And (Trim$(tlSort(ilTest).sShortTitle) = Trim$(tlOdf.sShortTitle)) And (tlSort(ilTest).iLen = ilLen) Then
'                                    ilFound = ilTest
'                                    Exit For
'                                End If
'                            Else
'                                If (tlSort(ilTest).iVefCode = tlOdf.iVefCode) And (StrComp(Trim$(tlSort(ilTest).sZone), Trim$(tlOdf.sZone), 1) = 0) And (tlSort(ilTest).iAdfCode = tlOdf.iAdfCode) And (Trim$(tlSort(ilTest).sProduct) = Trim$(tlOdf.sProduct)) And (tlSort(ilTest).iLen = ilLen) Then
'                                    ilFound = ilTest
'                                    Exit For
'                                End If
'                            End If
                            '4-19-17 change this version of the log to show 1 line per advt & contract; not by the copy products if multiple defined
                            If (tlSort(ilTest).iVefCode = tlOdf.iVefCode) And (StrComp(Trim$(tlSort(ilTest).sZone), Trim$(tlOdf.sZone), 1) = 0) And (tlSort(ilTest).iAdfCode = tlOdf.iAdfCode) And (Trim$(tlSort(ilTest).sProduct) = Trim$(tmChf.sProduct)) And (tlSort(ilTest).iLen = ilLen) Then
                                ilFound = ilTest
                                Exit For
                            End If
                        Next ilTest
                        If ilFound = -1 Then
                        ilFound = ilUpper
                        ilUpper = ilUpper + 1
                        ReDim Preserve tlSort(0 To ilUpper) As CMMLSUM
                        If tlOdf.iAdfCode <> tmAdf.iCode Then
                            tmAdfSrchKey.iCode = tlOdf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                            If ilRet <> BTRV_ERR_NONE Then
                            tmAdf.sName = "Missing"
                            End If
                        End If
                        slStr = Trim$(str$(ilZone))
                        If slStr = "" Then
                            slStr = "0"
                        End If
                        'tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tlOdf.sProduct & slLen
'                        If tgSpf.sUseProdSptScr = "P" Then      '4-18-17 use advt product (vs short title
'                            tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tlOdf.sShortTitle & slLen
'                        Else
'                            tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tlOdf.sProduct & slLen
'                        End If
                        '4-19-17 change this version of the log to show 1 line per advt & contract; not by the copy products if multiple defined
                        tlSort(ilFound).sKey = tmVef.sName & slStr & tlOdf.sZone & tmAdf.sName & tmChf.sProduct & slLen

                        tlSort(ilFound).iVefCode = tlOdf.iVefCode
                        tlSort(ilFound).sVehicle = Trim$(tmVef.sName)
                        tlSort(ilFound).sZone = tlOdf.sZone
                        tlSort(ilFound).sAdvt = Trim$(tmAdf.sName)
                        tlSort(ilFound).iAdfCode = tlOdf.iAdfCode
                        'tlSort(ilFound).sProduct = Trim$(tlOdf.sProduct)
                        '4-19-17 change this version of the log to show 1 line per advt & contract; not by the copy products if multiple defined
                        tlSort(ilFound).sProduct = Trim$(tmChf.sProduct)
                        tlSort(ilFound).sShortTitle = Trim$(tlOdf.sShortTitle)   'Trim$(tlOdf.sProduct)
                        tlSort(ilFound).iLen = ilLen
                        tlSort(ilFound).iMFAM = 0
                        tlSort(ilFound).iSaAM = 0
                        tlSort(ilFound).iSuAM = 0
                        tlSort(ilFound).iMFMid = 0
                        tlSort(ilFound).iSaMid = 0
                        tlSort(ilFound).iSuMid = 0
                        tlSort(ilFound).iMFPM = 0
                        tlSort(ilFound).iSaPM = 0
                        tlSort(ilFound).iSuPM = 0
                        tlSort(ilFound).iMFEve = 0
                        tlSort(ilFound).iSaEve = 0
                        tlSort(ilFound).iSuEve = 0
                        tlSort(ilFound).iTotal = 0
                        End If
                        If llTime < llDPStartTime(2) Then      '6-10am
                        If ilDay <= 4 Then  'M-F
                            tlSort(ilFound).iMFAM = tlSort(ilFound).iMFAM + 1
                        ElseIf ilDay = 5 Then   'Sa
                            tlSort(ilFound).iSaAM = tlSort(ilFound).iSaAM + 1
                        Else    'Sun
                            tlSort(ilFound).iSuAM = tlSort(ilFound).iSuAM + 1
                        End If
                        ElseIf llTime < llDPStartTime(3) Then  '10-3pm
                        If ilDay <= 4 Then  'M-F
                            tlSort(ilFound).iMFMid = tlSort(ilFound).iMFMid + 1
                        ElseIf ilDay = 5 Then   'Sa
                            tlSort(ilFound).iSaMid = tlSort(ilFound).iSaMid + 1
                        Else    'Sun
                            tlSort(ilFound).iSuMid = tlSort(ilFound).iSuMid + 1
                        End If
                        ElseIf llTime < llDPStartTime(4) Then  '3-7pm
                        If ilDay <= 4 Then  'M-F
                            tlSort(ilFound).iMFPM = tlSort(ilFound).iMFPM + 1
                        ElseIf ilDay = 5 Then   'Sa
                            tlSort(ilFound).iSaPM = tlSort(ilFound).iSaPM + 1
                        Else    'Sun
                            tlSort(ilFound).iSuPM = tlSort(ilFound).iSuPM + 1
                        End If
                        Else
                        If ilDay <= 4 Then  'M-F
                            tlSort(ilFound).iMFEve = tlSort(ilFound).iMFEve + 1
                        ElseIf ilDay = 5 Then   'Sa
                            tlSort(ilFound).iSaEve = tlSort(ilFound).iSaEve + 1
                        Else    'Sun
                            tlSort(ilFound).iSuEve = tlSort(ilFound).iSuEve + 1
                        End If
                        End If
                        tlSort(ilFound).iTotal = tlSort(ilFound).iTotal + 1
                        tlSort(ilFound).iDay(ilDay) = 1
                    End If
                    End If
                End If
                End If
            End If
            End If
        End If
        ilRet = btrGetNext(hlODF, tlOdf, ilOdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next llDate
    'Loop thru tlsort and create one disk record for entry
    For ilIndex = LBound(tlSort) To UBound(tlSort) - 1 Step 1
        Select Case Trim$(tlSort(ilIndex).sZone)
        Case "EST"
            If ilUseZone Then
            tmGrf.sBktType = "E"
            Else
            tmGrf.sBktType = " "
            End If
        Case "CST"
            tmGrf.sBktType = "C"
        Case "MST"
            tmGrf.sBktType = "M"
        Case "PST"
            tmGrf.sBktType = "P"
        End Select
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        '10-10-01
        tmGrf.lGenTime = lgNowTime
        'tmGrf.iGenTime(0) = igNowTime(0)
        'tmGrf.iGenTime(1) = igNowTime(1)
        tmGrf.iVefCode = tlSort(ilIndex).iVefCode
        tmGrf.iAdfCode = tlSort(ilIndex).iAdfCode
        '4-19-17 change this version of the log to show 1 line per advt & contract; not by the copy products if multiple defined
'        If tgSpf.sUseProdSptScr = "P" Then      'use short title vs advt/prod
'            tmGrf.sGenDesc = Trim$(tlSort(ilIndex).sShortTitle)
'        Else
            tmGrf.sGenDesc = Trim$(tlSort(ilIndex).sProduct)
'        End If
        tmGrf.iCode2 = tlSort(ilIndex).iLen
        '8-10-16 for new customized log, set up the header and footer comments
        If sgRnfRptName = "L88" Then
            tmGrf.lLong = tgVof.lHd1CefCode
            tmGrf.lCode4 = tgVof.lFt1CefCode
            tmGrf.lChfCode = tgVof.lFt2CefCode
        End If
        
        'loop to setup the spot counts per day and daypart
'        tmGrf.lDollars(1) = tlSort(ilIndex).iMFEarly
'        tmGrf.lDollars(2) = tlSort(ilIndex).iMFAM
'        tmGrf.lDollars(3) = tlSort(ilIndex).iMFMid
'        tmGrf.lDollars(4) = tlSort(ilIndex).iMFPM
'        tmGrf.lDollars(5) = tlSort(ilIndex).iMFEve
'        tmGrf.lDollars(6) = tlSort(ilIndex).iSaEarly
'        tmGrf.lDollars(7) = tlSort(ilIndex).iSaAM
'        tmGrf.lDollars(8) = tlSort(ilIndex).iSaMid
'        tmGrf.lDollars(9) = tlSort(ilIndex).iSaPM
'        tmGrf.lDollars(10) = tlSort(ilIndex).iSaEve
'        tmGrf.lDollars(11) = tlSort(ilIndex).iSuEarly
'        tmGrf.lDollars(12) = tlSort(ilIndex).iSuAM
'        tmGrf.lDollars(13) = tlSort(ilIndex).iSuMid
'        tmGrf.lDollars(14) = tlSort(ilIndex).iSuPM
'        tmGrf.lDollars(15) = tlSort(ilIndex).iSuEve
        tmGrf.lDollars(0) = tlSort(ilIndex).iMFEarly
        tmGrf.lDollars(1) = tlSort(ilIndex).iMFAM
        tmGrf.lDollars(2) = tlSort(ilIndex).iMFMid
        tmGrf.lDollars(3) = tlSort(ilIndex).iMFPM
        tmGrf.lDollars(4) = tlSort(ilIndex).iMFEve
        tmGrf.lDollars(5) = tlSort(ilIndex).iSaEarly
        tmGrf.lDollars(6) = tlSort(ilIndex).iSaAM
        tmGrf.lDollars(7) = tlSort(ilIndex).iSaMid
        tmGrf.lDollars(8) = tlSort(ilIndex).iSaPM
        tmGrf.lDollars(9) = tlSort(ilIndex).iSaEve
        tmGrf.lDollars(10) = tlSort(ilIndex).iSuEarly
        tmGrf.lDollars(11) = tlSort(ilIndex).iSuAM
        tmGrf.lDollars(12) = tlSort(ilIndex).iSuMid
        tmGrf.lDollars(13) = tlSort(ilIndex).iSuPM
        tmGrf.lDollars(14) = tlSort(ilIndex).iSuEve
        For ilLoop = 0 To 6
            'tmGrf.iPerGenl(ilLoop + 1) = tlSort(ilIndex).iDay(ilLoop)
            tmGrf.iPerGenl(ilLoop) = tlSort(ilIndex).iDay(ilLoop)
        Next ilLoop

        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    Next ilIndex
    Next ilVehicle

    Erase llDPStartTime
    Erase llDPEndTime
    Erase slValues
    Erase ilTotals
    Erase tlSort
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hmGrf)
    btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmAnf
    btrDestroy hmAdf
    btrDestroy hlODF
    btrDestroy hmGrf
    Exit Sub
    Return
End Sub
'*******************************************************
'*
'*      Procedure Name:gL37ComlSch
'*
'*             Created:6/16/93       By:D. LeVine
'*            Modified:12/14/97      By:d. Hosaka
'*
'*       Comments: Generate Seven day log report
'*       12/14/97 New log generation scheme processes
'*       only 1 vehicle at a time automatically
'*       from the main log screen
'*      3/9/99 CommlSchedules are generated forL07
'*          and L27 (new event ids)
'*      11-3-99 Convert ABC coml Schedule (bridge rept)
'*              to Crystal
'*      5-24-01 Create another version of coml sch (c80)
'*              that doesn't show pgm ID and positions/breaks
'*      5-30-01 C80 didnt show spots when program time was optioned
'               not to be shown
'       6-19-01 Change C80 to show pgm name/time on same line as spot,
'               not on a separate line.  Remove pgm length & avail event ID
'       7-9-03 Start date of week was not set properly when multiple weeks requested
'       5-11-05 event comments are printed in the footer (special for ABC); if
'               the length of char exceeds 250, then truncate it.
'*******************************************************
Sub gL37ComlSch(ilUrfCode As Integer)
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilZone As Integer
    Dim slZone As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim hlODF As Integer            'One day log file handle
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim ilAnyEvt As Integer
    Dim ilFind As Integer       '0=Prog; 1=Note, 2=Break
    Dim ilPass As Integer
    Dim ilBreakNo As Integer
    Dim ilPositionNo As Integer
    Dim llCurrTime As Long
    Dim ilMode As Integer
    Dim ilVehCount As Integer
    Dim ilZCount As Integer
    Dim ilDBEof As Integer    'end of weeks processing
    Dim hlSvr As Integer
    ReDim ilEvtIndex(0 To 6) As Integer
    ReDim ilSvEvtIndex(0 To 6) As Integer
    ReDim ilPrgFnd(0 To 6) As Integer
    Dim slEvtTitle As String
    'ReDim slComment(1 To 3) As String
    ReDim slComment(0 To 3) As String       'Index zero ignored
    'ReDim slTempComment(1 To 3) As String
    ReDim slTempComment(0 To 3) As String   'Index zero ignored
    Dim slNewComment As String
    Dim ilFound As Integer
    Dim ilComm As Integer
    Dim ilTempComm As Integer
    Dim ilFirstZoneIndex As Integer
    ReDim slEvt(0 To 6) As String 'Events to be printed
    'ReDim tlOdf0(1 To 1) As ODFEXT
    'ReDim tlOdf1(1 To 1) As ODFEXT
    'ReDim tlOdf2(1 To 1) As ODFEXT
    'ReDim tlOdf3(1 To 1) As ODFEXT
    'ReDim tlOdf4(1 To 1) As ODFEXT
    'ReDim tlOdf5(1 To 1) As ODFEXT
    'ReDim tlOdf6(1 To 1) As ODFEXT
    ReDim tlOdf0(0 To 0) As ODFEXT
    ReDim tlOdf1(0 To 0) As ODFEXT
    ReDim tlOdf2(0 To 0) As ODFEXT
    ReDim tlOdf3(0 To 0) As ODFEXT
    ReDim tlOdf4(0 To 0) As ODFEXT
    ReDim tlOdf5(0 To 0) As ODFEXT
    ReDim tlOdf6(0 To 0) As ODFEXT
    Dim ilLastBreakNo As Integer
    Dim slVehName As String
    Dim ilUseZone As Integer
    Dim ilVpfIndex As Integer
    Dim ilLoop As Integer
    Dim ilMajorSort As Integer      'key field used for Crystal reporting
    Dim ilCount As Integer
    Dim ilPgmLength As Integer      'length of pgm descrip string
    Dim ilPgmIndex As Integer       'pos of pgm description string
    Dim llStartOfWk   As Long
    ReDim ilSpotLength(0 To 6) As Integer      'length of spot descrip string for m-su
    ReDim ilSpotIndex(0 To 6) As Integer       'pos. of spot description string for m-su
    ReDim ilStartofWk(0 To 1) As Integer
    ReDim slLogComment(0 To 0) As String
    ReDim slSegment(0 To 6) As String * 30
    Dim slSegmentStr As String
    Dim slTemp As String
    slLogComment(0) = ""
    smLogName = UCase$(Trim$(sgRnfRptName))
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    tmVefSrchKey.iCode = igcodes(0)         'new scheme only processes one vehicle at a time in this module
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
    slVehName = tmVef.sName
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    Else
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    Exit Sub
    End If
    slDate = RptSelLg!edcSelCFrom.Text   'Start date
    If (slDate = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCFrom.SetFocus
    Exit Sub
    End If
    If Not gValidDate(slDate) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCFrom.SetFocus
    Exit Sub
    End If
    llStartDate = gDateValue(slDate)
    llEndDate = llStartDate + Val(RptSelLg!edcSelCFrom1.Text) - 1
    slTime = RptSelLg!edcSelCTo.Text   'Start Time
    If (slTime = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo.SetFocus
    Exit Sub
    End If
    If Not gValidTime(slTime) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo.SetFocus
    Exit Sub
    End If
    llStartTime = CLng(gTimeToCurrency(slTime, False))
    slTime = RptSelLg!edcSelCTo1.Text   'End Time
    If (slTime = "") Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo1.SetFocus
    Exit Sub
    End If
    If Not gValidTime(slTime) Then
    igGenRpt = False
    RptSelLg!frcOutput.Enabled = igOutput
    RptSelLg!frcCopies.Enabled = igCopies
    'RptSelLg!frcWhen.Enabled = igWhen
    RptSelLg!frcFile.Enabled = igFile
    RptSelLg!frcOption.Enabled = igOption
    'RptSelLg!frcRptType.Enabled = igReportType
    Beep
    RptSelLg!edcSelCTo1.SetFocus
    Exit Sub
    End If
    llEndTime = CLng(gTimeToCurrency(slTime, True)) - 1
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hlODF)
    btrDestroy hlODF
    Exit Sub
    End If
    hmEnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmEnf)
    ilRet = btrClose(hlODF)
    btrDestroy hmEnf
    btrDestroy hlODF
    Exit Sub
    End If
    imEnfRecLen = Len(tmEnf)
    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCef)
    ilRet = btrClose(hmEnf)
    ilRet = btrClose(hlODF)
    btrDestroy hmCef
    btrDestroy hmEnf
    btrDestroy hlODF
    Exit Sub
    End If
    imCefRecLen = Len(tmCef)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmCef)
    ilRet = btrClose(hmEnf)
    ilRet = btrClose(hlODF)
    btrDestroy hmCef
    btrDestroy hmAnf
    btrDestroy hmEnf
    btrDestroy hlODF
    Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)
    hlSvr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlSvr, "", sgDBPath & "Svr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmCef)
    ilRet = btrClose(hmEnf)
    ilRet = btrClose(hlODF)
    btrDestroy hmCef
    btrDestroy hmAnf
    btrDestroy hmEnf
    btrDestroy hlODF
    Exit Sub
    End If
    imSvrRecLen = Len(tmSvr)
    llStartDate = gDateValue(sgLogStartDate)
    llEndDate = llStartDate + Val(sgLogNoDays) - 1
    llStartOfWk = llStartDate                       'get to Monday start week
    Do While (gWeekDayLong(llStartOfWk)) <> 0
    llStartOfWk = llStartOfWk - 1
    Loop
    'convert to btrieve format
    gPackDateLong llStartOfWk, ilStartofWk(0), ilStartofWk(1)
    tmEnf.iCode = 0
    tmSEnf.iCode = 0
    tmAnf.iCode = 0
    ilZCount = 0
    ilFirstZoneIndex = -1
    For ilZone = 0 To 3 Step 1
    If RptSelLg!ckcSelC3(ilZone).Value = vbChecked Then
        ilZCount = ilZCount + 1
        If ilFirstZoneIndex = -1 Then
        ilFirstZoneIndex = ilZone
        End If
    End If
    Next ilZone
    ilVehCount = 1
    ilDBEof = False
    slComment(1) = ""
    slComment(2) = ""
    slComment(3) = ""
    For ilVehicle = 0 To UBound(igcodes) - 1 Step 1
    ilVefCode = igcodes(ilVehicle)

    ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)
    ilUseZone = False
    If ilVpfIndex > 0 Then
        'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
        If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
            ilUseZone = True
        Else
            'Zones not used,
            RptSelLg!ckcSelC3(0).Value = vbChecked  'True            'fake out to do just one
            RptSelLg!ckcSelC3(1).Value = vbUnchecked    'False
            RptSelLg!ckcSelC3(2).Value = vbUnchecked    'False
            RptSelLg!ckcSelC3(3).Value = vbUnchecked    'False
        End If
    End If
    'Set titles for week
    For ilZone = 0 To 3 Step 1
        If RptSelLg!ckcSelC3(ilZone).Value = vbChecked Then
        'set fields in prepass file that dont change
        tmSvr.iGenDate(0) = igNowDate(0)
        tmSvr.iGenDate(1) = igNowDate(1)
        '10-10-01
        tmSvr.lGenTime = lgNowTime
        'tmSvr.iGenTime(0) = igNowTime(0)
        'tmSvr.iGenTime(1) = igNowTime(1)
        tmSvr.iVefCode = ilVefCode
        ilMajorSort = 0         'initializekeys for prepass record
        For llDate = llStartDate To llEndDate Step 7    'One week at a time
            ilDBEof = False
            llCurrTime = 86401
            Select Case ilZone
            Case 0  'Eastern
                If ilUseZone Then
                slZone = "EST"
                Else
                slZone = "   "
                End If
            Case 1  'Central
                slZone = "CST"
            Case 2  'Mountain
                slZone = "MST"
            Case 3  'Pacific
                slZone = "PST"
            End Select
            If ilZone = ilFirstZoneIndex Then
            For ilDay = 0 To 6 Step 1
                slDate = Format$(llDate + ilDay, "m/d/yy")
                Select Case ilDay
                Case 0
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0()
                Case 1
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf1()
                Case 2
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf2()
                Case 3
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf3()
                Case 4
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf4()
                Case 5
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf5()
                Case 6
                    mObtainOdf hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf6()
                End Select
            Next ilDay
            End If
            For ilDay = 0 To 6 Step 1
            Select Case ilDay
                Case 0
                mMoveODF slZone, tmOdf0(), tlOdf0()
                ilEvtIndex(0) = LBound(tlOdf0)
                Case 1
                mMoveODF slZone, tmOdf1(), tlOdf1()
                ilEvtIndex(1) = LBound(tlOdf1)
                Case 2
                mMoveODF slZone, tmOdf2(), tlOdf2()
                ilEvtIndex(2) = LBound(tlOdf2)
                Case 3
                mMoveODF slZone, tmOdf3(), tlOdf3()
                ilEvtIndex(3) = LBound(tlOdf3)
                Case 4
                mMoveODF slZone, tmOdf4(), tlOdf4()
                ilEvtIndex(4) = LBound(tlOdf4)
                Case 5
                mMoveODF slZone, tmOdf5(), tlOdf5()
                ilEvtIndex(5) = LBound(tlOdf5)
                Case 6
                mMoveODF slZone, tmOdf6(), tlOdf6()
                ilEvtIndex(6) = LBound(tlOdf6)
            End Select
            Next ilDay
            Do While (Not ilDBEof)
            slNewComment = ""
            slComment(1) = ""
            slComment(2) = ""
            slComment(3) = ""
            slTempComment(1) = ""
            slTempComment(2) = ""
            slTempComment(3) = ""
            'Find The next earliest time
            slTempComment(1) = ""
            slTempComment(2) = ""
            slTempComment(3) = ""
            ilAnyEvt = False
            For ilDay = 0 To 6 Step 1
                ilPrgFnd(ilDay) = False
                slEvt(ilDay) = ""
                Select Case ilDay
                Case 0
                    mFindEarliestLogPrg llCurrTime, tlOdf0(), ilEvtIndex(ilDay), ilAnyEvt
                Case 1
                    mFindEarliestLogPrg llCurrTime, tlOdf1(), ilEvtIndex(ilDay), ilAnyEvt
                Case 2
                    mFindEarliestLogPrg llCurrTime, tlOdf2(), ilEvtIndex(ilDay), ilAnyEvt
                Case 3
                    mFindEarliestLogPrg llCurrTime, tlOdf3(), ilEvtIndex(ilDay), ilAnyEvt
                Case 4
                    mFindEarliestLogPrg llCurrTime, tlOdf4(), ilEvtIndex(ilDay), ilAnyEvt
                Case 5
                    mFindEarliestLogPrg llCurrTime, tlOdf5(), ilEvtIndex(ilDay), ilAnyEvt
                Case 6
                    mFindEarliestLogPrg llCurrTime, tlOdf6(), ilEvtIndex(ilDay), ilAnyEvt
                End Select
            Next ilDay
            If Not ilAnyEvt Then
                ilDBEof = True
                Exit Do
            End If
            'Obtain number of break for matching times
            For ilIndex = 0 To 6 Step 1
                ilSvEvtIndex(ilIndex) = ilEvtIndex(ilIndex)
            Next ilIndex
            ilAnyEvt = False
            ilFind = 0
            ilPass = 0
            ilBreakNo = 1
            ilPositionNo = 1
            ilLastBreakNo = 1
            ilMode = 0  'Test
            Do
                Do
                    ilAnyEvt = False
                    For ilDay = 0 To 6 Step 1
                        Select Case ilDay
                        Case 0
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf0(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 1
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf1(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 2
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf2(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 3
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf3(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 4
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf4(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 5
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf5(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        Case 6
                            mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf6(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
                        End Select
                        If ilMode = 1 Then
                            If slNewComment <> "" Then
                                ilFound = False
                                For ilComm = LBound(slLogComment) To UBound(slLogComment) - 1 Step 1
                                    If StrComp(Trim$(slNewComment), Trim$(slLogComment(ilComm))) = 0 Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilComm
                                If Not ilFound Then
                                    slLogComment(UBound(slLogComment)) = slNewComment
                                    ReDim Preserve slLogComment(0 To UBound(slLogComment) + 1) As String
                                End If

                                For ilComm = 1 To 3 Step 1
                                If slTempComment(ilComm) = "" Then
                                    slTempComment(ilComm) = slNewComment
                                    Exit For
                                End If
                                If StrComp(slTempComment(ilComm), slNewComment, 1) = 0 Then
                                    Exit For
                                End If
                                Next ilComm
                            End If
                        End If
                    Next ilDay
                    'process events at same time/break
                    If Not ilAnyEvt Then
                        Exit Do
                    End If
                    If ilMode = 1 Then
                        If ilPass = 0 Then
                            Exit Do
                        ElseIf ilPass = 1 Then
                            ilMode = 0  'Another note
                        ElseIf ilPass = 2 Then
                            ilMode = 0  'Another spot with this break no
                            ilPositionNo = ilPositionNo + 1
                        ElseIf ilPass = 3 Then
                            ilMode = 0  'Another spot with this break no
                        End If
                    Else
                        ilMode = 1  'Add events
                        If ilPass = 0 Then
                            slEvtTitle = gCurrencyToTime(CCur(llCurrTime)) & Chr$(10)
                        ElseIf (ilPass = 1) Or (ilPass = 3) Then
                            slEvtTitle = slEvtTitle & Chr$(10)
                        ElseIf ilPass = 2 Then
                            slStr = Trim$(str$(ilBreakNo))
                            If Len(slStr) = 1 Then
                                slStr = "0" & slStr
                            End If
                            If StrComp(smLogName, "L27", 1) <> 0 And StrComp(smLogName, "L37", 1) <> 0 And StrComp(smLogName, "L39", 1) <> 0 Then    '*************
                                slEvtTitle = slEvtTitle & Chr$(10) & slStr & Trim$(str$(ilPositionNo))
                            Else
                                If ilPositionNo = 1 Then
                                    slEvtTitle = slEvtTitle & Chr$(10) & Chr$(10) & slStr & Trim$(str$(ilPositionNo))
                                Else
                                    slEvtTitle = slEvtTitle & Chr$(10) & slStr & Trim$(str$(ilPositionNo))
                                End If
                            End If
                        End If
                    End If
                Loop
                ilMode = 0
                If ilPass = 0 Then
                    ilPass = 1
                ElseIf ilPass = 1 Then
                    ilPass = 2
                ElseIf ilPass = 2 Then
                    If (ilPositionNo = 1) Then
                        If ilBreakNo >= ilLastBreakNo + 6 Then
                            ilPass = 3
                        Else
                            ilBreakNo = ilBreakNo + 1
                            ilPositionNo = 1
                        End If
                    Else
                        ilLastBreakNo = ilBreakNo
                        ilBreakNo = ilBreakNo + 1
                        ilPositionNo = 1
                    End If
                ElseIf ilPass = 3 Then
                    Exit Do
                End If
            Loop

            If right$(slEvtTitle, 1) <> Chr$(10) Then
                slEvtTitle = slEvtTitle & Chr$(10)
            End If
            For ilDay = 0 To 6 Step 1
                If right$(slEvt(ilDay), 1) <> Chr$(10) Then
                    slEvt(ilDay) = slEvt(ilDay) & Chr$(10)
                End If
            Next ilDay
            slEvtTitle = slEvtTitle & Chr$(0)



            'Create records in prepass record- 7 days across, an avail at a time
            ilLoop = Len(slEvtTitle)     'determine loop in the event title string
            ilCount = 0
            tmSvr.iSeq = 0
            ilMajorSort = ilMajorSort + 1
            For ilIndex = 1 To ilLoop
                ilRet = InStr(ilIndex, slEvtTitle, Chr$(10))
                If ilRet = 0 Then
                    Exit For
                Else
                    ilCount = ilCount + 1
                    ilIndex = ilRet     'set the current loop to the loc of the carriage return for next set
                End If
            Next ilIndex

            'ilCount contains # of lines in this avail
            ilPgmIndex = 1
            For ilDay = 0 To 6
                ilSpotIndex(ilDay) = 1      'set start loc to 1 for each day
            Next ilDay
            For ilIndex = 1 To ilCount + 1
                ilPgmLength = InStr(ilPgmIndex, slEvtTitle, Chr$(10))
                tmSvr.sProgramInfo = ""
                If ilPgmLength > 0 Then
                    If StrComp(smLogName, "C80", 1) = 0 Then        'c80 doesnt may not show break/positions
                        'If tgVof.sShowHour = "Y" Then
                            If ilPgmLength - ilPgmIndex > 0 Then  '(start loc of c/r - previous start loc) = pos of current c/r. tet for c/r back to back, no data
                                tmSvr.sProgramInfo = Mid$(slEvtTitle, ilPgmIndex, ilPgmLength - ilPgmIndex) 'look for carriage return
                                'there was something in the programInfo to show, make sure only a time shows and not the break/position
                                If (InStr(tmSvr.sProgramInfo, "AM") > 0 Or InStr(tmSvr.sProgramInfo, "PM") > 0) Then

                                    '6-16-01 for C80- show program start time and name in one column and spots on the same line, not on separate line
                                    'i.e.   6:00A Show Name    Mon Advt     Tue Advt     Wed Advt, etc

                                    For ilDay = 0 To 6      'Bypass the program lengths
                                        ilTempComm = InStr(ilSpotIndex(ilDay), slEvt(ilDay), Chr$(10))
                                        If ilTempComm > 0 Then      'more to come, otherwise blank out field field
                                            ilSpotLength(ilDay) = ilTempComm
                                           ilSpotIndex(ilDay) = ilSpotLength(ilDay) + 1
                                        End If
                                    Next ilDay

                                    For ilDay = 0 To 6      'do monday thru sunday spot information
                                        slStr = ""
                                        ilTempComm = InStr(ilSpotIndex(ilDay), slEvt(ilDay), Chr$(10))
                                        If ilTempComm > 0 Then      'more to come, otherwise blank out field field
                                            ilSpotLength(ilDay) = ilTempComm
                                            If ilDay = 0 Then
                                               If ilSpotLength(ilDay) - ilSpotIndex(ilDay) > 0 Then  '(start loc of c/r - previous start loc) = pos of current c/r. tet for c/r back to back, no data
                                                  slStr = Mid$(slEvt(ilDay), ilSpotIndex(ilDay), ilSpotLength(ilDay) - ilSpotIndex(ilDay)) 'look for carriage return
                                               End If
                                               tmSvr.sProgramInfo = Trim$(tmSvr.sProgramInfo) & " " & slStr
                                           End If
                                           ilSpotIndex(ilDay) = ilSpotLength(ilDay) + 1
                                        Else
                                           tmSvr.sProduct(ilDay) = ""
                                        End If
                                    Next ilDay
                                    'bypass the program length array
                                    ilPgmIndex = ilPgmIndex + 1
                                    ilPgmLength = InStr(ilPgmIndex, slEvtTitle, Chr$(10))

                                Else
                                    tmSvr.sProgramInfo = ""
                                End If
                            End If
                        'End If

                    Else
                        If ilPgmLength - ilPgmIndex > 0 Then  '(start loc of c/r - previous start loc) = pos of current c/r. tet for c/r back to back, no data
                            tmSvr.sProgramInfo = Mid$(slEvtTitle, ilPgmIndex, ilPgmLength - ilPgmIndex) 'look for carriage return
                        End If
                    End If
                End If
                '*********
                For ilDay = 0 To 6      'do monday thru sunday spot information

                    slStr = ""
                    slSegmentStr = ""
                    ilTempComm = InStr(ilSpotIndex(ilDay), slEvt(ilDay), Chr$(10))
                    If ilTempComm > 0 Then      'more to come, otherwise blank out field field
                       'ilSpotLength(ilDay) = InStr(ilSpotIndex(ilDay), slEvt(ilDay), Chr$(10))
                       ilSpotLength(ilDay) = ilTempComm
                       If ilSpotLength(ilDay) - ilSpotIndex(ilDay) > 0 Then  '(start loc of c/r - previous start loc) = pos of current c/r. tet for c/r back to back, no data
                            slTemp = Mid$(slEvt(ilDay), ilSpotIndex(ilDay), ilSpotLength(ilDay) - ilSpotIndex(ilDay)) 'look for carriage return
                            ilRet = InStr(slTemp, "^")    'look for special char that separates the advt & prod from segment name (c80)
                            If ilRet = 0 Then   'no segment exists
                                slSegmentStr = ""
                                'slTemp contains the advt & prod desc
                                slStr = slTemp
                            Else
                                slStr = Trim$(Left$(slTemp, ilRet - 1))     'remove the segment name
                                slSegmentStr = Trim$(Mid$(slTemp, ilRet + 1)) 'get the segment name
                            End If

                       End If
                       tmSvr.sProduct(ilDay) = slStr
                       slSegment(ilDay) = slSegmentStr
                       ilSpotIndex(ilDay) = ilSpotLength(ilDay) + 1
                    Else
                       tmSvr.sProduct(ilDay) = slStr
                       slSegment(ilDay) = slSegmentStr
                    End If
                Next ilDay
                ilPgmIndex = ilPgmLength + 1
                tmSvr.iPosition = ilMajorSort
                tmSvr.iSeq = tmSvr.iSeq + 1
                tmSvr.sZone = slZone
                gPackDateLong llDate, ilStartofWk(0), ilStartofWk(1)    '7-9-03, start weeks not updated proper
                tmSvr.iStartofWk(0) = ilStartofWk(0)
                tmSvr.iStartofWk(1) = ilStartofWk(1)
                tmSvr.lHd1CefCode = tgVof.lHd1CefCode 'log header comment code
                tmSvr.lFt1CefCode = tgVof.lFt1CefCode 'log footer #1 comment code
                tmSvr.lFt2CefCode = tgVof.lFt2CefCode 'log footer #2 comment code
                tmSvr.iType = 0         'this is a flag for Crystal output to show boxes on C80
                ilRet = btrInsert(hlSvr, tmSvr, imSvrRecLen, INDEXKEY0)
                If StrComp(smLogName, "C80", 1) = 0 Then
                    'write out the segment descriptions associated with the spots
                    tmSvr.sProgramInfo = ""
                    tmSvr.iType = -1            'flag to show boxes on C80 detail
                    For ilDay = 0 To 6
                        tmSvr.sProduct(ilDay) = slSegment(ilDay)
                        tmSvr.iSeq = tmSvr.iSeq + 1     'keep the same times apart, and this must follow the spot
                    Next ilDay
                    ilRet = btrInsert(hlSvr, tmSvr, imSvrRecLen, INDEXKEY0)
                End If
            Next ilIndex

            For ilTempComm = 1 To 3 Step 1
                slNewComment = slTempComment(ilTempComm)
                If slNewComment <> "" Then
                ilFound = False
                For ilComm = LBound(slLogComment) To UBound(slLogComment) - 1 Step 1
                    If StrComp(slNewComment, slLogComment(ilComm)) = 0 Then
                        ilFound = True
                        Exit For
                    End If
                Next ilComm
                If Not ilFound Then
                    slLogComment(UBound(slLogComment)) = slNewComment
                    ReDim Preserve slLogComment(0 To UBound(slLogComment) + 1) As String
                End If


                For ilComm = 1 To 3 Step 1
                    If slComment(ilComm) = "" Then
                    slComment(ilComm) = slNewComment
                    Exit For
                    End If
                    If StrComp(slComment(ilComm), slNewComment, 1) = 0 Then
                    Exit For
                    End If
                Next ilComm
                End If
            Next ilTempComm
            llCurrTime = 86401
            Loop
            'Reset Index so same program can be found

            For ilIndex = 0 To 6 Step 1
            ilEvtIndex(ilIndex) = ilSvEvtIndex(ilIndex)
            Next ilIndex
        Next llDate
        End If
    Next ilZone
    Next ilVehicle
    'send any special comments associated with the program events
    'to crystal to print as footer
    slStr = ""
    For ilComm = LBound(slLogComment) To UBound(slLogComment) - 1
    If slStr = "" Then
        slStr = slLogComment(ilComm)
    Else
       'ideally we want each comment on a separate line, but c/r doesn't work
       'slStr = slStr & Chr(10) & slLogComment(ilComm)
        slStr = slStr & ", " & slLogComment(ilComm)
    End If

    Next ilComm
    If Len(slStr) <= 250 Then           '5-11-05 crystal limit is 255 characters
        If Not gSetFormula("Comments", "'" & slStr & "'") Then
            ilRet = ilRet
        End If
    Else
        slStr = Mid$(slStr, 1, 250)
        If Not gSetFormula("Comments", "'" & slStr & "...'") Then
            ilRet = ilRet
        End If

    End If
    Erase ilEvtIndex
    Erase ilSvEvtIndex
    Erase ilPrgFnd
    Erase slComment
    Erase slTempComment
    Erase slLogComment
    Erase slEvt
    Erase tlOdf0
    Erase tlOdf1
    Erase tlOdf2
    Erase tlOdf3
    Erase tlOdf4
    Erase tlOdf5
    Erase tlOdf6
    Erase tmOdf0
    Erase tmOdf1
    Erase tmOdf2
    Erase tmOdf3
    Erase tmOdf4
    Erase tmOdf5
    Erase tmOdf6
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmCef)
    ilRet = btrClose(hmEnf)
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hlSvr)
    btrDestroy hmAnf
    btrDestroy hmCef
    btrDestroy hmEnf
    btrDestroy hlODF
    btrDestroy hlSvr
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gLogSeven                       *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:12/14/97      By:d. Hosaka      *
'*                                                     *
'*       Comments: Generate Seven day log report       *
'*       12/14/97 New log generation scheme processes  *
'*       only 1 vehicle at a time automatically        *
'*       from the main log screen                      *
'           3/9/99 Add option to check for L08 or L27  *
'                  L08 is orig.comml schedule          *
'                  L27 is comml sch w/new event Ids    *
'
'*                                                     *
'*******************************************************
'Sub gLogSevenRptAll(ilPreview As Integer, slName As String, ilUrfCode As Integer, slLogName As String)
'    Dim slFileName As String * 80
'    Dim slPrinter As String
'    Dim slPort As String
'    Dim ilErrorFlag As Integer
'    Dim llRecNo As Long
'    Dim llRecsRemaining As Long
'    Dim slLangText As String
'    Dim ilDBEof As Integer
'    Dim ilDummy As Integer
'    Dim ilRet As Integer
'    Dim slStr As String
'    Dim ilZone As Integer
'    Dim sLZone As String
'    Dim ilCode As Integer
'    Dim slCode As String
'    Dim slNameCode As String
'    Dim llDate As Long
'    Dim llStartDate As Long
'    Dim llEndDate As Long
'    Dim slDate As String
'    Dim slTime As String
'    Dim slLen As String
'    Dim llStartTime As Long
'    Dim llEndTime As Long
'    Dim ilDay As Integer
'    Dim ilindex As Integer
'    Dim hlOdf As Integer            'One day log file handle
'    Dim tlOdfSrchKey As ODFKEY0            'ODF record image
'    Dim ilOdfRecLen As Integer        'ODF record length
'    Dim ilVefCode As Integer
'    Dim ilVehicle As Integer
'    Dim ilAnyEvt As Integer
'    Dim ilFind As Integer       '0=Prog; 1=Note, 2=Break
'    Dim ilPass As Integer
'    Dim ilPageNo As Integer
'    Dim ilBreakNo As Integer
'    Dim ilPositionNo As Integer
'    Dim llCurrTime As Long
'    Dim ilMode As Integer
'    Dim ilVehCount As Integer
'    Dim llHrCount As Long
'    Dim llWkCount As Long
'    Dim ilZCount As Integer
'    Dim llLastHour As Long
'    ReDim ilEvtIndex(0 To 6) As Integer
'    ReDim ilSvEvtIndex(0 To 6) As Integer
'    ReDim ilPrgFnd(0 To 6) As Integer
'    Dim slEvtTitle As String
'    ReDim slComment(1 To 3) As String
'    ReDim slTempComment(1 To 3) As String
'    Dim slNewComment As String
'    Dim ilComm As Integer
'    Dim ilTempComm As Integer
'    Dim ilCommEnabled As Integer
'    Dim ilFirstZoneIndex As Integer
'    Dim ilRetComm As Integer
'    ReDim slEvt(0 To 6) As String 'Events to be printed
'    ReDim tlOdf0(1 To 1) As ODFEXT
'    ReDim tlOdf1(1 To 1) As ODFEXT
'    ReDim tlOdf2(1 To 1) As ODFEXT
'    ReDim tlOdf3(1 To 1) As ODFEXT
'    ReDim tlOdf4(1 To 1) As ODFEXT
'    ReDim tlOdf5(1 To 1) As ODFEXT
'    ReDim tlOdf6(1 To 1) As ODFEXT
'    Dim ilNoteRet As Integer
'    Dim ilLastBreakNo As Integer
'    Dim ilAnyOutput As Integer
'    Dim slVehName As String
'    Dim ilUseZone As Integer
'    Dim ilVpfIndex As Integer
'    Dim ilLoop As Integer
'
'    smLogName = UCase$(Trim$(slLogName))
'    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmVef)
'        btrDestroy hmVef
'        Exit Sub
'    End If
'    imVefRecLen = Len(tmVef)
'    tmVefSrchKey.iCode = igCodes(0)         'new scheme only processes one vehicle at a time in this module
'    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If ilRet = BTRV_ERR_NONE Then
'        slVehName = tmVef.sName
'        ilRet = btrClose(hmVef)
'        btrDestroy hmVef
'    Else
'        ilRet = btrClose(hmVef)
'        btrDestroy hmVef
'        Exit Sub
'    End If
'    slDate = RptSelLg!edcSelCFrom.Text   'Start date
'    If (slDate = "") Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCFrom.SetFocus
'        Exit Sub
'    End If
'    If Not gValidDate(slDate) Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCFrom.SetFocus
'        Exit Sub
'    End If
'    llStartDate = gDateValue(slDate)
'    llEndDate = llStartDate + Val(RptSelLg!edcSelCFrom1.Text) - 1
'    slTime = RptSelLg!edcSelCTo.Text   'Start Time
'    If (slTime = "") Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCTo.SetFocus
'        Exit Sub
'    End If
'    If Not gValidTime(slTime) Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCTo.SetFocus
'        Exit Sub
'    End If
'    llStartTime = CLng(gTimeToCurrency(slTime, False))
'    slTime = RptSelLg!edcSelCTo1.Text   'End Time
'    If (slTime = "") Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCTo1.SetFocus
'        Exit Sub
'    End If
'    If Not gValidTime(slTime) Then
'        igGenRpt = False
'        RptSelLg!frcOutput.Enabled = igOutput
'        RptSelLg!frcCopies.Enabled = igCopies
'        'RptSelLg!frcWhen.Enabled = igWhen
'        RptSelLg!frcFile.Enabled = igFile
'        RptSelLg!frcOption.Enabled = igOption
'        'RptSelLg!frcRptType.Enabled = igReportType
'        Beep
'        RptSelLg!edcSelCTo1.SetFocus
'        Exit Sub
'    End If
'    llEndTime = CLng(gTimeToCurrency(slTime, True)) - 1
'    hlOdf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hlOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hlOdf)
'        btrDestroy hlOdf
'        Exit Sub
'    End If
'    hmEnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmEnf)
'        ilRet = btrClose(hlOdf)
'        btrDestroy hmEnf
'        btrDestroy hlOdf
'        Exit Sub
'    End If
'    imEnfRecLen = Len(tmEnf)
'    hmCef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCef)
'        ilRet = btrClose(hmEnf)
'        ilRet = btrClose(hlOdf)
'        btrDestroy hmCef
'        btrDestroy hmEnf
'        btrDestroy hlOdf
'        Exit Sub
'    End If
'    imCefRecLen = Len(tmCef)
'    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmAnf)
'        ilRet = btrClose(hmCef)
'        ilRet = btrClose(hmEnf)
'        ilRet = btrClose(hlOdf)
'        btrDestroy hmCef
'        btrDestroy hmAnf
'        btrDestroy hmEnf
'        btrDestroy hlOdf
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'    imAnfRecLen = Len(tmAnf)
'    tmEnf.iCode = 0
'    tmSEnf.iCode = 0
'    tmAnf.iCode = 0
'    slFileName = sgRptPath & slName & Chr$(0)
'    'define variables for the load check (yes, L&L now checks the definition file!)
'    If StrComp(smLogName, "L27", 1) = 0 Then
''Rm**   gLogSeven
'    Else
''Rm**   gLogSevenL27
'    End If
'    'initiate printing
'    ilAnyOutput = False
'    slLangText = "Printing..."
'    If ilPreview <> 0 Then
''VB6**  ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, RptSelLg.hWnd, slLangText)
''VB6**  ilDummy = LlPreviewSetTempPath(hdJob, sgRptSavePath)
''VB6**  ilDummy = LlPreviewSetResolution(hdJob, 200)
'    Else
''VB6**  ilRet = LlPrintWithBoxStart(hdJob, LL_PROJECT_LIST, slFileName, LL_PRINT_NORMAL Or LL_PRINT_MULTIPLE_JOBS, LL_BOXTYPE_BRIDGEMETER, RptSelLg.hWnd, slLangText)
'    End If
'    If (ilRet = 0) Then
'
'        'Compute Time of job by #Weeks*#Hours*#Zones*#Vehicles
'        llWkCount = (llEndDate - llStartDate) \ 7 + 1
'        llHrCount = (llEndTime - llStartTime) \ 3600 + 1
'        ilZCount = 0
'        ilFirstZoneIndex = -1
'        For ilZone = 0 To 3 Step 1
'            If RptSelLg!ckcSelC3(ilZone).Value = vbChecked Then
'                ilZCount = ilZCount + 1
'                If ilFirstZoneIndex = -1 Then
'                    ilFirstZoneIndex = ilZone
'                End If
'            End If
'        Next ilZone
'        ilVehCount = 1
'        'ilVehCount = 0
'        'For ilVehicle = 0 To RptSelLg!lbcSelection(0).ListCount - 1 Step 1
'        '    If RptSelLg!lbcSelection(0).Selected(ilVehicle) Then
'        '        ilVehCount = ilVehCount + 1
'        '    End If
'        'Next ilVehicle
'        llRecsRemaining = llWkCount * llHrCount * ilZCount * ilVehCount
'        llLastHour = llStartTime \ 3600 + 1
'        llRecNo = 0
'        ilErrorFlag = 0
'        ilDBEof = False
''VB6**  slPrinter = LlVBPrintGetPrinter(hdJob)
''VB6**  slPort = LlVBPrintGetPort(hdJob)
'        slComment(1) = ""
'        slComment(2) = ""
'        slComment(3) = ""
'        'For ilVehicle = 0 To RptSelLg!lbcSelection(0).ListCount - 1 Step 1
'        '    If RptSelLg!lbcSelection(0).Selected(ilVehicle) Then
'        '        slNameCode = tgVehicle(ilVehicle).sKey 'Traffic!lbcVehicle.List(ilVehicle)
'        '        ilRet = gParseItem(slNameCode, 1, "\", slName)
'        '        ilRet = gParseItem(slName, 3, "|", slName)
'        '        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'        '        ilVefCode = Val(slCode)
'        For ilVehicle = 0 To UBound(igCodes) - 1 Step 1
'                ilVefCode = igCodes(ilVehicle)
'
'                ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)
'                ilUseZone = False
'                If ilVpfIndex > 0 Then
'                    If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
'                        ilUseZone = True
'                    Else
'                        'Zones not used,
'                        RptSelLg!ckcSelC3(0).Value = vbChecked  'True            'fake out to do just one
'                        RptSelLg!ckcSelC3(1).Value = vbUnchecked    'False
'                        RptSelLg!ckcSelC3(2).Value = vbUnchecked    'False
'                        RptSelLg!ckcSelC3(3).Value = vbUnchecked    'False
'                    End If
'                End If
'                'Set titles for week
'                For ilZone = 0 To 3 Step 1
'                    If RptSelLg!ckcSelC3(ilZone).Value = vbUnchecked Then
''VB6**                  ilErrorFlag = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter & Chr$(10) & "Gathering Information", (100# * llRecNo / llRecsRemaining))
'                        If ilErrorFlag <> 0 Then
'                            Exit For
'                        End If
'                        DoEvents
'                        Screen.MousePointer = vbHourglass
'                        slStr = "Printed" & Chr$(10) & Format$(gNow(), "m/d/yy") & " " & Format(Now, "h:mm AM/PM") & Chr$(0)
''VB6**                  ilDummy = LLDefineVariableExt(hdJob, "Printed", slStr, LL_TEXT, "")
'                        'slStr = "ABC Radio Network" & " " & "Commercial Schedule" & Chr$(0)
'                        If StrComp(smLogName, "L27", 1) <> 0 Then
'                            slStr = "ABC Radio Network" & " " & "Commercial Schedule" & Chr$(0)
'                        Else
'                            slStr = "ABC Radio Networks" & " " & "Commercial Schedule" & Chr$(0)
'                        End If
''VB6**                  ilDummy = LLDefineVariableExt(hdJob, "Name", slStr, LL_TEXT, "")
'                        For llDate = llStartDate To llEndDate Step 7    'One week at a time
'                            ilErrorFlag = 0
'                            ilDBEof = False
'                            llCurrTime = 86401
'                            ilPageNo = 1
'                            llLastHour = llStartTime \ 3600 + 1
'                            slDate = Format$(llDate, "m/d/yy")
'                            slStr = Format$(llDate, "dddd") & ", " & slDate
'                            slDate = Format$(llDate + 6, "m/d/yy")
'                            slStr = slStr & " to " & Format$(llDate + 6, "dddd") & ", " & slDate & Chr$(0)
''VB6**                      ilDummy = LLDefineVariableExt(hdJob, "LogDate", slStr, LL_TEXT, "")
'                            slStr = Trim$(Str$(ilPageNo)) & Chr$(0)
''VB6**                      ilDummy = LLDefineVariableExt(hdJob, "*Page*", slStr, LL_TEXT, "")
'                            slStr = Trim$(slVehName)
'                            Select Case ilZone
'                                Case 0  'Eastern
'                                    If ilUseZone Then
'                                        slStr = slStr & ", Eastern Time Zone" & Chr$(0)
'                                        sLZone = "EST"
'                                    Else
'                                        sLZone = "   "
'                                    End If
'                                Case 1  'Central
'                                    slStr = slStr & ", Central Time Zone" & Chr$(0)
'                                    sLZone = "CST"
'                                Case 2  'Mountain
'                                    slStr = slStr & ", Mountain Time Zone" & Chr$(0)
'                                    sLZone = "MST"
'                                Case 3  'Pacific
'                                    slStr = slStr & ", Pacific Time Zone" & Chr$(0)
'                                    sLZone = "PST"
'                            End Select
''VB6**                      ilDummy = LLDefineVariableExt(hdJob, "Vehicle", slStr, LL_TEXT, "")
'                            If ilZone = ilFirstZoneIndex Then
'                                For ilDay = 0 To 6 Step 1
'                                    slDate = Format$(llDate + ilDay, "m/d/yy")
'                                    Select Case ilDay
'                                        Case 0
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf0()
'                                        Case 1
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf1()
'                                        Case 2
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf2()
'                                        Case 3
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf3()
'                                        Case 4
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf4()
'                                        Case 5
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf5()
'                                       Case 6
'                                            mObtainOdf hlOdf, ilUrfCode, ilVefCode, sLZone, slDate, llStartTime, llEndTime, tmOdf6()
'                                    End Select
'                                Next ilDay
'                            End If
'                            For ilDay = 0 To 6 Step 1
'                                Select Case ilDay
'                                    Case 0
'                                        mMoveODF sLZone, tmOdf0(), tlOdf0()
'                                        ilEvtIndex(0) = LBound(tlOdf0)
'                                    Case 1
'                                        mMoveODF sLZone, tmOdf1(), tlOdf1()
'                                        ilEvtIndex(1) = LBound(tlOdf1)
'                                    Case 2
'                                        mMoveODF sLZone, tmOdf2(), tlOdf2()
'                                        ilEvtIndex(2) = LBound(tlOdf2)
'                                    Case 3
'                                        mMoveODF sLZone, tmOdf3(), tlOdf3()
'                                        ilEvtIndex(3) = LBound(tlOdf3)
'                                    Case 4
'                                        mMoveODF sLZone, tmOdf4(), tlOdf4()
'                                        ilEvtIndex(4) = LBound(tlOdf4)
'                                    Case 5
'                                        mMoveODF sLZone, tmOdf5(), tlOdf5()
'                                        ilEvtIndex(5) = LBound(tlOdf5)
'                                    Case 6
'                                        mMoveODF sLZone, tmOdf6(), tlOdf6()
'                                        ilEvtIndex(6) = LBound(tlOdf6)
'                                End Select
'                            Next ilDay
'                            Screen.MousePointer = vbDefault
''VB6**                      ilErrorFlag = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter & Chr$(10) & "Making Report Pages", (100# * llRecNo / llRecsRemaining))
'                            DoEvents
'                            'outer loop - one loop per page
'                            Do While (Not ilDBEof) And ilErrorFlag = 0 'VB6** And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
'                                ilAnyOutput = True
'                                slStr = Trim$(Str$(ilPageNo)) & Chr$(0)
''VB6**                          ilDummy = LLDefineVariableExt(hdJob, "*Page*", slStr, LL_TEXT, "")
''VB6**                          ilNoteRet = LLPrintEnableObject(hdJob, ":Notes", True)
''VB6**                          ilDummy = LLDefineVariableExt(hdJob, "Note1", slComment(1), LL_TEXT, "")
''VB6**                          ilDummy = LLDefineVariableExt(hdJob, "Note2", slComment(2), LL_TEXT, "")
''VB6**                          ilDummy = LLDefineVariableExt(hdJob, "Note3", slComment(3), LL_TEXT, "")
'                                slNewComment = ""
'                                slComment(1) = ""
'                                slComment(2) = ""
'                                slComment(3) = ""
'                                slTempComment(1) = ""
'                                slTempComment(2) = ""
'                                slTempComment(3) = ""
'                                ilPageNo = ilPageNo + 1
'                                'All object must be enabled to be printed
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Title", True)
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Events", True)
''VB6**                          ilRet = LLPrint(hdJob)
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Events", False)
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Title", True)
''VB6**                          LlDefineFieldStart hdJob
'                                If StrComp(smLogName, "L27", 1) <> 0 Then
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "DayTitle", "Program" & Chr$(10) & " Start Time" & Chr$(10) & Chr$(10) & "Com'l B & P", LL_TEXT, "")
'                                Else
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "DayTitle", "Program" & Chr$(10) & " Start Time" & Chr$(10) & Chr$(10) & Chr$(10) & "Com'l B & P", LL_TEXT, "")
'                                End If
'                                For ilDay = 0 To 6 Step 1
'                                    slStr = Format$(llDate + ilDay, "dddd" & ", " & "m/d/yy")
'                                    'slStr = gAddDayToDate(slDate)
'                                    If StrComp(smLogName, "L27", 1) <> 0 Then
'                                        slStr = slStr & Chr$(10) & "Pgm Length, Name" & Chr$(10) & "Pgm Number" & Chr$(10) & "Spot Length, Name" & Chr$(0)
'                                    Else
'                                        slStr = slStr & Chr$(10) & "Pgm ID" & Chr$(10) & "Pgm Length, Name" & Chr$(10) & "Avail ID" & Chr$(10) & "Spot Length, Name" & Chr$(0)
'                                    End If
'                                    Select Case ilDay
'                                        Case 0
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day1Title", slStr, LL_TEXT, "")
'                                        Case 1
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day2Title", slStr, LL_TEXT, "")
'                                        Case 2
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day3Title", slStr, LL_TEXT, "")
'                                        Case 3
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day4Title", slStr, LL_TEXT, "")
'                                        Case 4
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day5Title", slStr, LL_TEXT, "")
'                                        Case 5
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day6Title", slStr, LL_TEXT, "")
'                                        Case 6
''VB6**                                      ilDummy = LLDefineFieldExt(hdJob, "Day7Title", slStr, LL_TEXT, "")
'                                    End Select
'                                Next ilDay
'                                If ilRet = 0 Then
''VB6**                              ilRet = LlPrintFields(hdJob)
'                                End If
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Title", False)
''VB6**                          ilRet = LLPrintEnableObject(hdJob, ":Events", True)
'                                'any group painting delayed?
'                                'if so, print it here (just under the header)
'                                'we could put the text in a global variable too, but this method is nicer!
'                                'inner loop - one loop per data set
'                                Do While (Not ilDBEof) And ilErrorFlag = 0 And ilRet = 0 'VB6** And LlPrintGetCurrentPage(hdJob) <= LlPrintGetOption(hdJob, LL_OPTION_LASTPAGE)
'                                    'Find The next earliest time
'                                    slTempComment(1) = ""
'                                    slTempComment(2) = ""
'                                    slTempComment(3) = ""
'                                    ilAnyEvt = False
'                                    For ilDay = 0 To 6 Step 1
'                                        ilPrgFnd(ilDay) = False
'                                        slEvt(ilDay) = ""
'                                        Select Case ilDay
'                                            Case 0
'                                                mFindEarliestLogPrg llCurrTime, tlOdf0(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 1
'                                                mFindEarliestLogPrg llCurrTime, tlOdf1(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 2
'                                                mFindEarliestLogPrg llCurrTime, tlOdf2(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 3
'                                                mFindEarliestLogPrg llCurrTime, tlOdf3(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 4
'                                                mFindEarliestLogPrg llCurrTime, tlOdf4(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 5
'                                                mFindEarliestLogPrg llCurrTime, tlOdf5(), ilEvtIndex(ilDay), ilAnyEvt
'                                            Case 6
'                                                mFindEarliestLogPrg llCurrTime, tlOdf6(), ilEvtIndex(ilDay), ilAnyEvt
'                                        End Select
'                                    Next ilDay
'                                    If Not ilAnyEvt Then
'                                        ilDBEof = True
'                                        Exit Do
'                                    End If
'                                    'Obtain number of break for matching times
'                                    For ilindex = 0 To 6 Step 1
'                                        ilSvEvtIndex(ilindex) = ilEvtIndex(ilindex)
'                                    Next ilindex
'                                    ilAnyEvt = False
'                                    ilFind = 0
'                                    ilPass = 0
'                                    ilBreakNo = 1
'                                    ilPositionNo = 1
'                                    ilLastBreakNo = 1
'                                    ilMode = 0  'Test
'                                    Do
'                                        Do
'                                            ilAnyEvt = False
'                                            For ilDay = 0 To 6 Step 1
'                                                Select Case ilDay
'                                                    Case 0
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf0(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 1
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf1(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 2
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf2(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 3
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf3(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 4
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf4(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 5
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf5(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                    Case 6
'                                                        mLogPass ilMode, ilPass, llCurrTime, ilBreakNo, ilPositionNo, tlOdf6(), ilEvtIndex(ilDay), ilPrgFnd(ilDay), slEvt(ilDay), ilAnyEvt, slNewComment
'                                                End Select
'                                                If ilMode = 1 Then
'                                                    If slNewComment <> "" Then
'                                                        For ilComm = 1 To 3 Step 1
'                                                            If slTempComment(ilComm) = "" Then
'                                                                slTempComment(ilComm) = slNewComment
'                                                                Exit For
'                                                            End If
'                                                            If StrComp(slTempComment(ilComm), slNewComment, 1) = 0 Then
'                                                                Exit For
'                                                            End If
'                                                        Next ilComm
'                                                    End If
'                                                End If
'                                            Next ilDay
'                                            If Not ilAnyEvt Then
'                                                Exit Do
'                                            End If
'                                            If ilMode = 1 Then
'                                                If ilPass = 0 Then
'                                                    Exit Do
'                                                ElseIf ilPass = 1 Then
'                                                    ilMode = 0  'Another note
'                                                ElseIf ilPass = 2 Then
'                                                    ilMode = 0  'Another spot with this break no
'                                                    ilPositionNo = ilPositionNo + 1
'                                                ElseIf ilPass = 3 Then
'                                                    ilMode = 0  'Another spot with this break no
'                                                End If
'                                            Else
'                                                ilMode = 1  'Add events
'                                                If ilPass = 0 Then
'                                                    slEvtTitle = gCurrencyToTime(CCur(llCurrTime)) & Chr$(10)
'                                                ElseIf (ilPass = 1) Or (ilPass = 3) Then
'                                                    slEvtTitle = slEvtTitle & Chr$(10)
'                                                ElseIf ilPass = 2 Then
'                                                    slStr = Trim$(Str$(ilBreakNo))
'                                                    If Len(slStr) = 1 Then
'                                                        slStr = "0" & slStr
'                                                    End If
'                                                    'slEvtTitle = slEvtTitle & Chr$(10) & slStr & Trim$(Str$(ilPositionNo))
'                                                    If StrComp(smLogName, "L27", 1) <> 0 Then
'                                                        slEvtTitle = slEvtTitle & Chr$(10) & slStr & Trim$(Str$(ilPositionNo))
'                                                    Else
'                                                        If ilPositionNo = 1 Then
'                                                            slEvtTitle = slEvtTitle & Chr$(10) & Chr$(10) & slStr & Trim$(Str$(ilPositionNo))
'                                                        Else
'                                                            slEvtTitle = slEvtTitle & Chr$(10) & slStr & Trim$(Str$(ilPositionNo))
'                                                        End If
'                                                    End If
'                                                End If
'                                            End If
'                                        Loop
'                                        ilMode = 0
'                                        If ilPass = 0 Then
'                                            ilPass = 1
'                                        ElseIf ilPass = 1 Then
'                                            ilPass = 2
'                                        ElseIf ilPass = 2 Then
'                                            If (ilPositionNo = 1) Then
'                                                If ilBreakNo >= ilLastBreakNo + 6 Then
'                                                    ilPass = 3
'                                                Else
'                                                    ilBreakNo = ilBreakNo + 1
'                                                    ilPositionNo = 1
'                                                End If
'                                            Else
'                                                ilLastBreakNo = ilBreakNo
'                                                ilBreakNo = ilBreakNo + 1
'                                                ilPositionNo = 1
'                                            End If
'                                        ElseIf ilPass = 3 Then
'                                            Exit Do
'                                        End If
'                                    Loop
'                                    'for each record: define the fields according to their type
'                                    'Call DefineFields
'                                    'LlDefineFieldStart hdJob
'                                    If right$(slEvtTitle, 1) <> Chr$(10) Then
'                                        slEvtTitle = slEvtTitle & Chr$(10)
'                                    End If
'                                    For ilDay = 0 To 6 Step 1
'                                        If right$(slEvt(ilDay), 1) <> Chr$(10) Then
'                                            slEvt(ilDay) = slEvt(ilDay) & Chr$(10)
'                                        End If
'                                    Next ilDay
'                                    slEvtTitle = slEvtTitle & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "EventTitle", slEvtTitle, LL_TEXT, "")
'                                    slEvt(0) = slEvt(0) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day1Event", slEvt(0), LL_TEXT, "")
'                                    slEvt(1) = slEvt(1) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day2Event", slEvt(1), LL_TEXT, "")
'                                    slEvt(2) = slEvt(2) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day3Event", slEvt(2), LL_TEXT, "")
'                                    slEvt(3) = slEvt(3) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day4Event", slEvt(3), LL_TEXT, "")
'                                    slEvt(4) = slEvt(4) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day5Event", slEvt(4), LL_TEXT, "")
'                                    slEvt(5) = slEvt(5) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day6Event", slEvt(5), LL_TEXT, "")
'                                    slEvt(6) = slEvt(6) & Chr$(0)
''VB6**                              ilDummy = LLDefineFieldExt(hdJob, "Day7Event", slEvt(6), LL_TEXT, "")
'                                    'notify the user (how far have we come?)
''VB6**                              ilRet = LlPrintSetBoxText(hdJob, "Printing To " & slPrinter, (100# * llRecNo / llRecsRemaining))
'                                    DoEvents
'                                    'tell L&L to print the table line
'                                    If ilRet = 0 Then
''VB6**                                  ilRet = LlPrintFields(hdJob)
'                                    End If
'                                    'next data set if no error or warning
'                                    If ilRet = 0 Then
'                                        'Place space between rows of programs boxes
'                                        'ilRet = LLPrintGroupLine(hdJob, 0, " ") 'Print space within table
'                                        'If ilRet = LL_WRN_REPEAT_DATA Then
'                                        '   ilRet = 0  'Ignore repeat data
'                                        'End If
'                                        'Call DBNext
'                                        For ilTempComm = 1 To 3 Step 1
'                                            slNewComment = slTempComment(ilTempComm)
'                                            If slNewComment <> "" Then
'                                                For ilComm = 1 To 3 Step 1
'                                                    If slComment(ilComm) = "" Then
'                                                        slComment(ilComm) = slNewComment
'                                                        Exit For
'                                                    End If
'                                                    If StrComp(slComment(ilComm), slNewComment, 1) = 0 Then
'                                                        Exit For
'                                                    End If
'                                                Next ilComm
'                                            End If
'                                        Next ilTempComm
'                                        llRecNo = llRecNo + (llCurrTime \ 3600) + 1 - llLastHour
'                                        llLastHour = llCurrTime \ 3600 + 1
'                                        llCurrTime = 86401 'Advance to next program
'                                        'If ilRecNo > ilRecsRemaining Then
'                                        '    ilDBEOF = True
'                                        'End If
'                                    End If
'                                Loop  ' inner loop
'
'                                'if error or warning: different reactions:
'                                If ilRet < 0 Then
'                                    If ilRet <> LL_WRN_REPEAT_DATA Then
'                                        ilErrorFlag = ilRet
'                                    Else
'                                        'ilRet = LLPrintEnableObject(hdJob, ":Title", True)
'                                        'ilRet = LLPrintEnableObject(hdJob, ":Events", False)
'                                        'Reset Index so same program can be found
'
'                                        For ilindex = 0 To 6 Step 1
'                                            ilEvtIndex(ilindex) = ilSvEvtIndex(ilindex)
'                                        Next ilindex
'                                    End If
'                                End If
'                            Loop    ' while not EOF
'                            If ilErrorFlag <> 0 Then
'                                Exit For
'                            End If
'                        Next llDate
'                    End If
'                    If ilErrorFlag <> 0 Then
'                        Exit For
'                    End If
'                Next ilZone
'            'End If  'Vehicle selection
'            If ilErrorFlag <> 0 Then
'                Exit For
'            End If
'        Next ilVehicle
''VB6**  ilNoteRet = LLPrintEnableObject(hdJob, ":Notes", True)
''VB6**  ilDummy = LLDefineVariableExt(hdJob, "Note1", slComment(1), LL_TEXT, "")
''VB6**  ilDummy = LLDefineVariableExt(hdJob, "Note2", slComment(2), LL_TEXT, "")
''VB6**  ilDummy = LLDefineVariableExt(hdJob, "Note3", slComment(3), LL_TEXT, "")
'
'        'end print
'        If Not ilAnyOutput Then
''VB6**      ilRet = LLPrintEnableObject(hdJob, ":Title", False)
''VB6**      ilRet = LLPrintEnableObject(hdJob, ":Events", False)
''VB6**      ilRet = LLPrint(hdJob)
'        Else
''VB6**      ilRet = LLPrintEnableObject(hdJob, ":Title", True)
''VB6**      ilRet = LLPrintEnableObject(hdJob, ":Events", True)
'        End If
''VB6**  ilRet = LlPrintEnd(hdJob, 0)
''in case of preview: show the preview
'        If ilPreview <> 0 Then
'            If ilErrorFlag = 0 Then
''VB6**          ilDummy = LlPreviewDisplay(hdJob, slFileName, sgRptSavePath, RptSelLg.hWnd)
'            Else
'                mErrMsg ilErrorFlag
'            End If
''VB6**      ilDummy = LlPreviewDeleteFiles(hdJob, slFileName, sgRptSavePath)
'        Else
'            If ilErrorFlag <> 0 Then
'                mErrMsg ilErrorFlag
'            End If
'        End If
'    Else  ' LlPrintWithBoxStart
'        ilErrorFlag = ilRet
''VB6**  ilRet = LlPrintEnd(hdJob, 0)
'        mErrMsg ilErrorFlag
'    End If  ' LlPrintWithBoxStart
'
'    Erase ilEvtIndex
'    Erase ilSvEvtIndex
'    Erase ilPrgFnd
'    Erase slComment
'    Erase slTempComment
'    Erase slEvt
'    Erase tlOdf0
'    Erase tlOdf1
'    Erase tlOdf2
'    Erase tlOdf3
'    Erase tlOdf4
'    Erase tlOdf5
'    Erase tlOdf6
'    Erase tmOdf0
'    Erase tmOdf1
'    Erase tmOdf2
'    Erase tmOdf3
'    Erase tmOdf4
'    Erase tmOdf5
'    Erase tmOdf6
'    ilRet = btrClose(hmAnf)
'    ilRet = btrClose(hmCef)
'    ilRet = btrClose(hmEnf)
'    ilRet = btrClose(hlOdf)
'    btrDestroy hmAnf
'    btrDestroy hmCef
'    btrDestroy hmEnf
'    btrDestroy hlOdf
'End Sub
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
'*      Procedure Name:mFindEarliestLogPrg             *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the earliest time for an   *
'*                     log event from the specified    *
'*                     current time                    *
'*                                                     *
'*******************************************************
Sub mFindEarliestLogPrg(llCurrTime As Long, tlOdf() As ODFEXT, ilEvtIndex As Integer, ilAnyEvt As Integer)
    Dim slTime As String
    Dim llTime As Long
    Do While ilEvtIndex < UBound(tlOdf)
        If tlOdf(ilEvtIndex).iEtfCode = 1 Then
            gUnpackTime tlOdf(ilEvtIndex).iLocalTime(0), tlOdf(ilEvtIndex).iLocalTime(1), "A", "1", slTime
            llTime = CLng(gTimeToCurrency(slTime, False))
            If llTime <= llCurrTime Then
                llCurrTime = llTime
                ilAnyEvt = True
            End If
            Exit Sub
        End If
        ilEvtIndex = ilEvtIndex + 1
    Loop
    Exit Sub
End Sub
'
'
'           mGetmnfSegment - obtain the MNF record type "K" (Segments)
'               whose pointer is stored in the contract header
'
'           Return the string description (or blank if none)
'
'           DH: Created 6-19-01
'
Function mGetMnfSegment(ilMnfSegmentCode As Integer) As String
Dim ilLoop As Integer
Dim ilRet As Integer
    If ilMnfSegmentCode > 0 Then
        ilRet = gObtainMnfForType("K", sgMNFCodeTagLg, tmMnfSegs())
        If ilRet = False Then
            mGetMnfSegment = ""
        Else
            'loop to find the matching one from header
            For ilLoop = LBound(tmMnfSegs) To UBound(tmMnfSegs) - 1
                If ilMnfSegmentCode = tmMnfSegs(ilLoop).iCode Then
                    mGetMnfSegment = tmMnfSegs(ilLoop).sName
                    Exit For
                End If
            Next ilLoop
        End If
    Else
        mGetMnfSegment = ""
    End If
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
    '10-9-01
    Dim ilSSFType As Integer
    'Dim slSsfType As String
    Dim ilRet As Integer
    '10-9-01
    '11/24/12
    'ilSSFType = 0
    ilSSFType = tlSdf.iGameNo
    ''slSsfType = "O" 'On Air
    ilSpotSeqNo = 0
    If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        '10-9-01 If (tmSsf.sType <> slSsfType) Or (tmSsf.iVefCode <> tlSdf.iVefCode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
        If (tmSsf.iType <> ilSSFType) Or (tmSsf.iVefCode <> tlSdf.iVefCode) Or (tmSsf.iDate(0) <> tlSdf.iDate(0)) Or (tmSsf.iDate(1) <> tlSdf.iDate(1)) Or (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then

            '10-9-01 tmSsfSrchKey.sType = slSsfType
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
        '10-9-01 If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
            ilEvtIndex = 1
            Do
                If ilEvtIndex > tmSsf.iCount Then
                    imSsfRecLen = Len(tmSsf)
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    '10-9-01 If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slSsfType) And (tmSsf.iVefCode = tlSdf.iVefCode) And (tmSsf.iDate(0) = tlSdf.iDate(0)) And (tmSsf.iDate(1) = tlSdf.iDate(1)) Then
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
'*
'*      Procedure Name:mLogPass
'*
'*             Created:4/21/94       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:Obtain event
'*
'*      5-24-01 Create another version of coml summary(c80)
'              that is similar to L39
'       5-11-05 if no event name, blank out so garbage char are not used
'*******************************************************
Sub mLogPass(ilMode As Integer, ilPass As Integer, llCurrTime As Long, ilBreakNo As Integer, ilPositionNo As Integer, tlOdf() As ODFEXT, ilEvtIndex As Integer, ilPrgFnd As Integer, slEvt As String, ilAnyEvt As Integer, slComment As String)
    Dim slTime As String
    Dim slLen As String
    Dim llTime As Long
    Dim ilRet As Integer
    Dim slEvtID As String
    Dim slSegName As String
    Dim slStr As String
    slComment = ""
    slEvtID = ""
    If ilEvtIndex < UBound(tlOdf) Then
        If ilPass = 0 Then
            gUnpackTime tlOdf(ilEvtIndex).iLocalTime(0), tlOdf(ilEvtIndex).iLocalTime(1), "A", "1", slTime
            llTime = CLng(gTimeToCurrency(slTime, False))
            If llTime = llCurrTime Then
                ilAnyEvt = True
                ilPrgFnd = True
                If ilMode = 0 Then
                    Exit Sub
                End If
                slEvt = ""

                If StrComp(smLogName, "L27", 1) = 0 Or StrComp(smLogName, "L37", 1) = 0 Or StrComp(smLogName, "L39", 1) = 0 Then
                    mReadCefRec tlOdf(ilEvtIndex).lEvtIDCefCode
                    'If tmCef.iStrLen > 0 Then
                    '    slEvtID = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                    'End If
                    slEvtID = gStripChr0(tmCef.sComment)
                    slEvt = slEvtID & Trim$(tlOdf(ilEvtIndex).sDupeAvailID) & Chr$(10)    'Add blank line after program and before other events
                End If
                'Convert length
                gUnpackLength tlOdf(ilEvtIndex).iLen(0), tlOdf(ilEvtIndex).iLen(1), "1", False, slLen
                If StrComp(smLogName, "C80", 1) = 0 Then
                    slEvt = slEvt & slLen & Chr$(10)
                Else
                    slEvt = slEvt & slLen
                End If
                If tlOdf(ilEvtIndex).iEnfCode > 0 Then
                    'Get Event name
                    If tmEnf.iCode <> tlOdf(ilEvtIndex).iEnfCode Then
                        tmEnfSrchKey.iCode = tlOdf(ilEvtIndex).iEnfCode
                        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) Then
                            If StrComp(smLogName, "C80", 1) = 0 Then
                                slEvt = slEvt & Trim$(tmEnf.sName)        'no space required, this field is at beginning of column
                            Else
                                slEvt = slEvt & " " & Trim$(tmEnf.sName) 'need space between length and event name
                            End If
                        End If
                    Else
                        If StrComp(smLogName, "C80", 1) = 0 Then
                            slEvt = slEvt & Trim$(tmEnf.sName)         'no space required, this field is at beginning of column
                        Else
                            slEvt = slEvt & " " & Trim$(tmEnf.sName)   'need space between length & event name
                        End If
                    End If
                Else                    '5-11-05 no event name
                    tmEnf.sName = ""
                End If
                'slEvt = slEvt & Chr$(10) & Trim$(tlOdf(ilEvtIndex).sProgCode)    'Add blank line after program and before other events
                If StrComp(smLogName, "L27", 1) <> 0 And StrComp(smLogName, "L37", 1) <> 0 And StrComp(smLogName, "L39", 1) <> 0 And StrComp(smLogName, "C80", 1) <> 0 Then
                    slEvt = slEvt & Chr$(10) & Trim$(tlOdf(ilEvtIndex).sProgCode)    'Add blank line after program and before other events
                'Else
                '    mReadCefRec tlOdf(ilEvtIndex).lCefCode
                '    If tmCef.iStrLen > 0 Then
                '        slEvtID = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                '    End If
                '    slEvt = slEvt & Chr$(10) & slEvtID & Trim$(tlOdf(ilEvtIndex).sDupeAvailID)    'Add blank line after program and before other events
                End If
                ilEvtIndex = ilEvtIndex + 1
            Else
                If ilMode = 0 Then
                    Exit Sub
                End If
                'slEvt = slEvt & Chr$(10)    'Blank line
            End If
        End If
        If (ilPass = 1) Or (ilPass = 3) Then  'Note
            If (tlOdf(ilEvtIndex).iEtfCode > 13) And (ilPrgFnd) Then
                ilAnyEvt = True
                If ilMode = 0 Then
                    Exit Sub
                End If
                'Get Event name
                slEvt = slEvt & Chr$(10)
                'Next line
                If tlOdf(ilEvtIndex).iEnfCode > 0 Then
                    If tmEnf.iCode <> tlOdf(ilEvtIndex).iEnfCode Then
                        tmEnfSrchKey.iCode = tlOdf(ilEvtIndex).iEnfCode
                        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If (ilRet = BTRV_ERR_NONE) Then
                            slEvt = slEvt & " " & Trim$(tmEnf.sName)
                        End If
                    Else
                        slEvt = slEvt & " " & Trim$(tmEnf.sName)
                    End If
                End If
                'Get comments
                mReadCefRec tlOdf(ilEvtIndex).lCefCode
                'If tmCef.iStrLen > 0 Then
                '    slComment = Trim$(tmEnf.sName) & ": " & Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                'End If
                slStr = gStripChr0(tmCef.sComment)
                If slStr <> "" Then
                    slComment = Trim$(tmEnf.sName) & ": " & slStr
                End If
                ilEvtIndex = ilEvtIndex + 1
            Else
                If ilMode = 0 Then
                    Exit Sub
                End If
                slEvt = slEvt & Chr$(10)    'Blank line
            End If
        End If
        If ilPass = 2 Then  'Spot
            If ((tlOdf(ilEvtIndex).iEtfCode = 0) And (ilBreakNo = tlOdf(ilEvtIndex).iBreakNo) And (ilPositionNo = tlOdf(ilEvtIndex).iPositionNo)) And (ilPrgFnd) Then
                ilAnyEvt = True
                If ilMode = 0 Then
                    Exit Sub
                End If
                slEvt = slEvt & Chr$(10)    'Next line
                'if first spot of break, show the avail event ID if it exists
                If (ilPositionNo = 1) And ((StrComp(smLogName, "L27", 1) = 0) Or (StrComp(smLogName, "L37", 1) = 0) Or (StrComp(smLogName, "L39", 1) = 0)) Then
                    mReadCefRec tlOdf(ilEvtIndex).lEvtIDCefCode
                    'If tmCef.iStrLen > 0 Then
                    '    slEvtID = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                    'End If
                    slEvtID = gStripChr0(tmCef.sComment)
                    slEvt = slEvt & slEvtID & Trim$(tlOdf(ilEvtIndex).sDupeAvailID) & Chr$(10)
                End If
                If tmAnf.iCode <> tlOdf(ilEvtIndex).ianfCode Then
                    If tlOdf(ilEvtIndex).ianfCode <> 0 Then
                        If tmAnf.iCode <> tlOdf(ilEvtIndex).ianfCode Then
                            tmAnfSrchKey.iCode = tlOdf(ilEvtIndex).ianfCode
                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                        If (ilRet = BTRV_ERR_NONE) Then
                            If StrComp(smLogName, "C80", 1) = 0 Then    'C80 doesnt show the first letter of avail name
                                slEvt = slEvt & ""
                            Else
                                If Left$(Trim$(tmAnf.sName), 1) = "N" Then
                                    slEvt = slEvt & "N"
                                Else
                                    slEvt = slEvt & "L"
                                End If
                            End If
                        End If
                    End If
                Else
                    If StrComp(smLogName, "C80", 1) = 0 Then
                        slEvt = slEvt & ""
                    Else
                        If Left$(Trim$(tmAnf.sName), 1) = "N" Then
                            slEvt = slEvt & "N"
                        Else
                            slEvt = slEvt & "L"
                        End If
                    End If
                End If
                gUnpackLength tlOdf(ilEvtIndex).iLen(0), tlOdf(ilEvtIndex).iLen(1), "1", True, slLen
                slEvt = slEvt & slLen
                slEvt = slEvt & " " & UCase$(Trim$(tlOdf(ilEvtIndex).sShortTitle)) 'UCase$(Trim$(tlOdf(ilEvtIndex).sProduct))
                If StrComp(smLogName, "C80", 1) = 0 And tgSpf.sCUseSegments = "Y" Then
                    slSegName = mGetMnfSegment(tlOdf(ilEvtIndex).imnfSeg)
                    slEvt = slEvt & "^" & slSegName     'insert delimeter to find the segment name
                End If
                ilEvtIndex = ilEvtIndex + 1
            Else
                If ilMode = 0 Then
                    Exit Sub
                End If
                slEvt = slEvt & Chr$(10)    'Blank line
            End If
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveODF                        *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move from all zone ODF to Zone  *
'*                     ODF                             *
'*                                                     *
'*      5-24-01 Too many records were processed from tlAllODF
'               (adjust for loop with minus one)
'*******************************************************
Sub mMoveODF(slZone As String, tlAllOdf() As ODFEXT, tlZoneOdf() As ODFEXT)
    Dim ilUpper As Integer
    Dim ilIndex As Integer
    ReDim tlZoneOdf(LBound(tlZoneOdf) To LBound(tlZoneOdf)) As ODFEXT
    ilUpper = LBound(tlZoneOdf)
    For ilIndex = LBound(tlAllOdf) To UBound(tlAllOdf) - 1 Step 1 '5-24-01
        If tlAllOdf(ilIndex).sZone = slZone Then
            tlZoneOdf(ilUpper) = tlAllOdf(ilIndex)
            ilUpper = ilUpper + 1
            ReDim Preserve tlZoneOdf(LBound(tlZoneOdf) To ilUpper) As ODFEXT
        End If
    Next ilIndex
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainOdf                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain the Odf records for date*
'*                                                     *
'*******************************************************
Sub mObtainOdf(hlODF As Integer, ilUrfCode As Integer, ilVefCode As Integer, slZone As String, slDate As String, llStartTime As Long, llEndTime As Long, tlOdfExt() As ODFEXT)
'
'   mObtainOdf
'   Where:
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slTime As String
    Dim ilUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlodfExtSrchKey As ODFKEY0
    Dim tlLongTypeBuff As POPLCODE      '10-10-01
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer
    Dim tlOdf As ODF
    'ReDim tlOdfExt(1 To 1) As ODFEXT
    'ilRecLen = Len(tlOdfExt(1)) 'btrRecordLength(hlAdf)  'Get and save record length
    ReDim tlOdfExt(0 To 0) As ODFEXT
    ilRecLen = Len(tlOdfExt(0)) 'btrRecordLength(hlAdf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlODF   'Clear any previous extend operation
    'tlOdfExtSrchKey.iUrfCode = ilUrfCode
    tlodfExtSrchKey.iVefCode = ilVefCode
    gPackDate slDate, tlodfExtSrchKey.iAirDate(0), tlodfExtSrchKey.iAirDate(1)
    gPackDate slDate, ilAirDate0, ilAirDate1
    slTime = gCurrencyToTime(CCur(llStartTime))
    gPackTime slTime, tlodfExtSrchKey.iLocalTime(0), tlodfExtSrchKey.iLocalTime(1)
    tlodfExtSrchKey.sZone = "" 'slZone
    tlodfExtSrchKey.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlodfExtSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlODF, llNoRec, -1, "UC", "ODFEXTPK", ODFEXTPK) 'Set extract limits (all records)
    'tlIntTypeBuff.iType = ilUrfCode
    'ilOffset = gFieldOffsetExtra("ODF", "OdfUrfCode")
    'ilRet = btrExtAddLogicConst(hlOdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
    tlIntTypeBuff.iType = ilVefCode
    ilOffSet = gFieldOffsetExtra("ODF", "OdfVefCode")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
    tlDateTypeBuff.iDate0 = igODFGenDate(0)     '5-25-01
    tlDateTypeBuff.iDate1 = igODFGenDate(1)
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    '10-10-01
    tlLongTypeBuff.lCode = lgGenTime
    'tlDateTypeBuff.iDate0 = igODFGenTime(0)     '5-25-01
    'tlDateTypeBuff.iDate1 = igODFGenTime(1)
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenTime")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    slTime = gCurrencyToTime(CCur(llEndTime))
    gPackTime slTime, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffsetExtra("ODF", "OdfLocalTime")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_TIME, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    'tlCharTypeBuff.sType = Left$(slZone, 1)
    'ilOffset = gFieldOffsetExtra("ODF", "OdfZone")
    'ilRet = btrExtAddLogicConst(hlOdf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = gFieldOffsetExtra("ODF", "OdfLocalTime")
'    ilOffset = gFieldOffsetExtra("ODF", "OdfAirTime")
    ilRet = btrExtAddField(hlODF, ilOffSet, 4)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfZone")
    ilRet = btrExtAddField(hlODF, ilOffSet, 3)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfEtfCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfEnfCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfProgCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 5)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfAnfCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfLength")
    ilRet = btrExtAddField(hlODF, ilOffSet, 4)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfProduct")
    ilRet = btrExtAddField(hlODF, ilOffSet, 35)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfMnfSubFeed")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfBreakNo")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfPositionNo")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfCefCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 4)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfShortTitle")
    ilRet = btrExtAddField(hlODF, ilOffSet, 15)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfAdfCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'extract advt code
    ilOffSet = gFieldOffsetExtra("ODF", "OdfEvtIDCefCode")
    ilRet = btrExtAddField(hlODF, ilOffSet, 4)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfDupeAvailID")
    ilRet = btrExtAddField(hlODF, ilOffSet, 5)  'Extract iCode field
    '6-19-01
    ilOffSet = gFieldOffsetExtra("ODF", "OdfmnfSeg")
    ilRet = btrExtAddField(hlODF, ilOffSet, 2)  'Extract iCode field
    ilOffSet = gFieldOffsetExtra("ODF", "OdfCifCode")        '8-1-16 copy pointer
    ilRet = btrExtAddField(hlODF, ilOffSet, 4)
   'ilRet = btrExtGetNextExt(hlOdf)    'Extract record
    ilUpper = UBound(tlOdfExt)
    ilRet = btrExtGetNext(hlODF, tlOdfExt(ilUpper), ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
    End If
    'ilRet = btrExtGetFirst(hlIcf, tlIcf, ilRecLen, llRecPos)
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hlODF, tlOdfExt(ilUpper), ilRecLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        If tlOdfExt(ilUpper).iMnfSubFeed = 0 Then   'bypass records with subfeed
           '4-8-05 this filter was put in wrong place, as it did not handle all zones properly
           '2-3-05 filter out non-matching zones
            'If Trim$(slZone) = "" Or Left$(slZone, 1) = Left$(tlOdfExt(ilUpper).sZone, 1) Then
                'ReDim Preserve tlOdfExt(1 To ilUpper + 1) As ODFEXT
                ReDim Preserve tlOdfExt(0 To ilUpper + 1) As ODFEXT
                ilUpper = ilUpper + 1
            'End If
        End If
        ilRet = btrExtGetNext(hlODF, tlOdfExt(ilUpper), ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlODF, tlOdfExt(ilUpper), ilRecLen, llRecPos)
        Loop
    Loop
    Exit Sub
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

Private Function mGetPrfFromCif(llCifCode As Long) As Long
Dim ilRet As Integer

            mGetPrfFromCif = 0
            tmCifSrchKey.lCode = llCifCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                mGetPrfFromCif = tmCif.lcpfCode     'get the product record for the ISCI pointer
            End If
            
            Exit Function

End Function
