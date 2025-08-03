Attribute VB_Name = "RPTCRAP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrap.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

'*******************************************************************
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' Created from RptCrCt.bas by W.Bjerke 10/30/97
'
' Description:
' This file contains the code for gathering the prepass data for the
' Actual/Projection Comparison Report
'*******************************************************************
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public lgPrintedCnts() As Long             'table to maintain the contr pointers
                                                'when contracts are finished printing, update print flag
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmPjf As PJF
Dim hmPjf As Integer
Dim imPjfRecLen As Integer
Dim tmPrf As PRF
Dim hmPrf As Integer
Dim hmSbf As Integer
Dim tmMnfNtr() As MNF
Dim tmSrchKey As LONGKEY0
Dim imPrfRecLen As Integer
Const NOT_SELECTED = 0
'  Receivables File
Dim tmRvf As RVF            'RVF record image
'*********************************************************************
'       gGetActProj()
'       Created 10/30/97 by W.Bjerke
'       Create Actual/Projection report prepass file
'
'       4/10/98 Correct the user inclusions/exclusions of contract types
'               and statuses.  Previously, parameters not tested, even
'               tho the questions were asked on screen
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)
'*********************************************************************
Sub gGetActProj(llCurrStart As Long, llCurrEnd As Long, llPrevStart As Long, llPrevEnd As Long)
Dim slMnfStamp As String
Dim slAirOrder As String * 1                'from site pref - bill as air or ordered
'ReDim ilLikePct(1 To 9) As Integer          'most likely percentage from potential code A, B & C
ReDim ilLikePct(0 To 9) As Integer          'most likely percentage from potential code A, B & C. Index zero ignored
'ReDim ilLikeCode(1 To 3) As Integer         'mnf most likely auto increment code for A, B, C
ReDim ilLikeCode(0 To 3) As Integer         'mnf most likely auto increment code for A, B, C. Index zero ignored
Dim ilLoop As Integer
Dim ilTemp As Integer
Dim ilSlsLoop As Integer
Dim ilChfLoop As Integer
Dim llRvfLoop As Long                       '2-11-05
Dim ilRet As Integer
ReDim ilRODate(0 To 1) As Integer           'Effective Date to match retrieval of Projection record
ReDim ilEnterDate(0 To 1) As Integer        'Btrieve format for date entered by user
Dim llClosestDate As Long                   'closest date to the rollover user entered date
Dim slDate As String
Dim slStr As String
Dim ilMonth As Integer
Dim ilYear As Integer
Dim ilTYYear As Integer
Dim llEnterFrom As Long                       'gather cnts whose entered date falls within llEnterFrom and llEnterTo
Dim llEnterTo As Long
'ReDim llProject(1 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array
ReDim llProject(0 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array. Index zero ignored
'ReDim llLYDates(1 To 2) As Long               'range of  qtr dates for contract retrieval (this year)
ReDim llLYDates(0 To 2) As Long               'range of  qtr dates for contract retrieval (this year). Index zero ignored
Dim llGross As Long
Dim slTYStartQtr As String
Dim slTYEndQtr As String
Dim ilAdfCode As Integer
Dim slNameCode As String
Dim slCode As String
ReDim tlSlsList(0 To 0) As SLSLIST
'ReDim llTYDates(1 To 2) As Long               'range of qtr dates for contract retrieval (last year)
ReDim llTYDates(0 To 2) As Long               'range of qtr dates for contract retrieval (last year). Index zero ignored
Dim slLYStartQtr As String
Dim slLYEndQtr As String
Dim tlChfAdvtExt() As CHFADVTEXT
'ReDim llStartDates(1 To 2) As Long            'temp array for last year vs this years range of dates
ReDim llStartDates(0 To 2) As Long            'temp array for last year vs this years range of dates. Index zero ignored
Dim ilFoundOne As Integer
Dim llLYGetFrom As Long                       'start date of last year
Dim llLYGetTo As Long                         'obtain last years qtr if cnt entered date equal/prior to this date (same time last year)
Dim slLYStartYr As String                     'start date of last year
Dim slLYEndYr As String
Dim llTYGetFrom As Long                       'start date of last year
Dim llTYGetTo As Long                         'obtain this years qtr if cnt entered date equal/prior to this date
Dim slTYStartYr As String                     'start date of this year
Dim slTYEndYr As String                       'end date of this year
Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer                        'include unsch holds/orders, sch holds/orders
Dim ilFound As Integer
Dim ilAdvtFound As Integer
Dim ilStartWk As Integer                        'starting week index to gather budget data
Dim ilEndWk As Integer                          'ending week index to gather budgets
Dim ilFirstWk As Integer                        'true if week 0 needs to be added when start wk = 1
Dim ilLastWk As Integer                         'true if week 53 needs to be added when end wk = 52
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilPastFut As Integer                        'loop to process past contracts, then current contracts
Dim ilClf As Integer                            'index to line from tgClfAP
Dim llAdjust As Long                            'Adjusted gross using the potential codes most likely %
Dim ilCorpStd As Integer            '1 = corp, 2 = std
Dim ilBvfCalType As Integer            '0=std, 1 = reg, 2 & 3 = julian, 4 = corp for jan thru dec, 5 = corp for fiscal year
Dim ilPjfCalType As Integer            '0=std, 1 = reg, 2 & 3 = julian, 4 = corp for jan thru dec, 5 = corp for fiscal year
Dim ilQuarter As Integer
Dim ilAdjust As Integer
ReDim ilAdjFlag(0 To 2) As Integer               'array to hold user selected adjustment flags
Dim ilAdjLoop As Integer
Dim slMnfRPU As String
Dim slMnfSSComm As String
Dim tlCntTypes As CNTTYPES
Dim tlTranType As TRANTYPES
Dim tlRvf() As RVF
Dim ilTY As Integer
Dim llDate As Long
Dim blIncludeNTR As Boolean
Dim blIncludeHardCost As Boolean
Dim blNTRWithTotal As Boolean
Dim tlNTRInfo() As NTRPacing
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim ilNTRCounter As Integer
Dim llSingleContract As Long            'Dan M added contract selectivity 7-14-08
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean
Dim blFailedBecauseInstallment As Boolean
Dim slTemp As String

    If Val(RptSelAp!edcSelC10.Text) > 0 And RptSelAp!edcSelC10.Text <> " " Then
         llSingleContract = Val(RptSelAp!edcSelC10.Text)
    Else
        llSingleContract = NOT_SELECTED
    End If
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPjf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imPjfRecLen = Len(tmPjf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPrf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)
    ' Dan M 6-23-08
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        btrDestroy hmPrf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    ReDim tmMnfNtr(0 To 0) As MNF
    imMnfRecLen = Len(tmMnfNtr(0))
    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02

    If RptSelAp!ckcSelC9(0).Value Or RptSelAp!ckcSelC9(1).Value Then    'don't waste time filling array if don't need.
         tlTranType.iNTR = True
        'did user choose ntr, hard cost?
        If RptSelAp!ckcSelC9(0).Value = 1 Then
             blIncludeNTR = True
        End If
        If RptSelAp!ckcSelC9(1).Value = 1 Then
            blIncludeHardCost = True
        End If

        ilRet = gObtainMnfForType("I", "", tmMnfNtr())
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered

    'get all the dates needed to work with
'    slDate = RptSelAp!edcSelCFrom.Text                          'effective date entred
    slDate = RptSelAp!CSI_CalFrom.Text                          'effective date entred 9-10-19 use csi calendar control vs edit box

    'obtain the entered dates year based on the std month
    llTYGetTo = gDateValue(slDate)                              'gather contracts thru this date
    gPackDateLong llTYGetTo, ilEnterDate(0), ilEnterDate(1)     'get btrieve date format for entered to pass to record to show on hdr
'    gGetRollOverDate RptSelAp, 2, slDate, llClosestDate         'send the lbcselection index to search, plus rollover date
    gGetRollOverDate RptSelAp, 1, slDate, llClosestDate         'send the lbcselection index to search, plus rollover date
    gPackDateLong llClosestDate, ilRODate(0), ilRODate(1)

    ilYear = Val(RptSelAp!edcSelCTo.Text)           'year requested
    ilQuarter = Val(RptSelAp!edcSelCTo1.Text)       'quarter requested
    igMonthOrQtr = ilQuarter

    If RptSelAp!rbcSelC7(0).Value Then
        ilCorpStd = 1                               'corp flag for genl subtrn
        ilPjfCalType = 4               'get week inx based on std for projections
        ilBvfCalType = 5               'get week inx based on fiscal dates
    Else
        ilCorpStd = 2
        ilPjfCalType = 0             'std month
        ilBvfCalType = 0             'both projections and budgets will be std
    End If

    'This Years start/end quarter and year dates
    gGetStartEndQtr ilCorpStd, ilYear, igMonthOrQtr, slTYStartQtr, slTYEndQtr
    llTYDates(1) = gDateValue(slTYStartQtr)
    llTYDates(2) = gDateValue(slTYEndQtr)
    gGetStartEndYear ilCorpStd, ilYear, slTYStartYr, slTYEndYr
    llTYGetFrom = gDateValue(slTYStartYr)
    'Last years start/end quarter and year dates
    gGetStartEndQtr ilCorpStd, ilYear - 1, igMonthOrQtr, slLYStartQtr, slLYEndQtr
    llLYDates(1) = gDateValue(slLYStartQtr)
    llLYDates(2) = gDateValue(slLYEndQtr)
    gGetStartEndYear ilCorpStd, ilYear - 1, slLYStartYr, slLYEndYr
    llLYGetFrom = gDateValue(slLYStartYr)

    'determine same time last year
    llLYGetTo = llLYGetFrom + (llTYGetTo - llTYGetFrom)
    tlCntTypes.iHold = gSetCheck(RptSelAp!ckcSelC3(0))     'Inc/excl holds
    tlCntTypes.iOrder = gSetCheck(RptSelAp!ckcSelC3(1))    'orders
    tlCntTypes.iTrade = gSetCheck(RptSelAp!ckcSelC5(1))    'trades (100% exclusion only)
    tlCntTypes.iCash = gSetCheck(RptSelAp!ckcSelC5(0))     'cash
    tlCntTypes.iReserv = gSetCheck(RptSelAp!ckcSelC6(1))     'reserved
    tlCntTypes.iRemnant = gSetCheck(RptSelAp!ckcSelC6(2))  'remnants
    tlCntTypes.iStandard = gSetCheck(RptSelAp!ckcSelC6(0)) 'standard
    tlCntTypes.iDR = gSetCheck(RptSelAp!ckcSelC6(3))       'Direct Response
    tlCntTypes.iPI = gSetCheck(RptSelAp!ckcSelC6(4))       'Per Inquiry
    'slCntrTypes = gBuildCntTypes()      'setup valid contract types to obtain based on user
    slCntrTypes = ""
    If tlCntTypes.iStandard Then
        slCntrTypes = "C"
    End If
    If tlCntTypes.iReserv And tgUrf(0).sResvType <> "H" Then
        slCntrTypes = slCntrTypes & "V"
    End If
    If tlCntTypes.iRemnant And tgUrf(0).sRemType <> "H" Then
        slCntrTypes = slCntrTypes & "T"
    End If
    If tlCntTypes.iDR And tgUrf(0).sDRType <> "H" Then
        slCntrTypes = slCntrTypes & "R"
    End If
    If tlCntTypes.iPI And tgUrf(0).sPIType <> "H" Then
        slCntrTypes = slCntrTypes & "Q"
    End If

    slCntrStatus = ""
    If tlCntTypes.iHold Then
        slCntrStatus = "HG"             'scheduled Holds and unsc holds
    End If
    If tlCntTypes.iOrder Then           'sch orders & uns orders
        slCntrStatus = slCntrStatus & "ON"
    End If
    'slCntrStatus = "HOGN"               'holds, orders, unsched holds, unsched orders
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    'use startwk & endwk to gather projections
    gObtainWkNo ilPjfCalType, slTYStartQtr, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    gObtainWkNo ilPjfCalType, slTYEndQtr, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)

    'Get all the Potential codes from MNF and save their adjustment percentages
    'based on the adjustment flag.
    'ReDim tlMMnf(1 To 1) As MNF
    ReDim tlMMnf(0 To 0) As MNF
    ilRet = gObtainMnfForType("P", slMnfStamp, tlMMnf())

    'Determine user selected adjust flag(s) and fill the array
    If (RptSelAp!ckcSelC8(0).Value = vbChecked) Then   'Most likely
        ilAdjFlag(0) = 1
    End If
    If (RptSelAp!ckcSelC8(1).Value = vbChecked) Then    'Optimisitc
        ilAdjFlag(1) = 2
    End If
    If (RptSelAp!ckcSelC8(2).Value = vbChecked) Then     'Pessimistic
        ilAdjFlag(2) = 3
    End If
    For ilAdjLoop = 0 To UBound(ilAdjFlag)
    'Retrieve potential code percentages from Mnf.btr and put into ilLikePct() array
        If ilAdjFlag(ilAdjLoop) = 1 Then
            'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
            For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
                If Trim$(tlMMnf(ilLoop).sName) = "A" Then
                    ilLikePct(1) = Val(tlMMnf(ilLoop).sUnitType)
                    ilLikeCode(1) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "B" Then
                    ilLikePct(2) = Val(tlMMnf(ilLoop).sUnitType)
                    ilLikeCode(2) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "C" Then
                    ilLikePct(3) = Val(tlMMnf(ilLoop).sUnitType)
                    ilLikeCode(3) = tlMMnf(ilLoop).iCode
                End If
            Next ilLoop
        End If
        If ilAdjFlag(ilAdjLoop) = 2 Then
            'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
            For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
                If Trim$(tlMMnf(ilLoop).sName) = "A" Then
                    slMnfRPU = tlMMnf(ilLoop).sRPU
                    gPDNToStr tlMMnf(ilLoop).sRPU, 2, slMnfRPU
                    ilLikePct(4) = Val(Mid$(slMnfRPU, 1, InStr(1, slMnfRPU, Chr$(46)) - 1))
                    ilLikeCode(1) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "B" Then
                    slMnfRPU = tlMMnf(ilLoop).sRPU
                    gPDNToStr tlMMnf(ilLoop).sRPU, 2, slMnfRPU
                    ilLikePct(5) = Val(Mid$(slMnfRPU, 1, InStr(1, slMnfRPU, Chr$(46)) - 1))
                    ilLikeCode(2) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "C" Then
                    slMnfRPU = tlMMnf(ilLoop).sRPU
                    gPDNToStr tlMMnf(ilLoop).sRPU, 2, slMnfRPU
                    ilLikePct(6) = Val(Mid$(slMnfRPU, 1, InStr(1, slMnfRPU, Chr$(46)) - 1))
                    ilLikeCode(3) = tlMMnf(ilLoop).iCode
                End If
            Next ilLoop
        End If
        If ilAdjFlag(ilAdjLoop) = 3 Then
            'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
            For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
                If Trim$(tlMMnf(ilLoop).sName) = "A" Then
                    slMnfSSComm = tlMMnf(ilLoop).sSSComm
                    gPDNToStr tlMMnf(ilLoop).sSSComm, 4, slMnfSSComm
                    ilLikePct(7) = Val(Mid$(slMnfSSComm, 1, InStr(1, slMnfSSComm, Chr$(46)) - 1))
                    ilLikeCode(1) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "B" Then
                    slMnfSSComm = tlMMnf(ilLoop).sSSComm
                    gPDNToStr tlMMnf(ilLoop).sSSComm, 4, slMnfSSComm
                    ilLikePct(8) = Val(Mid$(slMnfSSComm, 1, InStr(1, slMnfSSComm, Chr$(46)) - 1))
                    ilLikeCode(2) = tlMMnf(ilLoop).iCode
                ElseIf Trim$(tlMMnf(ilLoop).sName) = "C" Then
                    slMnfSSComm = tlMMnf(ilLoop).sSSComm
                    gPDNToStr tlMMnf(ilLoop).sSSComm, 4, slMnfSSComm
                    ilLikePct(9) = Val(Mid$(slMnfSSComm, 1, InStr(1, slMnfSSComm, Chr$(46)) - 1))
                    ilLikeCode(3) = tlMMnf(ilLoop).iCode
                End If
            Next ilLoop
        End If
    Next ilAdjLoop
    'Get all projection records
    'ReDim tmTPjf(1 To imPjfRecLen) As PJF
    ReDim tmTPjf(0 To 0) As PJF
    ilRet = gObtainPjf(RptSelAp, hmPjf, ilRODate(), tmTPjf())                 'Read all applicable Projection records into memory

    'Get advertisers.  build structure array of selected advt into tlSlsList
    ReDim tlSlsList(0 To RptSelAp!lbcSelection(0).ListCount)
    For ilSlsLoop = 0 To UBound(tlSlsList) - 1 Step 1
        If (RptSelAp!lbcSelection(0).Selected(ilSlsLoop)) Then
            slNameCode = tgAdvertiser(ilSlsLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAdfCode = Val(slCode)
            tlSlsList(ilAdvtFound).iAdfCode = ilAdfCode
            ilAdvtFound = ilAdvtFound + 1
        End If
    Next ilSlsLoop
    slDate = Format$(llLYGetTo, "m/d/yy")
    gPackDate slDate, ilMonth, ilYear
    tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
    tmGrf.iDate(1) = ilYear
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
    tmGrf.iStartDate(1) = ilEnterDate(1)
    ilTYYear = Val(RptSelAp!edcSelCTo.Text)           'year requested

    'Adjust projections using potential codes
    If UBound(tmTPjf) <> 0 Then
        For ilSlsLoop = 0 To ilAdvtFound Step 1
            For ilLoop = LBound(tmTPjf) To UBound(tmTPjf) - 1 Step 1
                If tmTPjf(ilLoop).iYear = ilTYYear Then
                    If tlSlsList(ilSlsLoop).iAdfCode = tmTPjf(ilLoop).iAdfCode Then
                        For ilAdjLoop = 0 To UBound(ilAdjFlag)
                            ilAdjust = ilAdjFlag(ilAdjLoop)
                            llAdjust = 0
                            For ilTemp = ilStartWk To ilEndWk Step 1
                                llAdjust = llAdjust + tmTPjf(ilLoop).lGross(ilTemp)
                            Next ilTemp
                            If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                                    'due to corp or calendar months
                                llAdjust = llAdjust + tmTPjf(ilLoop).lGross(0)
                            End If
                            If ilLastWk Then
                                llAdjust = llAdjust + tmTPjf(ilLoop).lGross(53)
                            End If
                            tmGrf.iSlfCode = tmTPjf(ilLoop).iSlfCode
                            tmGrf.iAdfCode = tgChfAP.iAdfCode
                            'access product file
                            tmSrchKey.lCode = tmTPjf(ilLoop).lPrfCode
                            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmPrf.sName = ""
                            End If
                            tmGrf.sGenDesc = tmPrf.sName
                            'tmGrf.lDollars(3) = 0
                            'tmGrf.lDollars(4) = 0
                            'tmGrf.lDollars(5) = 0
                            'tmGrf.lDollars(6) = 0
                            tmGrf.lDollars(2) = 0
                            tmGrf.lDollars(3) = 0
                            tmGrf.lDollars(4) = 0
                            tmGrf.lDollars(5) = 0
                            If ilAdjLoop = 0 And ilAdjust = 1 Then          'most likely
                                If tmTPjf(ilLoop).iMnfBus = ilLikeCode(1) Then
                                    'tmGrf.lDollars(3) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(1) / 100))
                                    tmGrf.lDollars(2) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(1) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(2) Then
                                    'tmGrf.lDollars(4) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(2) / 100))
                                    tmGrf.lDollars(3) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(2) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(3) Then
                                    'tmGrf.lDollars(5) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(3) / 100))
                                    tmGrf.lDollars(4) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(3) / 100))
                                End If
                                'tmGrf.iPerGenl(1) = 1
                                'tmGrf.iPerGenl(3) = ilLikePct(1)
                                'tmGrf.iPerGenl(4) = ilLikePct(2)
                                'tmGrf.iPerGenl(5) = ilLikePct(3)
                                tmGrf.iPerGenl(0) = 1
                                tmGrf.iPerGenl(2) = ilLikePct(1)
                                tmGrf.iPerGenl(3) = ilLikePct(2)
                                tmGrf.iPerGenl(4) = ilLikePct(3)
                                tmGrf.iSlfCode = tmTPjf(ilLoop).iSlfCode
                                tmGrf.iAdfCode = tmTPjf(ilLoop).iAdfCode

                                'write the record only if $ fields non-zero
                                'If (tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                                If (tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If

                            ElseIf ilAdjLoop = 1 And ilAdjust = 2 Then              'optimistic
                                If tmTPjf(ilLoop).iMnfBus = ilLikeCode(1) Then
                                    'tmGrf.lDollars(3) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(4) / 100))
                                    tmGrf.lDollars(2) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(4) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(2) Then
                                    'tmGrf.lDollars(4) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(5) / 100))
                                    tmGrf.lDollars(3) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(5) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(3) Then
                                    'tmGrf.lDollars(5) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(6) / 100))
                                    tmGrf.lDollars(4) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(6) / 100))
                                End If
                                'tmGrf.iPerGenl(1) = 2
                                'tmGrf.iPerGenl(3) = ilLikePct(4)
                                'tmGrf.iPerGenl(4) = ilLikePct(5)
                                'tmGrf.iPerGenl(5) = ilLikePct(6)
                                tmGrf.iPerGenl(0) = 2
                                tmGrf.iPerGenl(2) = ilLikePct(4)
                                tmGrf.iPerGenl(3) = ilLikePct(5)
                                tmGrf.iPerGenl(4) = ilLikePct(6)
                                tmGrf.iSlfCode = tmTPjf(ilLoop).iSlfCode
                                tmGrf.iAdfCode = tmTPjf(ilLoop).iAdfCode

                                'write the record only if $ fields non-zero
                                'If (tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                                If (tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            ElseIf ilAdjLoop = 2 And ilAdjust = 3 Then              'pessimistic
                                If tmTPjf(ilLoop).iMnfBus = ilLikeCode(1) Then
                                    'tmGrf.lDollars(3) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(7) / 100))
                                    tmGrf.lDollars(2) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(7) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(2) Then
                                    'tmGrf.lDollars(4) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(8) / 100))
                                    tmGrf.lDollars(3) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(8) / 100))
                                ElseIf tmTPjf(ilLoop).iMnfBus = ilLikeCode(3) Then
                                    'tmGrf.lDollars(5) = llAdjust
                                    'tmGrf.lDollars(6) = (llAdjust * (ilLikePct(9) / 100))
                                    tmGrf.lDollars(4) = llAdjust
                                    tmGrf.lDollars(5) = (llAdjust * (ilLikePct(9) / 100))
                                End If
                                'tmGrf.iPerGenl(1) = 3
                                'tmGrf.iPerGenl(3) = ilLikePct(7)
                                'tmGrf.iPerGenl(4) = ilLikePct(8)
                                'tmGrf.iPerGenl(5) = ilLikePct(9)
                                tmGrf.iPerGenl(0) = 3
                                tmGrf.iPerGenl(2) = ilLikePct(7)
                                tmGrf.iPerGenl(3) = ilLikePct(8)
                                tmGrf.iPerGenl(4) = ilLikePct(9)
                                tmGrf.iSlfCode = tmTPjf(ilLoop).iSlfCode
                                tmGrf.iAdfCode = tmTPjf(ilLoop).iAdfCode

                                'write the record only if $ fields non-zero
                                'If (tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                                If (tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            End If
                        Next ilAdjLoop
                    End If
                End If                              'Pjf.year = ilyear
            Next ilLoop
        Next ilSlsLoop
    End If

    'Search History and Receivables looking for AN (adjustments) to offset
    ilRet = gObtainPhfRvf(RptSelAp, slLYStartYr, slTYEndQtr, tlTranType, tlRvf(), 0)
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
               'dan M 7-14-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
            llDate = gDateValue(slStr)
            ilTY = False
            ilFoundOne = False
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)
            'Dan M 8-13-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
            blFailedMatchNtrOrHardCost = False
            'Dan M 8-13-8 don't allow installment option "I"
            blFailedBecauseInstallment = False
            If tmRvf.sType = "I" Then
                blFailedBecauseInstallment = True
            End If
            If ((blIncludeNTR) Xor (blIncludeHardCost)) And tmRvf.iMnfItem > 0 Then      'one or the other is true, but not both (if both true, don't have to isolate anything)
                ilRet = gIsItHardCost(tmRvf.iMnfItem, tmMnfNtr())
            'if is hard cost but blincludentr  or isn't hard cost but blincludehardcost then it needs to be removed. set failedmatchntrorhardcost true
                If (ilRet And blIncludeNTR) Or ((Not ilRet) And blIncludeHardCost) Then
                    blFailedMatchNtrOrHardCost = True
                End If
            End If
            If Not (blFailedMatchNtrOrHardCost Or blFailedBecauseInstallment) Then  'if both false, continue
                If Not (RptSelAp!ckcAll.Value = vbChecked) Then
                    'Find the selective advt
                    For ilSlsLoop = 0 To ilAdvtFound Step 1
                        If tlSlsList(ilSlsLoop).iAdfCode = tmRvf.iAdfCode Then
                            ilFoundOne = True
                            Exit For
                        End If
                    Next ilSlsLoop
                Else
                    ilFoundOne = True               'all advt selected, continue to test dates
                End If
            End If
            If ilFoundOne Then
                ilFoundOne = False          'find valid dates
                If llDate >= llTYDates(1) And llDate <= llTYDates(2) Then
                    ilTY = True
                    If llDateEntered <= llTYGetTo Then  'replaced below with this Dan M 8-13-08
                    'If llDate <= llTYGetTo Then
                        'llGenlDates(ilTemp) = llTYStarts(ilTemp)
                        'If llDate >= llTYDates(1) And llDate <= llTYDates(2) Then       'added equal sign to llTYDates(2) Dan M 8-14-08, then figured out don't need
                            ilFoundOne = True
                            gPDNToLong tmRvf.sGross, llProject(1)
                        'End If
                    End If
                'if trans date not within current year, assume last year
                Else
                    'If llDate <= llLYGetTo Then
                    If llDateEntered <= llLYGetTo Then
                        If llDate >= llLYDates(1) And llDate <= llLYDates(2) Then
                            ilFoundOne = True
                            gPDNToLong tmRvf.sGross, llProject(1)
                        End If
                    End If
                End If
            End If
            If ilFoundOne Then
                ilFoundOne = False
                'access product file
                tmSrchKey.lCode = tmRvf.lPrfCode
                ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    tmPrf.sName = ""
                End If

                'Create one record per adjustment requested (most likely, optimistic, pessimistic) for this contract
                For ilAdjLoop = 0 To UBound(ilAdjFlag)
                    ilAdjust = ilAdjFlag(ilAdjLoop)
                    'If ilAdjust > 0 Then
                        'Gather all Slsp projection records for the matching rollover date (exclude current records)
                        'Accumulate the $ projected into the buckets
                        llGross = llProject(1) \ 100                        'drop pennies

                        'Write the record to Grf.btr
                        'Calculate the adjust dollars based on potential code/percentage
                        'tmGrf.lDollars(1) = 0
                        'tmGrf.lDollars(2) = 0
                        tmGrf.lDollars(0) = 0
                        tmGrf.lDollars(1) = 0
                        If ilAdjLoop = 0 And ilAdjust = 1 Then
                            'tmGrf.iPerGenl(1) = 1
                            'tmGrf.iPerGenl(3) = ilLikePct(1)
                            'tmGrf.iPerGenl(4) = ilLikePct(2)
                            'tmGrf.iPerGenl(5) = ilLikePct(3)
                            tmGrf.iPerGenl(0) = 1
                            tmGrf.iPerGenl(2) = ilLikePct(1)
                            tmGrf.iPerGenl(3) = ilLikePct(2)
                            tmGrf.iPerGenl(4) = ilLikePct(3)
                            If Not ilTY Then             'past year  (holds & orders are combined)
                                'if beyond the effective date, it's still actuals for the qtr
                                'this record is last years qtr, accum the actuals
                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                    'show it in the same time last year column
                                    'tmGrf.lDollars(1) = llGross
                                    tmGrf.lDollars(0) = llGross
                                Else
                                    'tmGrf.lDollars(1) = 0
                                    tmGrf.lDollars(0) = 0
                                End If
                            Else                                'current year, holds and orders are added together
                                'tmGrf.lDollars(2) = llGross
                                tmGrf.lDollars(1) = llGross
                            End If

                        ElseIf ilAdjLoop = 1 And ilAdjust = 2 Then
                            'tmGrf.iPerGenl(1) = 2
                            'tmGrf.iPerGenl(3) = ilLikePct(4)
                            'tmGrf.iPerGenl(4) = ilLikePct(5)
                            'tmGrf.iPerGenl(5) = ilLikePct(6)
                            tmGrf.iPerGenl(0) = 2
                            tmGrf.iPerGenl(2) = ilLikePct(4)
                            tmGrf.iPerGenl(3) = ilLikePct(5)
                            tmGrf.iPerGenl(4) = ilLikePct(6)
                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                'if beyond the effective date, it's still actuals for the qtr
                                'this record is last years qtr, accum the actuals
                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                    'show it in the same time last year column
                                    'tmGrf.lDollars(1) = llGross
                                    tmGrf.lDollars(0) = llGross
                                Else
                                    'tmGrf.lDollars(1) = 0
                                    tmGrf.lDollars(0) = 0
                                End If
                            Else                                'current year, holds and orders are added together
                                'tmGrf.lDollars(2) = llGross
                                tmGrf.lDollars(1) = llGross
                            End If
                        ElseIf ilAdjLoop = 2 And ilAdjust = 3 Then
                            'tmGrf.iPerGenl(1) = 3
                            'tmGrf.iPerGenl(3) = ilLikePct(7)
                            'tmGrf.iPerGenl(4) = ilLikePct(8)
                            'tmGrf.iPerGenl(5) = ilLikePct(9)
                            tmGrf.iPerGenl(0) = 3
                            tmGrf.iPerGenl(2) = ilLikePct(7)
                            tmGrf.iPerGenl(3) = ilLikePct(8)
                            tmGrf.iPerGenl(4) = ilLikePct(9)
                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                'if beyond the effective date, it's still actuals for the qtr
                                'this record is last years qtr, accum the actuals
                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                    'show it in the same time last year column
                                    'tmGrf.lDollars(1) = llGross
                                    tmGrf.lDollars(0) = llGross
                                Else
                                    'tmGrf.lDollars(1) = 0
                                    tmGrf.lDollars(0) = 0
                                End If
                            Else                                'current year, holds and orders are added together
                                'tmGrf.lDollars(2) = llGross
                                tmGrf.lDollars(1) = llGross
                            End If
                        End If
                        tmGrf.iSlfCode = tmRvf.iSlfCode
                        tmGrf.iAdfCode = tmRvf.iAdfCode
                        tmGrf.sGenDesc = Trim$(tmPrf.sName)
                        'tmGrf.lDollars(3) = 0
                        'tmGrf.lDollars(4) = 0
                        'tmGrf.lDollars(5) = 0
                        'tmGrf.lDollars(6) = 0
                        tmGrf.lDollars(2) = 0
                        tmGrf.lDollars(3) = 0
                        tmGrf.lDollars(4) = 0
                        tmGrf.lDollars(5) = 0
                        slDate = Format$(llLYGetTo, "m/d/yy")
                        gPackDate slDate, ilMonth, ilYear
                        tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
                        tmGrf.iDate(1) = ilYear
                        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                        tmGrf.iGenDate(1) = igNowDate(1)
                        'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
                        'tmGrf.iGenTime(1) = igNowTime(1)
                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                        tmGrf.lGenTime = lgNowTime
                        tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
                        tmGrf.iStartDate(1) = ilEnterDate(1)
                        'write the record only if $ fields non-zero
                        'If (tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                        If (tmGrf.lDollars(0) + tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                Next ilAdjLoop

            End If      'ilfoundOne
        End If      'single contract selectivity
    Next llRvfLoop



    'Process last year, then this year.  Get all contracts for the active quarter dates and project their $ from the flights.
    For ilPastFut = 1 To 2 Step 1
        If ilPastFut = 1 Then                       'past
            'slStartDate = Format$(llLYDates(1), "m/d/yy")       'gather all cntrs whose start/end dates fall within requested qtr (last year)
            slStartDate = slLYStartYr                           'always gather all contr from beginning of corp or std year to
                                                                'get correct mods for the pacing
            'slEndDate = Format$(llLYDates(2), "m/d/yy")
            slEndDate = slLYEndYr
            llStartDates(1) = llLYDates(1)
            llStartDates(2) = llLYDates(2)
            llEnterFrom = llLYGetFrom                           'gather all cntrs whose entered date falls within these dates
            llEnterTo = llLYGetTo
        Else                                         'current
            'slStartDate = Format$(llTYDates(1), "m/d/yy")        'gather all cntrs whose start/end dates fall within requested qtr (this year)
            slStartDate = slTYStartYr
            'slEndDate = Format$(llTYDates(2), "m/d/yy")
            slEndDate = slTYEndYr
            llStartDates(1) = llTYDates(1)
            llStartDates(2) = llTYDates(2)
            llEnterFrom = llTYGetFrom           'gather cnts whose entered date falls within these dates
            llEnterTo = llTYGetTo
        End If

        'Build array of possible contracts that fall into last year or this years quarter and build into array tlChfAdvtExt
        ReDim Preserve tlChfAdvtExt(0 To 0)
        ilRet = gObtainCntrForDate(RptSelAp, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

        For ilSlsLoop = 0 To ilAdvtFound Step 1                 'loop on advertisers selected
            'For ilChfLoop = 1 To UBound(tlChfAdvtExt) - 1 Step 1  'loop on contracts for the year
            For ilChfLoop = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1  'loop on contracts for the year
                   '7-11-08 added single contract selectivity dan M
                If llSingleContract = NOT_SELECTED Or llSingleContract = tlChfAdvtExt(ilChfLoop).lCntrNo Then

                    If tlSlsList(ilSlsLoop).iAdfCode = tlChfAdvtExt(ilChfLoop).iAdfCode Then        'process only if contract has matching selected advt
                        llContrCode = tlChfAdvtExt(ilChfLoop).lCode
                        'Got the correct header that is equal or prior to the effective date entered
                        llContrCode = gPaceCntr(tlChfAdvtExt(ilChfLoop).lCntrNo, llEnterTo, hmCHF, tmChf)
                        ilFound = False
                        If llContrCode > 0 Then

                            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfAP, tgClfAP(), tgCffAP())   'get the latest version of this contract
                            gUnpackDateLong tgChfAP.iOHDDate(0), tgChfAP.iOHDDate(1), llAdjust      'date entered
                            If ilPastFut = 2 Then       'if current, need to test entered date against the requested effective
                                If llAdjust <= llEnterTo Then       'entered date must be entered thru effectve date
                                    ilFound = True
                                End If
                            Else                        'Past
                                ilFound = True          'past get all cnts affecting the qtr to get actuals as well as same wee last year
                            End If
                        End If
                        llProject(1) = 0                'init bkts to accum qtr $ for this line
                        If (ilFound) And ((tgChfAP.iPctTrade = 100 And tlCntTypes.iTrade) Or (tgChfAP.iPctTrade < 100 And tlCntTypes.iCash)) Then                'Loop thru all lines and project their $ from the flights

                            For ilClf = LBound(tgClfAP) To UBound(tgClfAP) - 1 Step 1
                                'llProject(1) = 0                'init bkts to accum qtr $ for this line
                                tmClf = tgClfAP(ilClf).ClfRec
                                If slAirOrder = "O" Then                'invoice all contracts as ordered
                                    If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                                        gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfAP(), tgCffAP()
                                    End If
                                Else                                    'inv all contracts as aired
                                    If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                                        'if hidden, will project if assoc. package is set to invoice as aired (real)
                                        For ilTemp = LBound(tgClfAP) To UBound(tgClfAP) - 1    'find the assoc. pkg line for these hidden
                                            If tmClf.iPkLineNo = tgClfAP(ilTemp).ClfRec.iLine Then
                                                If tgClfAP(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                                                    gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfAP(), tgCffAP() 'pkg bills as aired, project the hidden line
                                                End If
                                                Exit For
                                            End If
                                        Next ilTemp
                                    Else                            'conventional, VV, or Pkg line
                                        If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                                                                    'it has already been projected above with the hidden line
                                            gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfAP(), tgCffAP()
                                        End If
                                    End If
                                End If
                            Next ilClf                      'process nextline
                        End If                              'ilfound

                        If ilFound Then
                            'Create one record per adjustment requested (most likely, optimistic, pessimistic) for this contract
                            For ilAdjLoop = 0 To UBound(ilAdjFlag)
                                ilAdjust = ilAdjFlag(ilAdjLoop)
                                'If ilAdjust > 0 Then
                                    'Gather all Slsp projection records for the matching rollover date (exclude current records)
                                    'Accumulate the $ projected into the buckets
                                    llGross = llProject(1) \ 100                        'drop pennies
                                    'If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                        'if beyond the effective date, it's still actuals for the qtr
                                        'this record is last years qtr, accum the actuals
                                    '    If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                            'show it in the same time last year column
                                    '        tlSlsList(ilSlsLoop).lLYWeek = llGross
                                    '    Else
                                    '        tlSlsList(ilSlsLoop).lLYWeek = 0
                                    '    End If
                                    'Else                                'current year, holds and orders are added together
                                    '    tlSlsList(ilSlsLoop).lTYAct = llGross
                                    'End If

                                    'Write the record to Grf.btr
                                    'Calculate the adjust dollars based on potential code/percentage
                                    'tmGrf.lDollars(1) = 0
                                    'tmGrf.lDollars(2) = 0
                                    tmGrf.lDollars(0) = 0
                                    tmGrf.lDollars(1) = 0
                                    If ilAdjLoop = 0 And ilAdjust = 1 Then
                                        'tmGrf.iPerGenl(1) = 1
                                        'tmGrf.iPerGenl(3) = ilLikePct(1)
                                        'tmGrf.iPerGenl(4) = ilLikePct(2)
                                        'tmGrf.iPerGenl(5) = ilLikePct(3)
                                        tmGrf.iPerGenl(0) = 1
                                        tmGrf.iPerGenl(2) = ilLikePct(1)
                                        tmGrf.iPerGenl(3) = ilLikePct(2)
                                        tmGrf.iPerGenl(4) = ilLikePct(3)
                                        If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                            'if beyond the effective date, it's still actuals for the qtr
                                            'this record is last years qtr, accum the actuals
                                            If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                'show it in the same time last year column
                                                'tmGrf.lDollars(1) = llGross
                                                tmGrf.lDollars(0) = llGross
                                            Else
                                                'tmGrf.lDollars(1) = 0
                                                tmGrf.lDollars(0) = 0
                                            End If
                                        Else                                'current year, holds and orders are added together
                                            'tmGrf.lDollars(2) = llGross
                                            tmGrf.lDollars(1) = llGross
                                        End If

                                        'tmGrf.lDollars(1) = tlSlsList(ilSlsLoop).lLYWeek
                                        'tmGrf.lDollars(2) = tlSlsList(ilSlsLoop).lTYAct
                                    ElseIf ilAdjLoop = 1 And ilAdjust = 2 Then
                                        'tmGrf.iPerGenl(1) = 2
                                        'tmGrf.iPerGenl(3) = ilLikePct(4)
                                        'tmGrf.iPerGenl(4) = ilLikePct(5)
                                        'tmGrf.iPerGenl(5) = ilLikePct(6)
                                        tmGrf.iPerGenl(0) = 2
                                        tmGrf.iPerGenl(2) = ilLikePct(4)
                                        tmGrf.iPerGenl(3) = ilLikePct(5)
                                        tmGrf.iPerGenl(4) = ilLikePct(6)
                                        If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                            'if beyond the effective date, it's still actuals for the qtr
                                            'this record is last years qtr, accum the actuals
                                            If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                'show it in the same time last year column
                                                'tmGrf.lDollars(1) = llGross
                                                tmGrf.lDollars(0) = llGross
                                            Else
                                                'tmGrf.lDollars(1) = 0
                                                tmGrf.lDollars(0) = 0
                                            End If
                                        Else                                'current year, holds and orders are added together
                                            'tmGrf.lDollars(2) = llGross
                                            tmGrf.lDollars(1) = llGross
                                        End If
                                        'tmGrf.lDollars(1) = tlSlsList(ilSlsLoop).lLYWeek
                                        'tmGrf.lDollars(2) = tlSlsList(ilSlsLoop).lTYAct
                                    ElseIf ilAdjLoop = 2 And ilAdjust = 3 Then
                                        'tmGrf.iPerGenl(1) = 3
                                        'tmGrf.iPerGenl(3) = ilLikePct(7)
                                        'tmGrf.iPerGenl(4) = ilLikePct(8)
                                        'tmGrf.iPerGenl(5) = ilLikePct(9)
                                        tmGrf.iPerGenl(0) = 3
                                        tmGrf.iPerGenl(2) = ilLikePct(7)
                                        tmGrf.iPerGenl(3) = ilLikePct(8)
                                        tmGrf.iPerGenl(4) = ilLikePct(9)
                                        If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                            'if beyond the effective date, it's still actuals for the qtr
                                            'this record is last years qtr, accum the actuals
                                            If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                'show it in the same time last year column
                                                'tmGrf.lDollars(1) = llGross
                                                tmGrf.lDollars(0) = llGross
                                            Else
                                                'tmGrf.lDollars(1) = 0
                                                tmGrf.lDollars(0) = 0
                                            End If
                                        Else                                'current year, holds and orders are added together
                                            'tmGrf.lDollars(2) = llGross
                                            tmGrf.lDollars(1) = llGross
                                        End If
                                        'tmGrf.lDollars(1) = tlSlsList(ilSlsLoop).lLYWeek
                                        'tmGrf.lDollars(2) = tlSlsList(ilSlsLoop).lTYAct
                                    End If
                                    'If UBound(tmTPjf) = 0 Then
                                        'tlSlsList(ilSlsLoop).iSlfCode = tgChfAP.iSlfCode(0)
                                        'tmGrf.iSlfCode = tlSlsList(ilSlsLoop).iSlfCode
                                        tmGrf.iSlfCode = tgChfAP.iSlfCode(0)
                                        'tlSlsList(ilSlsLoop).iAdfCode = tgChfAP.iAdfCode
                                        'tmGrf.iAdfCode = tlSlsList(ilSlsLoop).iAdfCode
                                        tmGrf.iAdfCode = tgChfAP.iAdfCode
                                        tmGrf.sGenDesc = Trim$(tgChfAP.sProduct)
                                        'tmGrf.lDollars(3) = 0
                                        'tmGrf.lDollars(4) = 0
                                        'tmGrf.lDollars(5) = 0
                                        'tmGrf.lDollars(6) = 0
                                        tmGrf.lDollars(2) = 0
                                        tmGrf.lDollars(3) = 0
                                        tmGrf.lDollars(4) = 0
                                        tmGrf.lDollars(5) = 0
                                    'End If
                                    slDate = Format$(llLYGetTo, "m/d/yy")
                                    gPackDate slDate, ilMonth, ilYear
                                    tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
                                    tmGrf.iDate(1) = ilYear
                                    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                                    tmGrf.iGenDate(1) = igNowDate(1)
                                    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
                                    'tmGrf.iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmGrf.lGenTime = lgNowTime
                                    tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
                                    tmGrf.iStartDate(1) = ilEnterDate(1)
                                    'write the record only if $ fields non-zero
                                    'If (tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                                    If (tmGrf.lDollars(0) + tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                    End If
                                'End If               'iladjust > 0   (mostlikely, pessimistic or optimistic requested)
                            Next ilAdjLoop
                        End If                   'ilfound
                                   ' Dan M 7-03-08 Add NTR/Hard Cost option
                        'Does user want to see HardCost/NTR? Contract# <> 0? have hard cost? Not pure trade?    added ilfound 7-28-08
                        If (blIncludeNTR Or blIncludeHardCost) And (tgChfAP.sNTRDefined = "Y") And (tlChfAdvtExt(ilChfLoop).iPctTrade <> 100) And ilFound Then
                        'call routine to fill array with choice
                            gNtrByContract llContrCode, llStartDates(1), llStartDates(2), tlNTRInfo(), tmMnfNtr(), hmSbf, blIncludeNTR, blIncludeHardCost, RptSelAp
                            ilLowerboundNTR = LBound(tlNTRInfo)
                            ilUpperboundNTR = UBound(tlNTRInfo)
                            'blNTRWithTotal = False         'moved into loop
                            For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                         'look at each ntr record's date to see if falls into specific time period. changed < to <= 7/30/08
                                If tlNTRInfo(ilNTRCounter).lSbfDate >= llStartDates(1) And tlNTRInfo(ilNTRCounter).lSbfDate <= llStartDates(2) Then
                            blNTRWithTotal = False
                            'flag so won't write record if all values are 0
                                    If tlNTRInfo(ilNTRCounter).lSBFTotal > 0 Then
                '            'clear value
                                        llProject(1) = 0
                                        blNTRWithTotal = True
                                        llProject(1) = llProject(1) + tlNTRInfo(ilNTRCounter).lSBFTotal 'looks like running total: isn't
                                    End If
                                End If
                            'send to routine to write to grf
                                If blNTRWithTotal = True Then
                                ' copied from above
                                    For ilAdjLoop = 0 To UBound(ilAdjFlag)
                                        ilAdjust = ilAdjFlag(ilAdjLoop)
                                        'Gather all Slsp projection records for the matching rollover date (exclude current records)
                                        'Accumulate the $ projected into the buckets
                                        llGross = llProject(1) \ 100                        'drop pennies
                                        'Write the record to Grf.btr
                                        'Calculate the adjust dollars based on potential code/percentage
                                        'tmGrf.lDollars(1) = 0
                                        'tmGrf.lDollars(2) = 0
                                        tmGrf.lDollars(0) = 0
                                        tmGrf.lDollars(1) = 0
                                        If ilAdjLoop = 0 And ilAdjust = 1 Then
                                            'tmGrf.iPerGenl(1) = 1
                                            'tmGrf.iPerGenl(3) = ilLikePct(1)
                                            'tmGrf.iPerGenl(4) = ilLikePct(2)
                                            'tmGrf.iPerGenl(5) = ilLikePct(3)
                                            tmGrf.iPerGenl(0) = 1
                                            tmGrf.iPerGenl(2) = ilLikePct(1)
                                            tmGrf.iPerGenl(3) = ilLikePct(2)
                                            tmGrf.iPerGenl(4) = ilLikePct(3)
                                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                                'if beyond the effective date, it's still actuals for the qtr
                                                'this record is last years qtr, accum the actuals
                                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year

                                                                                    'show it in the same time last year column
                                                    'tmGrf.lDollars(1) = llGross
                                                    tmGrf.lDollars(0) = llGross
                                                Else
                                                    'tmGrf.lDollars(1) = 0
                                                    tmGrf.lDollars(0) = 0
                                                End If
                                            Else                                'current year, holds and orders are added together
                                                'tmGrf.lDollars(2) = llGross
                                                tmGrf.lDollars(1) = llGross
                                            End If

                                            'tmGrf.lDollars(1) = tlSlsList(ilSlsLoop).lLYWeek
                                            'tmGrf.lDollars(2) = tlSlsList(ilSlsLoop).lTYAct
                                        ElseIf ilAdjLoop = 1 And ilAdjust = 2 Then
                                            'tmGrf.iPerGenl(1) = 2
                                            'tmGrf.iPerGenl(3) = ilLikePct(4)
                                            'tmGrf.iPerGenl(4) = ilLikePct(5)
                                            'tmGrf.iPerGenl(5) = ilLikePct(6)
                                            tmGrf.iPerGenl(0) = 2
                                            tmGrf.iPerGenl(2) = ilLikePct(4)
                                            tmGrf.iPerGenl(3) = ilLikePct(5)
                                            tmGrf.iPerGenl(4) = ilLikePct(6)
                                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                                'if beyond the effective date, it's still actuals for the qtr
                                                'this record is last years qtr, accum the actuals
                                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                    'show it in the same time last year column
                                                    'tmGrf.lDollars(1) = llGross
                                                    tmGrf.lDollars(0) = llGross
                                                Else
                                                    'tmGrf.lDollars(1) = 0
                                                    tmGrf.lDollars(0) = 0
                                                End If
                                            Else                                'current year, holds and orders are added together
                                                'tmGrf.lDollars(2) = llGross
                                                tmGrf.lDollars(1) = llGross
                                            End If
                                        ElseIf ilAdjLoop = 2 And ilAdjust = 3 Then
                                            'tmGrf.iPerGenl(1) = 3
                                            'tmGrf.iPerGenl(3) = ilLikePct(7)
                                            'tmGrf.iPerGenl(4) = ilLikePct(8)
                                            'tmGrf.iPerGenl(5) = ilLikePct(9)
                                            tmGrf.iPerGenl(0) = 3
                                            tmGrf.iPerGenl(2) = ilLikePct(7)
                                            tmGrf.iPerGenl(3) = ilLikePct(8)
                                            tmGrf.iPerGenl(4) = ilLikePct(9)
                                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                                'if beyond the effective date, it's still actuals for the qtr
                                                'this record is last years qtr, accum the actuals
                                                If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                    'show it in the same time last year column
                                                    'tmGrf.lDollars(1) = llGross
                                                    tmGrf.lDollars(0) = llGross
                                                Else
                                                    'tmGrf.lDollars(1) = 0
                                                    tmGrf.lDollars(0) = 0
                                                End If
                                            Else                                'current year, holds and orders are added together
                                                'tmGrf.lDollars(2) = llGross
                                                tmGrf.lDollars(1) = llGross
                                            End If
                                        End If
                                        ' dan note: changed tmchf to tgchfap
                                        tmGrf.iSlfCode = tgChfAP.iSlfCode(0)
                                        tmGrf.iAdfCode = tgChfAP.iAdfCode
                                        tmGrf.sGenDesc = Trim$(tgChfAP.sProduct)
                                        'tmGrf.lDollars(3) = 0
                                        'tmGrf.lDollars(4) = 0
                                        'tmGrf.lDollars(5) = 0
                                        'tmGrf.lDollars(6) = 0
                                        tmGrf.lDollars(2) = 0
                                        tmGrf.lDollars(3) = 0
                                        tmGrf.lDollars(4) = 0
                                        tmGrf.lDollars(5) = 0
                                        slDate = Format$(llLYGetTo, "m/d/yy")
                                        gPackDate slDate, ilMonth, ilYear
                                        tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
                                        tmGrf.iDate(1) = ilYear
                                        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                                        tmGrf.iGenDate(1) = igNowDate(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tmGrf.lGenTime = lgNowTime
                                        'don't need effective date, right? dan
                                        'tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
                                        'tmGrf.iStartDate(1) = ilEnterDate(1)
                                        'write the record only if $ fields non-zero
                                        'If (tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) <> 0 Then
                                        If (tmGrf.lDollars(0) + tmGrf.lDollars(1) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5)) <> 0 Then
                                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                        End If
                                    Next ilAdjLoop
                                End If      'if blntrwithtotal = true
                            Next ilNTRCounter
                        End If          'Ntr/hard cost
                    End If                       'tlSlsList(ilSlsLoop).iAdfCode = tlChfAdvtExt(ilChfLoop).iAdfCode
                End If              'contract selectivity
            Next ilChfLoop
        Next ilSlsLoop                       'goto next contract record

        Erase tlChfAdvtExt                      'Make sure last years contrcts are erased, go process this year
        sgCntrForDateStamp = ""
    Next ilPastFut                              'go from past to future dates


    'Cleanup
    Erase tlSlsList
    Erase tlMMnf
    Erase ilAdjFlag
    Erase tmTPjf
    Erase tlChfAdvtExt
    Erase tmMnfNtr
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmMnf)
End Sub
