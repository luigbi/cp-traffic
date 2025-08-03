Attribute VB_Name = "RPTCRPJ"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrpj.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
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
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
'Dim lmStartDates(1 To 13)  As Long
Dim lmStartDates(0 To 13)  As Long      'Index zero ignored
'Dim lmEndDates(1 To 13) As Long
Dim lmEndDates(0 To 13) As Long         'Index zero ignored
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlf As SLF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmPrf As Integer            'Product file handle
Dim tmPrf As PRF                'PRF record image
Dim tmSrchKey As LONGKEY0
Dim imPrfRecLen As Integer        'PRF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmPjf As PJF                  'Slsp Projections
Dim hmPjf As Integer
Dim imPjfRecLen As Integer        'PJF record length
Const LBONE = 1

'
'
'                   gCrProj - Create Salesperson, Vehicle or Office
'                       Projections for 6 months of data
'                       Current Projections - get all records with rollover date = 0
'                       Past Projections - use date entered and find closest rollover
'                           date to produce data.
'                       If Differences - (for past option only):  Get the
'                           records with closest rollover date to user entered date,
'                           then retrieve one week prior data.
'
'                   7/13/98
'       10-17-03 fix subscript out of range when potential codes are not used
Sub gCrProj()
Dim ilRet As Integer
Dim slPotn As String
Dim ilLoop As Integer
Dim ilFound As Integer
Dim ilTemp As Integer
Dim ilMnfLoop As Integer
Dim slMnfStamp As String
Dim slStr As String
Dim llProjYearEnd As Long
Dim llProjYearStart As Long
Dim llTotalGross As Long            'total $ for 6 month projection
Dim ilMonth As Integer
Dim llAdjust As Long
Dim ilLoopProj As Integer                   '
Dim ilMaxOption As Integer
Dim slCurrDate As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim slNameCode As String
Dim slCode As String
Dim llClosestDate As Long
Dim slBaseDate As String
Dim ilLoopWks As Integer
Dim ilListIndex As Integer           'report selected
ReDim ilRODate(0 To 1) As Integer
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmPjf
        btrDestroy hmGrf
        Exit Sub
    End If
    imPjfRecLen = Len(tmPjf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmGrf
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSlf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmGrf
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmPrf
        btrDestroy hmSlf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmGrf
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)
    ilListIndex = RptSelPJ!lbcRptType.ListIndex

    If RptSelPJ!rbcSelCSelect(0).Value Then     'corporate selected
        slBaseDate = "C"
    Else
        slBaseDate = "S"
    End If
    'ReDim tlMMnf(1 To 1) As MNF
    ReDim tlMMnf(0 To 0) As MNF
    'get all the Potential codes from MNF
    ilRet = gObtainMnfForType("P", slMnfStamp, tlMMnf())
    'ReDim tlPotn(1 To 1) As ADJUSTLIST          'arry of potential codes and their percentages
    ReDim tlPotn(0 To 1) As ADJUSTLIST          'arry of potential codes and their percentages. Index zero ignored
    slPotn = ""
    'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
    For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
        'Bypass any potential codes that isnt an A, B or C
        If Trim$(tlMMnf(ilLoop).sName) = "A" Or Trim$(tlMMnf(ilLoop).sName) = "B" Or Trim$(tlMMnf(ilLoop).sName) = "C" Then
            For ilMnfLoop = LBONE To UBound(tlPotn) - 1 Step 1
                If tlMMnf(ilLoop).iCode = tlPotn(ilMnfLoop).iVefCode Then    'see if this potn code has been created in mem yet
                    ilFound = True
                    Exit For
                End If
            Next ilMnfLoop
            If Not ilFound Then
                ilMnfLoop = UBound(tlPotn)
                tlPotn(ilMnfLoop).iVefCode = tlMMnf(ilLoop).iCode
                tlPotn(ilMnfLoop).lProject(1) = Val(tlMMnf(ilLoop).sUnitType)            'most likely percentage
                gPDNToLong tlMMnf(ilLoop).sRPU, tlPotn(ilMnfLoop).lProject(2)          'optimistc percentage
                tlPotn(ilMnfLoop).lProject(2) = tlPotn(ilMnfLoop).lProject(2) \ 100
                gPDNToLong tlMMnf(ilLoop).sSSComm, tlPotn(ilMnfLoop).lProject(3)           'pessimistic percentage
                tlPotn(ilMnfLoop).lProject(3) = tlPotn(ilMnfLoop).lProject(3) \ 10000
                'ReDim Preserve tlPotn(1 To UBound(tlPotn) + 1)
                ReDim Preserve tlPotn(0 To UBound(tlPotn) + 1)
                slPotn = Trim$(slPotn) & Trim$(tlMMnf(ilLoop).sName)
            End If
        End If
    Next ilLoop

    'if no date entered, get current date so that we can determine what months to produce output for
    If RptSelPJ!rbcSelCInclude(0).Value Then    'current (get blank, or zero, rollover dates)
        gUnpackCurrDateTime slCurrDate, slStr, slMonth, slDay, slYear                                'default to todays date
        ilRODate(0) = 0                     'current projections dont have a rollover date
        ilRODate(1) = 0
    Else                                        'past, setup the user entered past date
        slCurrDate = RptSelPJ!edcSelCFrom.Text           'Past date entered
        gObtainYearMonthDayStr slCurrDate, True, slYear, slMonth, slDay
        gGetRollOverDate RptSelPJ, 2, slCurrDate, llClosestDate   'get closest rollover date one week ago
        slStr = Format$(llClosestDate, "m/d/yy")
        gPackDate slStr, ilRODate(0), ilRODate(1)
    End If
    gCorpStdDates slBaseDate, slCurrDate, ilRet, lmStartDates(), lmEndDates()           '1st 6 months are dates to report

    'Obtain the start month of the quarter for report heading
    slStr = Format$(lmStartDates(1) + 15, "m/d/yy")             'index to middle of month to get the correct month index
    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
    'These fields remain constant in the GRF record
    'tmGrf.iPerGenl(1) = Val(slMonth)
    tmGrf.iPerGenl(0) = Val(slMonth)
    If RptSelPJ!rbcSelCInclude(0).Value Then                    'current wk
        'tmGrf.iPerGenl(2) = 0
        tmGrf.iPerGenl(1) = 0
    Else
        'tmGrf.iPerGenl(2) = 1                                   'past wk
        tmGrf.iPerGenl(1) = 1                                   'past wk
    End If
    gPackDate slCurrDate, tmGrf.iDate(0), tmGrf.iDate(1)        'Date entered (for past) or todays date  (for current)
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.sBktType = "A"                    'assume Actuals (vs difference option)
    If RptSelPJ!ckcSelC3(0).Value = vbChecked Then      'difference only?
        tmGrf.sBktType = "D"
    End If
    If UBound(tlPotn) > 1 Then      '10-17-03 handle case where user doesnt have any potential codes defined
'        tmGrf.iPerGenl(3) = CInt(tlPotn(1).lProject(1))             'Cat A Most Likely
'        tmGrf.iPerGenl(4) = CInt(tlPotn(2).lProject(1))             'Cat B Most likely
'        tmGrf.iPerGenl(5) = CInt(tlPotn(3).lProject(1))             'Cat C Most Likely
'        tmGrf.iPerGenl(6) = CInt(tlPotn(1).lProject(2))             'Cat A Optimistic
'        tmGrf.iPerGenl(7) = CInt(tlPotn(2).lProject(2))             'Cat B Optimistic
'        tmGrf.iPerGenl(8) = CInt(tlPotn(3).lProject(2))             'Cat C Optimistic
'        tmGrf.iPerGenl(9) = CInt(tlPotn(1).lProject(3))             'Cat A Pessimistic
'        tmGrf.iPerGenl(10) = CInt(tlPotn(2).lProject(3))             'Cat B Pessimistic
'        tmGrf.iPerGenl(11) = CInt(tlPotn(3).lProject(3))             'Cat C Pessimistic
        tmGrf.iPerGenl(2) = CInt(tlPotn(1).lProject(1))             'Cat A Most Likely
        tmGrf.iPerGenl(3) = CInt(tlPotn(2).lProject(1))             'Cat B Most likely
        tmGrf.iPerGenl(4) = CInt(tlPotn(3).lProject(1))             'Cat C Most Likely
        tmGrf.iPerGenl(5) = CInt(tlPotn(1).lProject(2))             'Cat A Optimistic
        tmGrf.iPerGenl(6) = CInt(tlPotn(2).lProject(2))             'Cat B Optimistic
        tmGrf.iPerGenl(7) = CInt(tlPotn(3).lProject(2))             'Cat C Optimistic
        tmGrf.iPerGenl(8) = CInt(tlPotn(1).lProject(3))             'Cat A Pessimistic
        tmGrf.iPerGenl(9) = CInt(tlPotn(2).lProject(3))             'Cat B Pessimistic
        tmGrf.iPerGenl(10) = CInt(tlPotn(3).lProject(3))             'Cat C Pessimistic
    End If
    'gather all Slsp projection records for the matching rollover date
    'if current - all zero or blank rollover dates.  If past, get projection rollover date closest to
    'date entered.  In addition, if differnces, do a second pass to the projection records and retrieve
    'the previous weeks projection records
    ilMaxOption = 1
    If RptSelPJ!ckcSelC3(0).Value = vbChecked And RptSelPJ!rbcSelCInclude(1).Value Then        'if both Past & differences only, do a 2nd pass thru Projections
                                                                                'to retrieve previous weeks data for comparisons
        ilMaxOption = 2
    End If
    For ilLoopProj = 1 To ilMaxOption                 '1st pass is for the current week (or dated week),
                                        '2nd pass is for the previous week for differences option
        ReDim tmTPjf(0 To 0) As PJF

        ilRet = gObtainPjf(RptSelPJ, hmPjf, ilRODate(), tmTPjf())                 'Read all applicable Projection records into memory
        For ilLoop = LBound(tmTPjf) To UBound(tmTPjf) Step 1
            If ilListIndex = PRJ_SALESPERSON And Not RptSelPJ!ckcAll.Value = vbChecked Then
                'setup selective salespeople
                ilFound = False
                For ilTemp = 0 To RptSelPJ!lbcSelection(2).ListCount - 1 Step 1
                    If RptSelPJ!lbcSelection(2).Selected(ilTemp) Then
                        slNameCode = tgSalesperson(ilTemp).sKey    'Traffic!lbcSalesperson.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                        If Val(slCode) = tmTPjf(ilLoop).iSlfCode Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilTemp
            ElseIf ilListIndex = PRJ_VEHICLE And Not RptSelPJ!ckcAll.Value = vbChecked Then
                'setup selective vehicles
                For ilTemp = 0 To RptSelPJ!lbcSelection(6).ListCount - 1 Step 1
                    If RptSelPJ!lbcSelection(6).Selected(ilTemp) Then
                        slNameCode = tgCSVNameCode(ilTemp).sKey    'rptselpj!lbcCSVNameCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                        If Val(slCode) = tmTPjf(ilLoop).iVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilTemp
            Else
                ilFound = True                  'all other options, use the projction recd with filtering
            End If      'if not ckall
            If ilFound Then
                'Determine start & end dates of the standard year from the proj recd
                slStr = "12/15/" & Trim$(str$(tmTPjf(ilLoop).iYear))
                llProjYearEnd = gDateValue(gObtainEndStd(slStr))
                slStr = "1/15/" & Trim$(str$(tmTPjf(ilLoop).iYear))
                llProjYearStart = gDateValue(gObtainStartStd(slStr))
                'For ilMonth = 1 To 6
                For ilMonth = 0 To 5
                    tmGrf.lDollars(ilMonth) = 0
                Next ilMonth
                llTotalGross = 0
                For ilLoopWks = 1 To 53
                    'llProjYearStart is the first weeks to test for validity
                    For ilMonth = 1 To 6            'loop for # months in  qtr
                        llAdjust = 0
                        If llProjYearStart >= lmStartDates(1) Then
                            If llProjYearStart >= lmStartDates(ilMonth) And llProjYearStart < lmEndDates(ilMonth) Then
                                llAdjust = llAdjust + tmTPjf(ilLoop).lGross(ilLoopWks)
                                If ilLoopProj = 2 Then                      '2nd pass, differences (negate these to get diff)
                                    llAdjust = -llAdjust
                                End If
                                tmGrf.lDollars(ilMonth - 1) = tmGrf.lDollars(ilMonth - 1) + llAdjust
                                llTotalGross = llTotalGross + llAdjust                  'accum the total of all months, dont write to disk if zero
                                Exit For
                            End If
                        End If
                    Next ilMonth
                    llProjYearStart = llProjYearStart + 7       'increment next week
                Next ilLoopWks
                'Format remainders of fields required, then Write GRF record to disk
                'Gen date and time and current date entered are updated at the beginning of code
                tmGrf.iVefCode = tmTPjf(ilLoop).iVefCode
                tmGrf.iSlfCode = tmTPjf(ilLoop).iSlfCode
                tmGrf.iSofCode = tmTPjf(ilLoop).iSofCode
                tmGrf.iAdfCode = tmTPjf(ilLoop).iAdfCode
                tmGrf.lCode4 = tmTPjf(ilLoop).lCxfChgR          'change reason
                tmGrf.iCode2 = tmTPjf(ilLoop).iMnfBus           'potential code
                tmGrf.sDateType = Trim$(slBaseDate)              'C = corp, S = std
                tmGrf.iYear = tmTPjf(ilLoop).iYear
                tmGrf.lChfCode = tmTPjf(ilLoop).lChfCode
                If llTotalGross <> 0 Then
                    'get the product name and built into record (cannot place prf code into
                    'grf; not enuf long integers
                    tmSrchKey.lCode = tmTPjf(ilLoop).lPrfCode
                    ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmPrf.sName = ""
                    End If
                    tmGrf.sGenDesc = tmPrf.sName
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If
            End If                              'ilfound
        Next ilLoop
        If ilLoopProj = 1 And ilMaxOption = 2 Then              'just finished 1st pass, should there be a 2nd pass for differences?
            'setup processing for previous weeks projections
            slStr = gDecOneWeek(slCurrDate)                     'backup the Current date by one week
            llClosestDate = 0           '1-3-02 init the date for the next pass
            gGetRollOverDate RptSelPJ, 2, slStr, llClosestDate   'get closest rollover date one week ago
            slStr = Format$(llClosestDate, "m/d/yy")
            gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
            gPackDate slStr, ilRODate(0), ilRODate(1)
        End If
    Next ilLoopProj
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPrf)
    btrDestroy hmSlf
    btrDestroy hmMnf
    btrDestroy hmPjf
    btrDestroy hmGrf
    btrDestroy hmPrf
    Erase tlPotn
    Erase tlMMnf
    Erase tmTPjf
End Sub
