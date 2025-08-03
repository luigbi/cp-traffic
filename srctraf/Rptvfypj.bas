Attribute VB_Name = "RPTVFYPJ"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfypj.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelPj.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelPjAgencyCode() As SORTCODE
'Public sgRptSelPjAgencyCodeTag As String
'Public tgRptSelPjSalespersonCode() As SORTCODE
'Public sgRptSelPjSalespersonCodeTag As String
'Public tgRptSelPjAdvertiserCode() As SORTCODE
'Public sgRptSelPjAdvertiserCodeTag As String
'Public tgRptSelPjNameCode() As SORTCODE
'Public sgRptSelPjNameCodeTag As String
'Public tgRptSelPjBudgetCode() As SORTCODE
'Public sgRptSelPjBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelPjDemoCode() As SORTCODE
'Public sgRptSelPjDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
''Rate Card job report constants
''Global Const RC_RCITEMS = 0             'Rate carditems
''Global Const RC_DAYPARTS = 1            'Dayparts
''Sales commissions job report constants
''Global Const COMM_SALESCOMM = 0         'sales commission
''Global Const COMM_PROJECTION = 1        'projection report
''Projections job report constants
'Public Const PRJ_SALESPERSON = 0
'Public Const PRJ_VEHICLE = 1
'Public Const PRJ_OFFICE = 2
'Public Const PRJ_CATEGORY = 3
''Global Const PRJ_SCENARIO = 4
'Public Const PRJ_POTENTIAL = 4
''Orders, Proposals, and Spots jobs report constants
''Global Const CNT_BR = 0                     'BRoadcast contracts (proposal, narrow & wide)
''Global Const CNT_PAPERWORK = 1              'Paperwork, summary
''Global Const CNT_SPTSBYADVT = 2             'Spots by Advt
''Global Const CNT_SPTSBYDATETIME = 3         'Spots by Date and Time
''Global Const CNT_BOB_BYCNT = 4              'Business Booked (projection)
''Global Const CNT_RECAP = 5                  'Recap
''Global Const CNT_PLACEMENT = 6              'Spot Placement
''Global Const CNT_DISCREP = 7                'discrepancy
''Global Const CNT_MG = 8                     'makegoods (MG)
''Global Const CNT_SPOTTRAK = 9               'Spot Tracking
''Global Const CNT_COMLCHG = 10               'Commercial Change
''Global Const CNT_HISTORY = 11               'History
''Global Const CNT_AFFILTRAK = 12             'Affiliate Tracking
''Global Const CNT_SPOTSALES = 13             'Spot Sales
''Global Const CNT_MISSED = 14                'Missed
''Global Const CNT_BOB_BYSPOT = 15            'business Booked by Spots (Spot Projection)
''Global Const CNT_BOB_BYSPOT_REPRINT = 16    'Business Booked Reprint (Projection reprint)
''Global Const CNT_QTRLY_AVAILS = 17          'Quarterly Avails
''Global Const CNT_AVG_PRICES = 18            'Weekly Average Prices
''Global Const CNT_ADVT_UNITS = 19            'Advt Units Ordered
''Global Const CNT_SALES_CPPCPM = 20          'Sales Analysis by CPP CPM
''Global Const CNT_AVGRATE = 21               'Average Rate
''Global Const CNT_TIEOUT = 22                'Tie Out
''Global Const CNT_BOB = 23                   'Billed & Booked Report
''Global Const CNT_SALESACTIVITY = 24         'Sales Activity
''Invoice report options
''Global Const INV_REGISTER = 0               'Invoice Registers (by inv #, advt, slsp, vehicle)
''Global Const INV_VIEWEXPORT = 1             'View Export
''Global Const INV_DISTRIBUTE = 2             'Billing distribution
''Collections report options
''Global Const COLL_CASH = 0                  'Cash receipts
''Global Const COLL_AGEPAYEE = 1              'Ageing by Payee
''Global Const COLL_AGESLSP = 2               'Ageing by Salesperson
''Global Const COLL_AGEVEHICLE = 3            'Ageing by Vehicle
''Global Const COLL_DELINQUENT = 4            'Delinquent report
''Global Const COLL_STATEMENT = 5             'Statments
''Global Const COLL_PAYHISTORY = 6            'Payment History
''Global Const COLL_CREDITSTATUS = 7          'Credit Status
''Global Const COLL_DISTRIBUTE = 8            'Cash Distribution
''Global Const COLL_CASHSUM = 9               'Cash summary
''Global Const COLL_ACCTHIST = 10             'Account History
''Global Const COLL_MERCHANT = 11             'Merchandising & Promotions
''Global Const COLL_MERCHRECAP = 12            'Merchandising & Promotions Recap
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
''Global spot types for Spots by Advt & spots by Date & Time
''bit selectivity for charged and different types of no charge spots
''bits defined right to left (0 to 9)
''Global Const SPOT_CHARGE = &H1         'charged
''Global Const SPOT_00 = &H2          '0.00
''Global Const SPOT_ADU = &H4         'ADU
''Global Const SPOT_BONUS = &H8       'bonus
''Global Const SPOT_EXTRA = &H10      'Extra
''Global Const SPOT_FILL = &H20       'Fill
''Global Const SPOT_NC = &H40         'no charge
''Global Const SPOT_MG = &H80         'mg
''Global Const SPOT_RECAP = &H100     'recapturable
''Global Const SPOT_SPINOFF = &H200   'spinoff
'Library calendar file- used to obtain post log date status
'  Slsp Projection
Dim lmStartDates() As Long          'array of 13 bdcst or corp start dates
Dim lmEndDates() As Long            'array of 13 bdcst or corp end dates
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmcGenPjct                      *
'*                                                     *
'*             Created:7/24/97       By:W. Bjerke      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Duplicate of gCmcGen due to       *
'*                   expanding code base.              *
'*                   Used for projections              *
'*******************************************************
Function gCmcGenPjct(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
    Dim ilLoop As Integer
    Dim slSelection As String
    Dim ilresult As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slOr As String
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slStatus As String
    Dim slTime As String
    Dim slBaseDate As String
    Dim slMoreStr As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim slMnfRPU As String
    Dim slMnfRPU2 As String
    Dim slMnfRPU3 As String
    Dim slMnfUnitType As String
    Dim slMnfUnitType1 As String
    Dim slMnfUnitType2 As String
    Dim slMnfSSComm As String
    Dim slMnfSSComm2 As String
    Dim slMnfSSComm3 As String
    Dim ilRecLen As Integer
    Dim llNoRec As Long
    Dim ilLgth As Integer
    Dim ilSlfCode As Integer
    'ReDim lmStartDates(1 To 13) As Long
    'ReDim lmEndDates(1 To 13) As Long
    ReDim lmStartDates(0 To 13) As Long 'Index zero ignored
    ReDim lmEndDates(0 To 13) As Long   'Index zero ignored
    Dim llClosestDate As Long               'projection extraction should match this date
gCmcGenPjct = 0
ilSlfCode = tgUrf(0).iSlfCode               'determine if this a slsp requesting the projection
Select Case igRptCallType
    Case PROPOSALPROJECTION
        slSelection = ""                    'initialize for record selections to pass to Crystal
        slDate = RptSelPJ!edcSelCFrom.Text
        If gValidDate(slDate) Or (slDate = "" And RptSelPJ!rbcSelCInclude(0)) Then        'must be valid date entered or
                                                                                        'if not date must be requesting current
            gCurrDateTime slStr, slTime, slMonth, slDay, slYear
            'change 7/13/98 to use NOW formula to get time in all Crystal reports
            'If Not gSetFormula("AsOfT", Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))) Then
            '    gCmcGenPjct = -1
            '    Exit Function
            'End If
            'If illistindex = PRJ_SALESPERSON Or illistindex = PRJ_VEHICLE Or illistindex = PRJ_OFFICE Or illistindex = PRJ_CATEGORY Or illistindex = PRJ_POTENTIAL Then
            If ilListIndex = PRJ_POTENTIAL Then

                'all options (salesp, office, vehicle), send if actual or differences
                If RptSelPJ!ckcSelC3(0).Value = vbChecked Then              'diff only
                    slMoreStr = "D"
                Else
                    slMoreStr = "A"
                End If
                If RptSelPJ!edcSelCFrom.Text = "" Then
                    slStatus = "C"                            'no report date entered, get current
                Else
                    slStatus = "P"                            'get past , must match rollover dates
                End If
                slMoreStr = slMoreStr & slStatus
                'concatenate Actual or Diff flag with Current/Past flag
                If Not gSetFormula("ActualOrDifF", "'" & slMoreStr & "'") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
                'send formulas for start dates of the 6 months to report
                If RptSelPJ!rbcSelCSelect(0).Value Then          'use corp qtr or std qtr
                    slBaseDate = "C"
                Else
                    slBaseDate = "S"
                End If

                'if no date entered, get current date so that we can determine what months to produce output for
                If RptSelPJ!edcSelCFrom.Text = "" Then
                    slDate = slStr                          'default to todays date
                End If
                gCorpStdDates slBaseDate, slDate, ilresult, lmStartDates(), lmEndDates()
                slStr = Format$(lmStartDates(1) - 1, "m/d/yy")

                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If

                For ilLoop = 1 To 6
                    slStr = Format$(lmEndDates(ilLoop), "m/d/yy")
                    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(ilLoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenPjct = -1
                        Exit Function
                    End If
                    If ilLoop = 1 Then              'pass the start month of data gathering
                        'get to the middle of the month to insure the correct month # is passed
                        slStr = Format$(lmEndDates(1) - 15, "m/d/yy")
                        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                        If Not gSetFormula("StartMonth", Val(slMonth)) Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    End If
                Next ilLoop
                'End If
                slMoreStr = Format$(lmEndDates(3) + 15, "m/d/yy") 'get middle of start of 2nd quarter to put to determine year
                gObtainYearMonthDayStr slMoreStr, True, slYear, slMonth, slDay
                slMoreStr = slYear          'save the year of the 2nd qtr to be processed
                                            'If Quarter 2 is different year, need to calc the
                                            'start date of the year and send to Crystal

                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("EffDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
                slMonth = "01"          'set to calc std or corp start date of year
                slDay = "15"
                'slYear set above for year of effective date for Quarter 1
                'If RptSelPj!rbcSelCSelect(0).Value Then         'corporate month?  (vs std)
                '    slStr = gObtainStartCorp(slMonth & "/" & slDay & "/" & slYear, True)
                '    slStr = Format$(gDateValue(slStr), "m/d/yy")    'Start of the corporate year
                'Else
                    slStr = gObtainStartStd(slMonth & "/" & slDay & "/" & slYear)
                    slStr = Format$(gDateValue(slStr), "m/d/yy")    'start of stdyear
                'End If
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                'report needs the start date of year so that Crystal knows the correct week in a period
                If Not gSetFormula("StartOfYear", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
                slMonth = "01"          'set to calc std or corp start date of year
                slDay = "15"
                'slMoreStr = year of last month to process, could be going into 1st qutr of next year
                'If RptSelPj!rbcSelCSelect(0).Value Then         'corporate month?  (vs std)
                '    slStr = gObtainStartCorp(slMonth & "/" & slDay & "/" & slMoreStr, True)
                '    slStr = Format$(gDateValue(slStr), "m/d/yy")    'Start of the corporate year
                'Else
                    slStr = gObtainStartStd(slMonth & "/" & slDay & "/" & slMoreStr)
                    slStr = Format$(gDateValue(slStr), "m/d/yy")    'start of stdyear
                'End If
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                'report needs the start date of year so that Crystal knows the correct week in a period
                If Not gSetFormula("StartOfYearQ2", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
                'pass year end dates of the 1st and 2nd quarters  (2nd & 5th qtr will always be within the proper year to obtain the year end of each qtr
                If ilListIndex <> 4 Then    'exclude PRJ_POTENTIAL
                    slStr = Format$(lmEndDates(2), "m/d/yy")
                End If
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                'If RptSelPj!rbcSelCSelect(0).Value Then
                '    slStr = gObtainEndCorp("12" & "/" & "15" & "/" & slYear, True)
                '    slStr = Format$(gDateValue(slStr), "m/d/yy")    'Start of the corporate year
                'Else
                    slStr = gObtainEndStd("12" & "/" & "15" & "/" & slYear)
                    slStr = Format$(gDateValue(slStr), "m/d/yy")    'start of stdyear
                'End If
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                If Not gSetFormula("EndYearWk", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If

                'pass year end dates of the 1st and 2nd quarters
                slStr = Format$(lmEndDates(5), "m/d/yy")
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                'If RptSelPj!rbcSelCSelect(0).Value Then
                '    slStr = gObtainEndCorp("12" & "/" & "15" & "/" & slYear, True)
                '    slStr = Format$(gDateValue(slStr), "m/d/yy")    'Start of the corporate year
                'Else
                    slStr = gObtainEndStd("12" & "/" & "15" & "/" & slYear)
                    slStr = Format$(gDateValue(slStr), "m/d/yy")    'start of stdyear
                'End If
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                If Not gSetFormula("EndYearWkQ2", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
                If RptSelPJ!edcSelCFrom.Text = "" Then                    'no dates entered, select with 0 rollover
                    slSelection = "{PJF_Projections.pjfRollOverDate} = Date(0,0,0)"
                Else
                    gGetRollOverDate RptSelPJ, 2, slDate, llClosestDate   'send the lbcselection index to search, plust rollover date
                    slDate = Format$(llClosestDate, "m/d/yy")
                    'send date to match on in the projection file
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slSelection = "{PJF_Projections.pjfRollOverDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If
                If RptSelPJ!ckcSelC3(0).Value = vbChecked Then         'differences report, also retrieve previous weeks stuff
                    slDate = gDecOneWeek(slDate)
                    gGetRollOverDate RptSelPJ, 2, slDate, llClosestDate   'get closest rollover date one week ago
                    slDate = Format$(llClosestDate, "m/d/yy")
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay

                    'Date totest for differences option
                    If Not gSetFormula("DiffDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenPjct = -1
                        Exit Function
                    End If
                    slSelection = "(" & slSelection & " or " & "{PJF_Projections.pjfRollOverDate} = Date(" & slYear & "," & slMonth & "," & slDay & "))"
                End If
                If Not gSetFormula("CorpStd", "'" & slBaseDate & "'") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
            End If
            If ilListIndex <> PRJ_POTENTIAL Then
                If Not RptSelPJ!ckcAll.Value = vbChecked Then         'not all vehicles selected
                    If slSelection <> "" Then
                        slSelection = "(" & slSelection & ") " & " and ("
                        slOr = ""
                    Else
                        slSelection = "("
                        slOr = ""
                    End If
                    'If illistindex = PRJ_SALESPERSON Then
                        'setup selective salespeople
                    '    For ilLoop = 0 To RptSelPj!lbcSelection(2).ListCount - 1 Step 1
                    '        If RptSelPj!lbcSelection(2).Selected(ilLoop) Then
                    '            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
                    '            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    '            slSelection = slSelection & slOr & "{PJF_Projections.pjfslfCode} = " & Trim$(slCode)
                    '            slOr = " Or "
                    '        End If
                    '    Next ilLoop
                    'ElseIf illistindex = PRJ_VEHICLE Then
                    If ilListIndex = PRJ_VEHICLE Then
                        'setup selective vehicles
                        For ilLoop = 0 To RptSelPJ!lbcSelection(6).ListCount - 1 Step 1
                            If RptSelPJ!lbcSelection(6).Selected(ilLoop) Then
                                slNameCode = tgCSVNameCode(ilLoop).sKey    'rptselpj!lbcCSVNameCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                slSelection = slSelection & slOr & "{PJF_Projections.pjfvefCode} = " & Trim$(slCode)
                                slOr = " Or "
                            End If
                        Next ilLoop
                    End If                  'ILLISTINDEX = VEHICLE
                    slSelection = slSelection & ")"
                End If      'if not ckall
            End If
            If ilListIndex = PRJ_VEHICLE Then               'send to crystal which sort should be output
                If Not gSetFormula("SortOption", "'V'") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
            ElseIf ilListIndex = PRJ_OFFICE Then
                If Not gSetFormula("SortOption", "'O'") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
            ElseIf ilListIndex = PRJ_CATEGORY Then
                If Not gSetFormula("SortOption", "'C'") Then
                    gCmcGenPjct = -1
                    Exit Function
                End If
            ElseIf ilListIndex = PRJ_POTENTIAL Then
                'pass potential codes from MNF file to report column headings
                hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
                ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        btrDestroy hlMnf
                        Exit Function
                    End If

                ilRecLen = Len(tlMnf)
                llNoRec = gExtNoRec(ilRecLen)
                btrExtClear hlMnf
                ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'get values for name = A
                    Do Until Trim$(tlMnf.sName) = "A"
                        ilRet = btrGetNext(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE)
                    Loop

                    slMnfRPU = tlMnf.sRPU
                    gPDNToStr tlMnf.sRPU, 2, slMnfRPU2
                    slMnfRPU3 = Mid$(slMnfRPU2, 1, InStr(1, slMnfRPU2, Chr$(46)) - 1)

                        If Not gSetFormula("AOp", "'" & slMnfRPU3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfUnitType = tlMnf.sUnitType
                    ilLgth = InStr(1, slMnfUnitType, Chr$(46)) - 1
                    If ilLgth < 0 Then
                        slMnfUnitType2 = Mid$(slMnfUnitType, 1, 2)
                    Else
                        slMnfUnitType2 = Mid$(slMnfUnitType, 1, InStr(1, slMnfUnitType, Chr$(46)) - 1)
                    End If
                        If Not gSetFormula("AMl", "'" & slMnfUnitType2 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfSSComm = tlMnf.sSSComm
                    gPDNToStr slMnfSSComm, 4, slMnfSSComm2
                    slMnfSSComm3 = Mid$(slMnfSSComm2, 1, InStr(1, slMnfSSComm2, Chr$(46)) - 1)
                        If Not gSetFormula("APs", "'" & slMnfSSComm3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If

                    'get values for name = B
                    ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do Until Trim$(tlMnf.sName) = "B"
                        ilRet = btrGetNext(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE)
                    Loop

                    slMnfRPU = tlMnf.sRPU
                    gPDNToStr tlMnf.sRPU, 2, slMnfRPU2
                    slMnfRPU3 = Mid$(slMnfRPU2, 1, InStr(1, slMnfRPU2, Chr$(46)) - 1)
                        If Not gSetFormula("BOp", "'" & slMnfRPU3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfUnitType = tlMnf.sUnitType
                    If ilLgth < 0 Then
                        slMnfUnitType1 = Mid$(slMnfUnitType, 1, 2)
                    Else
                        slMnfUnitType1 = Mid$(slMnfUnitType, 1, InStr(1, slMnfUnitType, Chr$(46)) - 1)
                    End If
                        If Not gSetFormula("BMl", "'" & slMnfUnitType1 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfSSComm = tlMnf.sSSComm
                    gPDNToStr slMnfSSComm, 4, slMnfSSComm2
                    slMnfSSComm3 = Mid$(slMnfSSComm2, 1, InStr(1, slMnfSSComm2, Chr$(46)) - 1)
                    If Not gSetFormula("BPs", "'" & slMnfSSComm3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If

                    'get values for name = C
                    ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do Until Trim$(tlMnf.sName) = "C"
                        ilRet = btrGetNext(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE)
                    Loop

                    slMnfRPU = tlMnf.sRPU
                    gPDNToStr slMnfRPU, 2, slMnfRPU2
                    slMnfRPU3 = Mid$(slMnfRPU2, 1, InStr(1, slMnfRPU2, Chr$(46)) - 1)
                        If Not gSetFormula("COp", "'" & slMnfRPU3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfUnitType = tlMnf.sUnitType
                    If ilLgth < 0 Then
                        slMnfUnitType1 = Mid$(slMnfUnitType, 1, 2)
                    Else
                        slMnfUnitType1 = Mid$(slMnfUnitType, 1, InStr(1, slMnfUnitType, Chr$(46)) - 1)
                    End If
                        If Not gSetFormula("CMl", "'" & slMnfUnitType1 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If
                    slMnfSSComm = tlMnf.sSSComm
                    gPDNToStr slMnfSSComm, 4, slMnfSSComm2
                    slMnfSSComm3 = Mid$(slMnfSSComm2, 1, InStr(1, slMnfSSComm2, Chr$(46)) - 1)
                        If Not gSetFormula("CPs", "'" & slMnfSSComm3 & "'") Then
                            gCmcGenPjct = -1
                            Exit Function
                        End If

                End If
            End If

        Else
            mReset
            RptSelPJ!edcSelCTo.SetFocus
            Exit Function
        End If
        If ilSlfCode > 0 Then               'its a slsp requesting this report, only
                                            'allow slsp to get their own stuff
            slSelection = slSelection & " and ({PJF_Projections.pjfslfCode} = " & Trim$(str$(ilSlfCode)) & ")"
        End If

        If ilListIndex = PRJ_SALESPERSON Or ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
            'If slSelection <> "" Then
            '    slSelection = "(" & slSelection & ") and " & "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            '    slSelection = slSelection & " And {GRF_Generic_Report.grfGenTime} = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
            'Else
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            'End If
        End If
        If Not gSetSelection(slSelection) Then
            gCmcGenPjct = -1
            Exit Function
        End If
    End Select
gCmcGenPjct = 1
Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportPj                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:8/14/97       By:W. Bjerke      *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportPj() As Integer
    Dim ilListIndex As Integer
    ilListIndex = RptSelPJ!lbcRptType.ListIndex

    If ilListIndex = PRJ_SALESPERSON Then
        If Not gOpenPrtJob("pjsls.Rpt") Then
            gGenReportPj = False
            Exit Function
        End If
    ElseIf ilListIndex = PRJ_VEHICLE Or ilListIndex = PRJ_OFFICE Or ilListIndex = PRJ_CATEGORY Then
        If Not gOpenPrtJob("pjvehofc.Rpt") Then
            gGenReportPj = False
            Exit Function
        End If
    ElseIf ilListIndex = PRJ_POTENTIAL Then
        If Not gOpenPrtJob("pjpot.Rpt") Then
            gGenReportPj = False
            Exit Function
        End If
    End If



    gGenReportPj = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReset                   *
'*                                                     *
'*             Created:1/31/96       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reset controls                 *
'*                                                     *
'*******************************************************
Sub mReset()
    igGenRpt = False
    RptSelPJ!frcOutput.Enabled = igOutput
    RptSelPJ!frcCopies.Enabled = igCopies
    'RptSelPj!frcWhen.Enabled = igWhen
    RptSelPJ!frcFile.Enabled = igFile
    RptSelPJ!frcOption.Enabled = igOption
    'RptSelPj!frcRptType.Enabled = igReportType
    Beep
End Sub
