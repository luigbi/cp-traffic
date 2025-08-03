Attribute VB_Name = "RPTVFYCT"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyct.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSel.Bas
'
' Release: 1.0
'
' Description:Favggr
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

'Public sgVehicleSetsTag As String
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelAgencyCode() As SORTCODE
'Public sgRptSelAgencyCodeTag As String
'Public tgRptSelSalespersonCode() As SORTCODE
'Public sgRptSelSalespersonCodeTag As String
'Public tgRptSelAdvertiserCode() As SORTCODE
'Public sgRptSelAdvertiserCodeTag As String
'Public tgRptSelNameCode() As SORTCODE
'Public sgRptSelNameCodeTag As String
'Public tgRptSelBudgetCode() As SORTCODE
'Public sgRptSelBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelDemoCode() As SORTCODE
'Public sgRptSelDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public tgBookName() As SORTCODE
'Public sgBookNameTag As String
'Public tgMnfCode() As SORTCODE
'Public sgMnfCodeTag As String
''Rate Card job report constants
'Public Const RC_RCITEMS = 0             'Rate carditems
'Public Const RC_DAYPARTS = 1            'Dayparts
''Sales commissions job report constants
'Public Const COMM_SALESCOMM = 0         'sales commission
'Public Const COMM_PROJECTION = 1        'projection report
''Projections job report constants
'Public Const PRJ_SALESPERSON = 0
'Public Const PRJ_VEHICLE = 1
'Public Const PRJ_OFFICE = 2
'Public Const PRJ_CATEGORY = 3
'Public Const PRJ_SCENARIO = 4
''Orders, Proposals, and Spots jobs report constants
'Public Const CNT_BR = 0                     'BRoadcast contracts (proposal, narrow & wide)
'Public Const CNT_PAPERWORK = 1              'Paperwork, summary
'Public Const CNT_SPTSBYADVT = 2             'Spots by Advt
'Public Const CNT_SPTSBYDATETIME = 3         'Spots by Date and Time
'Public Const CNT_BOB_BYCNT = 4              'Business Booked (projection)
'Public Const CNT_RECAP = 5                  'Recap
'Public Const CNT_PLACEMENT = 6              'Spot Placement
'Public Const CNT_DISCREP = 7                'discrepancy
'Public Const CNT_MG = 8                     'makegoods (MG)
'Public Const CNT_SPOTTRAK = 9               'Spot Tracking
'Public Const CNT_COMLCHG = 10               'Commercial Change
'Public Const CNT_HISTORY = 11               'History
'Public Const CNT_AFFILTRAK = 12             'Affiliate Tracking
'Public Const CNT_SPOTSALES = 13             'Spot Sales
'Public Const CNT_MISSED = 14                'Missed
'Public Const CNT_BOB_BYSPOT = 15            'business Booked by Spots (Spot Projection)
'Public Const CNT_BOB_BYSPOT_REPRINT = 16    'Business Booked Reprint (Projection reprint)
'Public Const CNT_QTRLY_AVAILS = 17          'Quarterly Avails
'Public Const CNT_AVG_PRICES = 18            'Weekly Average Prices
'Public Const CNT_ADVT_UNITS = 19            'Advt Units Ordered
'Public Const CNT_SALES_CPPCPM = 20          'Sales Analysis by CPP CPM
'Public Const CNT_AVGRATE = 21               'Average Rate
'Public Const CNT_TIEOUT = 22                'Tie Out
'Public Const CNT_BOB = 23                   'Billed & Booked Report
'Public Const CNT_SALESACTIVITY = 24         'Sales Activity
'Public Const CNT_SALESCOMPARE = 25          'Sales Comparison
'Public Const CNT_CUMEACTIVITY = 26          'Cumulative Activity
'Public Const CNT_MAKEPLAN = 27              'Avg prices needed to make plan
'Public Const CNT_VEHCPPCPM = 28             'Current CPP & CPM by vehicle
'Public Const CNT_SALESANALYSIS = 29         'Sales Analysis Summary
'Public Const CNT_INSERTION = 30             'Insertion Orders
'Public Const CNT_DAILY_SALESACTIVITY = 31   'Daily Sales Activity by Contract 6-5-01
''Invoice report options
'Public Const INV_REGISTER = 0               'Invoice Registers (by inv #, advt, slsp, vehicle)
'Public Const INV_VIEWEXPORT = 1             'View Export
'Public Const INV_DISTRIBUTE = 2             'Billing distribution
''Collections report options
'Public Const COLL_CASH = 0                  'Cash receipts
'Public Const COLL_AGEPAYEE = 1              'Ageing by Payee
'Public Const COLL_AGESLSP = 2               'Ageing by Salesperson
'Public Const COLL_AGEVEHICLE = 3            'Ageing by Vehicle
'Public Const COLL_DELINQUENT = 4            'Delinquent report
'Public Const COLL_STATEMENT = 5             'Statments
'Public Const COLL_PAYHISTORY = 6            'Payment History
'Public Const COLL_CREDITSTATUS = 7          'Credit Status
'Public Const COLL_DISTRIBUTE = 8            'Cash Distribution
'Public Const COLL_CASHSUM = 9               'Cash summary
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public igRCSelectedIndex As Integer         'selected r/c index
'Public igBSelectedIndex As Integer          'selected budget index
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
'Public Const SPOT_CHARGE = &H1         'charged
'Public Const SPOT_00 = &H2          '0.00
'Public Const SPOT_ADU = &H4         'ADU
'Public Const SPOT_BONUS = &H8       'bonus
'Public Const SPOT_EXTRA = &H10      'Extra
'Public Const SPOT_FILL = &H20       'Fill
'Public Const SPOT_NC = &H40         'no charge
'Public Const SPOT_MG = &H80         'mg
'Public Const SPOT_RECAP = &H100     'recapturable
'Public Const SPOT_SPINOFF = &H200   'spinoff

'Dim imBrSumZer As Integer       '1-28-10 need to know if its monthly summary to bring in correct filters for crystalDim hmChf As Integer            'Contract header file handle
'Dim imBRSum As Integer          '2-12-10 need to know its a research summary version to show NTR vehicle totals with the air time Research
Dim hmCHF As Integer
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim tmChf As CHF

'Library calendar file- used to obtain post log date status
'*******************************************************
'*                                                     *
'*      Procedure Name:mCntJob11Plus                   *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Contract reports    *
'*            Originally mContractJob which handled    *
'*            all formulas and selection for Contract. *
'*            It has been split into two modules con-  *
'*            sisting of ilListIndex 1-10 into and     *
'*            ilListIndex 11 Plus                      *
'*******************************************************
Function mCntJob11Plus(ilListIndex As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilAllSelected                                                                         *
'******************************************************************************************

    Dim slDate As String
    Dim slTime As String
    Dim slYear As String
    Dim slYear2 As String
    Dim slMonth As String
    Dim slMonth2 As String
    Dim slDay As String
    Dim slDay2 As String
    Dim slSelection As String
    Dim illoop As Integer
    Dim slStr As String
    Dim slOr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilMonth As Integer
    Dim slType As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slReserve As String             'show or hide reservations for qtrly detail
    Dim llDate As Long
    Dim llDate2 As Long
    Dim slEarliest As String
    Dim slLatest As String
    Dim ilSaveMonth As Integer
    Dim slSaveYear As String
    Dim ilTemp As Integer
    Dim ilErr As Integer
    Dim ilVehicleGroup As Integer
    Dim slMonthHdr As String * 36
    Dim ilNoneExists As Integer
    Dim ilMinorGroupHdr As Integer
    Dim ilDate(0 To 1) As Integer
    Dim ilPeriods As Integer
    Dim slPacingTY As String  '2-19-16
    Dim slPacingLY As String  '2-19-16
    Dim ilYear As Integer
    Dim llPacingDate As Long

    mCntJob11Plus = 0

    If (ilListIndex = CNT_HISTORY) Then 'Contract History
        'Date: 12/17/2019 added CSI calendar control for date entry
        'If (RptSelCt!edcSelCFrom.Text <> "") And (RptSelCt!edcSelCTo.Text <> "") Then
        If (RptSelCt!CSI_CalFrom.Text <> "") And (RptSelCt!CSI_CalTo.Text <> "") Then
            If StrComp(RptSelCt!edcSelCTo.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_CalFrom.Text      ' edcSelCFrom.Text
                If gValidDate(slDate) Then
                    slDate = RptSelCt!CSI_CalTo.Text    ' edcSelCTo.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCt!CSI_CalTo.SetFocus     ' edcSelCTo.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus       ' edcSelCFrom.SetFocus
                    Exit Function
                End If
            Else
                slDate = RptSelCt!CSI_CalFrom.Text      ' edcSelCFrom.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus       ' edcSelCFrom.SetFocus
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!CSI_CalFrom.Text <> "" Then     ' edcSelCFrom.Text <> "" Then
            slDate = RptSelCt!CSI_CalFrom.Text          ' edcSelCFrom.Text
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalFrom.SetFocus           ' edcSelCFrom.SetFocus
                Exit Function
            End If
        ElseIf RptSelCt!CSI_CalTo.Text <> "" Then       ' edcSelCTo.Text <> "" Then
            'If StrComp(RptSelCt!edcSelCTo.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_CalTo.Text        ' edcSelCTo.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_CalTo.SetFocus         ' edcSelCTo.SetFocus
                    Exit Function
                End If
            End If
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_AFFILTRAK Then 'Affiliate Spot Tracking, converted to crystal 10-25-00
        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_CalFrom.SetFocus   'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_CalTo.Text        'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalTo.SetFocus     'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_From1.Text        'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCTo.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_From1.SetFocus     'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_To1.Text          'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCTo1.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_To1.SetFocus       'Date: 12/6/2019 added CSI calendar control for date entries --> edcSelCTo1.SetFocus
                Exit Function
            End If
        End If
        ilErr = gGRFSelection(slSelection)      'build date & time filter to send to crystal
        If ilErr <> 0 Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        ilRet = mTrackAndComlChgDates(ilListIndex)     'send dates requested for report headings
        If ilRet <> 0 Then
            mCntJob11Plus = -1
            Exit Function
            End If
    'Missed spot code removed--see rptselcb
    'Spot Sales code removed--see rptselcb
    ElseIf (ilListIndex = CNT_BOB_BYSPOT) Or (ilListIndex = CNT_BOB_BYSPOT_REPRINT) Then    'Spot projection
        If ilListIndex = CNT_BOB_BYSPOT Then
            slStr = RptSelCt!edcSelCTo.Text            '# columns to print
            If Not mVerifyNumber(slStr) Then
                mReset
                RptSelCt!edcSelCTo.SetFocus
                Exit Function
            End If
            ilTemp = Val(RptSelCt!edcSelCTo.Text)
            If ilTemp > 13 Then
                mReset
                RptSelCt!edcSelCTo.SetFocus
                Exit Function
            End If
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{JSR_Spot_Projection.jsrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({JSR_Spot_Projection.jsrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        End If

        'send crystal whether its gross or net
        If RptSelCt!rbcSelC4(0).Value Then          'gross
            If Not gSetFormula("GrossOrNet", "'Gross'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                        'net
            If Not gSetFormula("GrossOrNet", "'Net'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        ilTemp = Val(RptSelCt!edcSelCTo.Text)
        If ilTemp = 0 Then
            ilTemp = 13
        End If
        If Not gSetFormula("NumberPeriods", ilTemp) Then
            mCntJob11Plus = -1
            Exit Function
        End If
        'determine if any vehicle groups selected
        illoop = RptSelCt!cbcSet1.ListIndex
        If illoop = 0 Then
            If Not gSetFormula("ShowVehGrp", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowVehGrp", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

         'spot bus booked by vehicle or agency uses the same rpt module
        If RptSelCt!rbcSelCInclude(2).Value = True Or RptSelCt!rbcSelCInclude(3).Value = True Then      'vehicle or agency options?
            If RptSelCt!ckcSelC8(0).Value = vbChecked Then           'summary
                If Not gSetFormula("TotalsBy", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("TotalsBy", "'D'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
            If RptSelCt!rbcSelCInclude(2).Value = True Then     'vehicle
                If Not gSetFormula("SortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
            If RptSelCt!rbcSelCInclude(3).Value = True Then     'agency
                If Not gSetFormula("SortBy", "'G'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        End If

        '4-9-08 if adjustments are included, show in report header
        slStr = ""
        If RptSelCt!ckcSelC10(0).Value = vbChecked Then
            gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
            If Trim$(slStr) = "" Then
                slStr = "1/1/1975"
            End If
            llDate = gDateValue(slStr)           'convert last month billed to long

            slStr = ", Last Billed Date: " & Format$(llDate, "m/d/yy")
        End If
        If Not gSetFormula("ShowLastBilled", "'" & slStr & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

        If ilListIndex = CNT_BOB_BYSPOT Then
            slDate = RptSelCt!CSI_CalFrom.Text      'Date: 1/7/2020 added CSI calendar controls for date entries --> edcSelCFrom.Text
            If RptSelCt!rbcSelCSelect(0).Value Then  'set last sunday of first week
                slType = "Weekly"
            ElseIf RptSelCt!rbcSelCSelect(1).Value Then  'set last date of 12 standard periods
                slType = "Standard"
            ElseIf RptSelCt!rbcSelCSelect(2).Value Then  'set last date of 12 corporate periods
                slType = "Corporate"
            ElseIf RptSelCt!rbcSelCSelect(3).Value Then  'set last date of 12 calendar periods
                slType = "Calendar"
            End If
        Else    'Get date from combo box
            slDate = RptSelCt!cbcSel.List(RptSelCt!cbcSel.ListIndex)
            ilRet = gParseItem(slDate, 3, " ", slDate)    'Get application name
            slType = RptSelCt!cbcSel.List(RptSelCt!cbcSel.ListIndex)
            ilRet = gParseItem(slType, 4, " ", slType)    'Get application name
        End If
        If gValidDate(slDate) Then
            If slType = "Weekly" Then  'set last sunday of first week
                If Not gSetFormula("Type", "'Weekly '") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                sgPdType = "W"
                slDate = gObtainPrevMonday(slDate)
                gPackDate slDate, igPdStartDate(0), igPdStartDate(1)
                slDate = gObtainNextSunday(slDate)
                For illoop = 1 To ilTemp   '1-31-00   was 13    , loop on # periods to print
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    slDate = gIncOneWeek(slDate)
                Next illoop
            ElseIf slType = "Standard" Then  'set last date of 12 standard periods
                If Not gSetFormula("Type", "'Standard Month '") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(0) = "Type= Standard Month,"
                sgPdType = "S"
                slDate = gObtainStartStd(slDate)
                gPackDate slDate, igPdStartDate(0), igPdStartDate(1)
                ilMonth = Month(Format$(gDateValue(slDate) + 15, "m/d/yy"))
                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'End date of previous month
                For illoop = 1 To ilTemp   '1-31-00 was 12 , loop on # periods to print
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    slDate = gObtainEndStd(slDate)
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            ElseIf slType = "Calendar" Then  'set last date of 12 standard periods
                If Not gSetFormula("Type", "'Calendar Month '") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(0) = "Type= Standard Month,"
                sgPdType = "C"
                slDate = gObtainStartCal(slDate)
                gPackDate slDate, igPdStartDate(0), igPdStartDate(1)
                ilMonth = Month(Format$(gDateValue(slDate) + 15, "m/d/yy"))
                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'End date of previous month
                For illoop = 1 To ilTemp   '1-31-00 was 12 , loop on # periods to print
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    slDate = gObtainEndCal(slDate)
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            ElseIf slType = "Corporate" Then  'set last date of 12 corporate periods
                If Not gSetFormula("Type", "'Corporate Month '") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(0) = "Type= Corporate Month,"
                sgPdType = "F"
                slDate = gObtainStartCorp(slDate, True)
                gPackDate slDate, igPdStartDate(0), igPdStartDate(1)
                ilMonth = Month(Format$(gDateValue(slDate) + 15, "m/d/yy"))
                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'End date of previous month
                'For ilLoop = ilMonth To 12 Step 1
                For illoop = 1 To 12 Step 1
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    slDate = gObtainEndCorp(slDate, True)
                    '7-17-09 ensure the entire corporate calendar defined
                    If slDate = "" Then     'invalid date returned if blank
                        MsgBox "Corporate Calendar must be defined for next year in Site"
                        'mReset
                        'RptSelCt!edcSelCFrom.SetFocus
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            End If
        Else
            mReset
            RptSelCt!CSI_CalFrom.SetFocus       'Date: 1/7/2020 added CSI calendar controls for date entries --> edcSelCFrom.SetFocus
            Exit Function
        End If
    ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Then    'Quarterly Avails by min or pct
        slDate = RptSelCt!edcSelCFrom.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        slStr = (RptSelCt!edcSelCFrom1.Text)
        ilRet = gVerifyInt(slStr, 1, 53)                    '53 weeks max
        If ilRet = -1 Then
            mReset
            RptSelCt!edcSelCFrom1.SetFocus                 'invalid
            Exit Function
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get current date and time for headers or keys to prepass files
        slSelection = "{AVR_Quarterly_Avails.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({AVR_Quarterly_Avails.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        If Not gSetFormula("AsOfD", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            mCntJob11Plus = -1
            Exit Function
        End If
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
'        If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'            mCntJob11Plus = -1
'            Exit Function
'        End If
        slExclude = ""
        slInclude = ""

        gIncludeExcludeCkc RptSelCt!ckcSelC10(0), slInclude, slExclude, "Holds"     '3-8-06 chged from ckcselc3
        gIncludeExcludeCkc RptSelCt!ckcSelC10(1), slInclude, slExclude, "Orders"     '3-8-06 chg from ckcselc3

        gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(5), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(6), slInclude, slExclude, "Promo"

        gIncludeExcludeCkc RptSelCt!ckcSelC6(0), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(1), slInclude, slExclude, "Missed"
        '5-16-05 nc, fills, all the spot types are now diff. options to select
        'gIncludeExcludeCkc RptSelCt!ckcSelC6(2), slInclude, slExclude, "N/C"
        'gIncludeExcludeCkc RptSelCt!ckcSelC6(3), slInclude, slExclude, "Fill"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(4), slInclude, slExclude, "Locked Avails"
        '5-16-05
        gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Charge"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "0.00"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(2), slInclude, slExclude, "ADU"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(3), slInclude, slExclude, "Bonus"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(4), slInclude, slExclude, "+Fill"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(5), slInclude, slExclude, "-Fill"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(6), slInclude, slExclude, "N/C"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(7), slInclude, slExclude, "MG"        'currently mg cost type always included
        gIncludeExcludeCkc RptSelCt!ckcSelC3(8), slInclude, slExclude, "Recap"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(9), slInclude, slExclude, "Spinoff"

        'gIncludeExcludeCkc RptSelCt!ckcSelC12(0), slInclude, slExclude, "Local"        contract spots are shown as Holds & Orders
        If tgSpf.sSystemType = "R" Then                 'only show Feeds if Radio statin
            gIncludeExcludeCkc RptSelCt!ckcSelC12(0), slInclude, slExclude, "Feed"
        End If
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        'One or two column avail report
        If RptSelCt!rbcSelC11(1).Value Then         'two column
            If Not gSetFormula("Columns1or2", "'2'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else        'one column
            If Not gSetFormula("Columns1or2", "'1'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If RptSelCt!rbcSelCSelect(3).Value Then   'Percent option, override previous formula
            If Not gSetFormula("Columns1or2", "'P'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If RptSelCt!rbcSelC4(1).Value Then           'qtrly detail
            'Decide what to do with the reserves
            slReserve = "H"
            If RptSelCt!rbcSelC7(1).Value Then           'show the reserves
                slReserve = "S"
            End If

            If Not gSetFormula("Reserve", "'" & slReserve & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
        If Not gSetFormula("RateCard", "'" & slNameCode & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

    ElseIf (ilListIndex = CNT_AVG_PRICES) Then     'avg spot prices
        slDate = RptSelCt!CSI_CalFrom.Text          'Date: 11/5/2019 using CSI calendar for date entry -->  edcSelCFrom.Text        'obtain date entered
        If gValidDate(slDate) Then
            slExclude = ""
            slInclude = ""
            If RptSelCt!rbcSelCSelect(0).Value Then   'weekly
                If gWeekDayStr(slDate) = 0 Then         'OK, its a monday
                    For illoop = 1 To 14 Step 1
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If Not gSetFormula("Per" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                        slDate = gIncOneWeek(slDate)            'incr one week
                    Next illoop
                    If Not gSetFormula("AvgInterval", "'W'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else                                    'not a monday
                    MsgBox "Enter Monday start date"
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus       'Date: 11/5/2019 using CSI calendar for date entry --> edcSelCFrom.SetFocus
                    Exit Function
                End If
            Else                                    'monthly
                slDate = gObtainStartStd(slDate)    'obtain std bdcst start date of date entered
                For illoop = 1 To 14 Step 1
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Per" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    slDate = gObtainEndStd(slDate)
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                Next illoop
                If Not gSetFormula("AvgInterval", "'M'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If                                  'rbcSelCSelect(0)
            'Setup selections for database
            'always exclude '= proposal,  s = psa, m = promo
            'always exclude package lines (ordered and aired)
'            slSelection = "{CLF_Contract_Line.clfType} <> 'A' and {CLF_Contract_Line.clfType} <> 'O' and {CLF_Contract_Line.clfType} <> 'E' and {CHF_Contract_Header.chfStatus} <> 'W' And  {CHF_Contract_Header.chfStatus} <> 'C' And {CHF_Contract_Header.chfStatus} <> 'I' And {CHF_Contract_Header.chfStatus} <> 'D'  And {CHF_Contract_Header.chfType} <> 'M' And {CHF_Contract_Header.chfType} <> 'S'"


            ilRet = mAvgReptOptions(slSelection, slExclude, slInclude)       'format the filter for Crystal, and format description fields
            If ilRet = -1 Then
                mCntJob11Plus = -1
                Exit Function
            End If

'            ilRet = mAvgUnitsOptions(slSelection, slExclude, slInclude)       'format the filter for Crystal, and format description fields
'            If ilRet = -1 Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If

'            If Not RptSelCt!ckcAll.Value = vbChecked Then         '9-12-02 not all vehicles selected
'                If slSelection <> "" Then
'                    slSelection = "(" & slSelection & ") " & " and ("
'                    slOr = ""
'                Else
'                    slSelection = "("
'                    slOr = ""
'                End If
'                If RptSelCt!rbcSelCInclude(0).Value Then        'slsp option
'                    For ilLoop = 0 To RptSelCt!lbcSelection(2).ListCount - 1 Step 1
'                        If RptSelCt!lbcSelection(2).Selected(ilLoop) Then
'                            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
'                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                            slSelection = slSelection & slOr & "{CHF_Contract_Header.chfslfCode1} = " & Trim$(slCode)
'                            slOr = " Or "
'                        End If
'                    Next ilLoop
'                Else
'                    'setup selective vehicles
'                    For ilLoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
'                        If RptSelCt!lbcSelection(6).Selected(ilLoop) Then
'                            slNameCode = tgCSVNameCode(ilLoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
'                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                            slSelection = slSelection & slOr & "{CLF_Contract_Line.clfvefCode} = " & Trim$(slCode)
'                            slOr = " Or "
'                        End If
'                    Next ilLoop
'                End If
'                slSelection = slSelection & ")"
'            End If
'            If RptSelCt!rbcSelCSelect(0).Value Then   'weekly
'                slSelection = "(" & slSelection & ")" & " And ({CLF_Contract_Line.clfDelete} <> 'Y' and {CFF_Contract_Flight.cffDelete} <> 'Y' and {CFF_Contract_Flight.cffEndDate} >= {CFF_Contract_Flight.cffStartDate} and {CFF_Contract_Flight.cffEndDate} >= {@Per1} and {CFF_Contract_Flight.cffStartDate} < {@Per14})"
'            Else                                    'monthly
'                slSelection = "(" & slSelection & ")" & " And ({CLF_Contract_Line.clfDelete} <> 'Y' and {CFF_Contract_Flight.cffDelete} <> 'Y' and {CFF_Contract_Flight.cffEndDate} >= {CFF_Contract_Flight.cffStartDate} and {CFF_Contract_Flight.cffEndDate} >= {@Per1} and {CFF_Contract_Flight.cffStartDate} < {@Per13})"
'            End If

            'Slsp or vehicle option
            'Date: 11/5/2019 commented out; added major/minor sorts
'            If RptSelCt!rbcSelCInclude(0).Value Then            'slsp
'                If Not gSetFormula("SortBy", "'S'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            Else
'                If Not gSetFormula("SortBy", "'V'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            End If

            'Date: 8/30/2019 pass major/minor sort parameter to crystal report
            ilVehicleGroup = 0
            If RptSelCt!cbcSet1.ListIndex = 0 Then
                If Not gSetFormula("MajorSortBy", "'A'") Then   'advt
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
                If Not gSetFormula("MajorSortBy", "'G'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
                If Not gSetFormula("MajorSortBy", "'B'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
                If Not gSetFormula("MajorSortBy", "'P'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 4 Then          'slsp
                If Not gSetFormula("MajorSortBy", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then            '4-25-06 vehicle option added
                If Not gSetFormula("MajorSortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                            'vehicle group
                If Not gSetFormula("MajorSortBy", "'H'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                illoop = RptSelCt!lbcSelection(12).ListIndex            '3-18-16 chg from lbcselection(4)
                ilVehicleGroup = tgVehicleSets1(illoop).iCode
            End If
            If Not gSetFormula("MajorVehicleGroupHdr", ilVehicleGroup) Then
                mCntJob11Plus = -1
                Exit Function
            End If
    
            ilVehicleGroup = 0
            If RptSelCt!cbcSet2.ListIndex = 0 Then              'no minor sort selected
                If Not gSetFormula("MinorSortBy", "''") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then          'advt
                If Not gSetFormula("MinorSortBy", "'A'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then          'agency
                If Not gSetFormula("MinorSortBy", "'G'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then          'bus cat
                If Not gSetFormula("MinorSortBy", "'B'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then          'prod prot
                If Not gSetFormula("MinorSortBy", "'P'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 5 Then          'slsp
                If Not gSetFormula("MinorSortBy", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then            'vehicle
                If Not gSetFormula("MinorSortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                            'vehicle group
                If Not gSetFormula("MinorSortBy", "'H'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                'get the vehicle group selected for report heading (participant, format, market, etc)
                illoop = RptSelCt!lbcSelection(4).ListIndex
                ilVehicleGroup = tgVehicleSets1(illoop).iCode
            End If
            If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
                mCntJob11Plus = -1
                Exit Function
            End If
    
            ilNoneExists = True                    'NONE  allowed in this list
            ilMinorGroupHdr = True                 'there is no minor vehicle group hdr to send to crystal for ths report
            If mCBCSet2Test(ilNoneExists, ilMinorGroupHdr) Then
                mCntJob11Plus = -1
                Exit Function
            End If

'----   Major/Minor sorts

             'Slsp or vehicle option
            If RptSelCt!ckcSelC13(0).Value = vbChecked Then             'use Sales source as major sort
                If Not gSetFormula("UseSSAsMajor", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("UseSSAsMajor", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If

            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

            If Not gSetSelection(slSelection) Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                    'not valid date
            mReset
            RptSelCt!CSI_CalFrom.SetFocus   'Date: 11/5/2019 using CSI calendar for date entry --> edcSelCFrom.SetFocus
            Exit Function
        End If                                  'gValidDate(slDate)
    ElseIf (ilListIndex = CNT_ADVT_UNITS) Then              'advert units sold
        slStr = RptSelCt!CSI_CalFrom.Text       'using CSI calendar for date entry --> edcSelCFrom1.Text            '6-21-18 get # of weeks
        ilRet = gVerifyInt(slStr, 1, 13)
        If ilRet = -1 Then                          'bad conversion or illegal #
            mReset
            RptSelCt!CSI_CalFrom.SetFocus       'using CSI calendar for date entry --> edcSelCFrom1.SetFocus                 'invalid # weeks
            mCntJob11Plus = -1
            Exit Function
        End If
        
        'Date: 8/30/2019 pass major/minor sort parameter to crystal report
        If RptSelCt!cbcSet1.ListIndex = 0 Then
            If Not gSetFormula("MajorSortBy", "'A'") Then   'advt
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
            If Not gSetFormula("MajorSortBy", "'G'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
            If Not gSetFormula("MajorSortBy", "'B'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
            If Not gSetFormula("MajorSortBy", "'P'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet1.ListIndex = 4 Then          'slsp
            If Not gSetFormula("MajorSortBy", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then            '4-25-06 vehicle option added
            If Not gSetFormula("MajorSortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                            'vehicle group
            If Not gSetFormula("MajorSortBy", "'H'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
            illoop = RptSelCt!lbcSelection(12).ListIndex            '3-18-16 chg from lbcselection(4)
            ilVehicleGroup = tgVehicleSets1(illoop).iCode
        End If
        If Not gSetFormula("MajorVehicleGroupHdr", ilVehicleGroup) Then
            mCntJob11Plus = -1
            Exit Function
        End If

        ilVehicleGroup = 0
        If RptSelCt!cbcSet2.ListIndex = 0 Then              'no minor sort selected
            If Not gSetFormula("MinorSortBy", "''") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then          'advt
            If Not gSetFormula("MinorSortBy", "'A'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then          'agency
            If Not gSetFormula("MinorSortBy", "'G'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then          'bus cat
            If Not gSetFormula("MinorSortBy", "'B'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then          'prod prot
            If Not gSetFormula("MinorSortBy", "'P'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 5 Then          'slsp
            If Not gSetFormula("MinorSortBy", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then            'vehicle
            If Not gSetFormula("MinorSortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                            'vehicle group
            If Not gSetFormula("MinorSortBy", "'H'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
            'get the vehicle group selected for report heading (participant, format, market, etc)
            illoop = RptSelCt!lbcSelection(4).ListIndex
            ilVehicleGroup = tgVehicleSets1(illoop).iCode
        End If
        If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
            mCntJob11Plus = -1
            Exit Function
        End If

        ilNoneExists = True                    'NONE  allowed in this list
        ilMinorGroupHdr = True                 'there is no minor vehicle group hdr to send to crystal for ths report
        If mCBCSet2Test(ilNoneExists, ilMinorGroupHdr) Then
            mCntJob11Plus = -1
            Exit Function
        End If

        'if Advertiser is primary selection, force to show the advt totals; or if the
        'Include advt subtotals is checked because major and minor do not include the advt totals
        'If (RptSelCt!cbcSet1.ListIndex = 0 And RptSelCt!cbcSet1.ListIndex = 0) Or RptSelCt!cbcSet2.ListIndex = 1 Or RptSelCt!ckcSelC13(0).Value = vbChecked Then
'        If RptSelCt!ckcSelC13(0).Value = vbChecked Then
'            If Not gSetFormula("InclAdvtTotals", "'Y'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        Else
'            If Not gSetFormula("InclAdvtTotals", "'N'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If

''        If RptSelCt!ckcSelC13(1).Value = vbChecked Then
''            If Not gSetFormula("SeparatePoliticals", "'Y'") Then
''                mCntJob11Plus = -1
''                Exit Function
''            End If
''        Else
''            If Not gSetFormula("SeparatePoliticals", "'N'") Then
''                mCntJob11Plus = -1
''                Exit Function
''            End If
''        End If

'        If mBOBCrystal() < 0 Then                           'send Crystl formulas for Header notations (pkg vs hidden),
'            mCntJob11Plus = -1                              'As of Time,  Gross, Net
'            Exit Function
'        End If
        
        '10-13-10 new page each major sort
        If RptSelCt!ckcSelC8(1).Value = vbChecked Then
            If Not gSetFormula("NewPage", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        
        slDate = RptSelCt!CSI_CalFrom.Text  'Date: 9/22/2019    using CSI calendar for date entry   --> edcSelCFrom.Text         'obtain date entered
        'If (gValidDate(slDate)) And (gWeekDayStr(slDate) = 0) Then
        If (gValidDate(slDate)) Then
            If (gWeekDayStr(slDate) = 0) Then
                For illoop = 1 To 13 Step 1
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Wk " & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                    slDate = gIncOneWeek(slDate)            'incr one week
                Next illoop
                If RptSelCt!rbcSelC4(0).Value Then           'Spot counts (vs unit count)
                    If Not gSetFormula("UnitsOrSpots", "'S'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("UnitsOrSpots", "'U'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
                If RptSelCt!ckcSelC8(0).Value = vbChecked Then           'Show spot rates (y/N)
                    If Not gSetFormula("ShowRate", "'Y'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowRate", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
                If RptSelCt!ckcSelC8(1).Value = vbChecked Then           'new page each vehicle
                    If Not gSetFormula("Skip", "'Y'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("Skip", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
                
                '6-21-18
                'send # of weeks selected
                ilPeriods = Val(RptSelCt!edcSelCFrom1.Text)
                If Not gSetFormula("Weeks", ilPeriods) Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
    
    '            '9-23-09 gross, net tnet option
                If RptSelCt!rbcSelC9(0).Value = True Then           'gross
                    If Not gSetFormula("GrossNetTNet", "'G'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                ElseIf RptSelCt!rbcSelC9(1).Value = True Then           'net
                    If Not gSetFormula("GrossNetTNet", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("GrossNetTNet", "'T'") Then      't-net
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
                'Setup selections for database
                'Ignore all package type lines
                'always exclude '= proposal, j = rejection, h = hold, s = psa, m = promo , Q = PI, and altered cnts
    '            slSelection = "{CLF_Contract_Line.clfType} <> 'A' and {CLF_Contract_Line.clfType} <> 'O' and {CLF_Contract_Line.clfType} <> 'E' and {CHF_Contract_Header.chfStatus} <> 'W' And {CHF_Contract_Header.chfStatus} <> 'C' And {CHF_Contract_Header.chfStatus} <> 'I' And {CHF_Contract_Header.chfStatus} <> 'D' And {CHF_Contract_Header.chfType} <> 'M' And {CHF_Contract_Header.chfType} <> 'S' And {CHF_Contract_Header.chfSchStatus} <> 'A'"
    '            slSelection = "(" & slSelection & ")" & " And ({CLF_Contract_Line.clfDelete} <> 'Y' and {CFF_Contract_Flight.cffDelete} <> 'Y' and {CFF_Contract_Flight.cffEndDate} >= {CFF_Contract_Flight.cffStartDate} and {CFF_Contract_Flight.cffEndDate} >= {@Wk 1} and {CFF_Contract_Flight.cffStartDate} < {@Wk 13}+7)"
    
                ilRet = mAvgReptOptions(slSelection, slExclude, slInclude)       'format the filter for Crystal, and format description fields
                'ilRet = mAvgUnitsOptions(slSelection, slExclude, slInclude)       'format the filter for Crystal, and format description fields
                If ilRet = -1 Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                
    '            If Not RptSelCt!ckcAll.Value = vbChecked Then         '9-12-02 not all vehicles selected
    '                If slSelection <> "" Then
    '                    slSelection = "(" & slSelection & ") " & " and ("
    '                    slOr = ""
    '                Else
    '                    slSelection = "("
    '                    slOr = ""
    '                End If
    '                'setup selective vehicles
    '                For ilLoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
    '                    If RptSelCt!lbcSelection(6).Selected(ilLoop) Then
    '                        slNameCode = tgCSVNameCode(ilLoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
    '                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
    '                        slSelection = slSelection & slOr & "{CLF_Contract_Line.clfvefCode} = " & Trim$(slCode)
    '                        slOr = " Or "
    '                    End If
    '                Next ilLoop
    '                slSelection = slSelection & ")"
    '            End If
    
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
                 If Not gSetSelection(slSelection) Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                MsgBox "Enter Monday start date"
                mReset
                RptSelCt!CSI_CalFrom.SetFocus       'Date: 9/22/2019    use CSI calendar for date entry -->edcSelCFrom.SetFocus
                Exit Function
            End If
        Else
            mReset
            RptSelCt!CSI_CalFrom.SetFocus           'Date: 9/22/2019    use CSI calendar for date entry -->edcSelCFrom.SetFocus
            Exit Function
        End If
    ElseIf ilListIndex = CNT_SALES_CPPCPM Then
        slDate = RptSelCt!CSI_CalFrom.Text          'Date: 12/10/2019 added CSI calendar control for date entry -->  edcSelCFrom.Text
        If Not gValidDate(slDate) Then
            mReset
             RptSelCt!CSI_CalFrom.SetFocus          'Date: 12/10/2019 added CSI calendar control for date entry -->  edcSelCFrom.SetFocus
            Exit Function
        End If
        slStr = RptSelCt!edcSelCTo.Text                 'entered year
        igYear = gVerifyYear(slStr)
        If igYear = 0 Then
            mReset
            RptSelCt!edcSelCTo.SetFocus                 'invalid year
            mCntJob11Plus = -1
            Exit Function
        End If
        If RptSelCt!rbcSelCInclude(0).Value Then             'corporate procesing
            'check to see if budget year exists
            ilRet = gGetCorpCalIndex(igYear)
            If ilRet < 0 Then
                MsgBox "Corporate Year Missing for" & str$(igYear), vbOKOnly + vbExclamation, "RptselCt"
                mReset
                RptSelCt!edcSelCTo.SetFocus                 'invalid year
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        slStr = RptSelCt!edcSelCTo1.Text                  'edit qtr
        ilRet = gVerifyInt(slStr, 1, 4)
        If ilRet = -1 Then
            mReset
            RptSelCt!edcSelCTo1.SetFocus                 'invalid qtr
            mCntJob11Plus = -1
            Exit Function
        End If
        igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
        If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not mSendCorpStd(RptSelCt!rbcSelCInclude(0), RptSelCt!rbcSelCInclude(1)) Then           'pass corp or std as formula to crystal report
            mCntJob11Plus = -1
            Exit Function
        End If
        
        slStr = "G"
        If RptSelCt!rbcSelC4(1).Value Then
            slStr = "N"
        End If
        If Not gSetFormula("GrossNet", "'" & Trim$(slStr) & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not RptSelCt!ckcAll.Value = vbChecked Then         '9-12-02 not all vehicles selected
            If slSelection <> "" Then
                slSelection = "(" & slSelection & ") " & " and ("
                slOr = ""
            Else
                slSelection = "("
                slOr = ""
            End If
            'setup selective vehicles
            For illoop = 0 To RptSelCt!lbcSelection(11).ListCount - 1 Step 1
                If RptSelCt!lbcSelection(11).Selected(illoop) Then
                    slNameCode = tgRptSelDemoCodeCT(illoop).sKey 'RptSelCt!lbcDemoCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get demo name
                    slSelection = slSelection & slOr & "{CBF_Contract_BR.cbfdnfCode} = " & Trim$(slCode)
                    slOr = " Or "
                End If
            Next illoop
            slSelection = slSelection & ")"
        End If
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_AVGRATE Then       'Average rate
        ilPeriods = 14                          'default for max # weeks to print
        If RptSelCt!rbcSelCSelect(0).Value Then     'week
            If RptSelCt!rbcOutput(4).Value = False Then
                If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
                    mCntJob11Plus = -1
                    Exit Function
                End If
                igMonthOrQtr = Val(RptSelCt!edcSelCTo1.Text)
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("W1", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        Else                                    '9-28-11 month option
            slDate = Trim$(str(igMonthOrQtr)) & "/15/" & Trim$(str(igYear))
            slDate = gObtainStartStd(slDate)
            slCode = gObtainEndStd(slDate)
            slStr = gMonthYearFormat(slCode)
            'strip out comma from string
            slMonth = ""
            For illoop = 1 To Len(slStr)
                If Mid(slStr, illoop, 1) <> "," Then
                    slMonth = slMonth & Mid(slStr, illoop, 1)
                End If
            Next illoop
            slMonth = "Std " & slMonth      'this text has to be formatted as is.  If changed, need to look
                                            'at crystal formula (StartOfYear) to change starting index
            If Not gSetFormula("WeekQtrHeader", "'" & slMonth & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If

            For illoop = 1 To 12   '1-31-00 was 12 , loop on # periods to print
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("W" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                slDate = gObtainEndStd(slDate)
                slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
            Next illoop
            ilPeriods = Val(RptSelCt!edcSelCFrom1.Text)
        End If
        If RptSelCt!rbcOutput(4).Value = False Then
            ilRet = mAvgReptOptions(slSelection, slExclude, slInclude)       'format the filter for Crystal, and format description fields
            If ilRet = -1 Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

'        slExclude = ""
'        slInclude = ""
'        gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Holds"
'        gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Orders"
'
'        gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
'        gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
'        gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
'        gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
'        gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
'        gIncludeExcludeCkc RptSelCt!ckcSelC6(0), slInclude, slExclude, "Trade"
'        gIncludeExcludeCkc RptSelCt!ckcSelC6(2), slInclude, slExclude, "N/C"
'
'        If Len(slInclude) > 0 Then
'            If Not gSetFormula("Included", "'" & slInclude & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If
'        If Len(slExclude) > 0 Then
'            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If
'        If Len(slInclude) > 0 Then
'            If Not gSetFormula("Included", "'" & slInclude & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If
'        If Len(slExclude) > 0 Then
'            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If

        If tgSpf.sUsingBBs = "Y" Then
            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then   'use closest avail, dont know if its open or close
            'use closest, any avail.  send formula for legend
                If Not gSetFormula("BBLegend", "'A'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                'find specific avail for bb
                If Not gSetFormula("BBLegend", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        Else
            If Not gSetFormula("BBLegend", "'N'") Then      'not used
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        If RptSelCt!rbcSelC7(0).Value Then          'using dp name
             If Not gSetFormula("DPType", "'using Dayparts without Overrides'") Then      'not used
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!rbcSelC7(1).Value Then     '12-9-16 using dp with override
            If Not gSetFormula("DPType", "'using Dayparts with Overrides'") Then      'not used
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                        'agency option
             If Not gSetFormula("DPType", "'by Agency'") Then      'not used
                mCntJob11Plus = -1
                Exit Function
            End If
        End If


        ilRet = RptSelCt!cbcSet1.ListIndex
        If Not gSetFormula("VehicleGroup", ilRet) Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        ilPeriods = Val(RptSelCt!edcSelCFrom1.Text)
        If Not gSetFormula("ShowPeriods", ilPeriods) Then
            mCntJob11Plus = -1
            Exit Function
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_TIEOUT Then
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
             'dan M 7-28-08 NTR/Hard Cost option
        slStr = ""
        If RptSelCt!ckcSelC3(0).Value = 1 Or RptSelCt!ckcSelC3(1).Value = 1 Then
            slStr = slStr & "'With"
            If RptSelCt!ckcSelC3(0).Value = 1 Then
                slStr = slStr & " NTR"
                If RptSelCt!ckcSelC3(1).Value = 1 Then
                    slStr = slStr & " and Hard Cost"
                End If
            Else
                slStr = slStr & " Hard Cost"
            End If
        End If
        If slStr <> "" Then
            slStr = slStr & "'"
        End If
        If Not gSetFormula("reporttitle", slStr) Then
            mCntJob11Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_BOB Then
        If RptSelCt!edcText.Text <> "" Then
            slDate = RptSelCt!edcText.Text
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!edcText.SetFocus
                Exit Function
            End If
        End If

        If mBOBCrystal() < 0 Then                                'send Crystl formulas for Header notations (pkg vs hidden),
            mCntJob11Plus = -1                                 'As of Time,  Gross, Net
            Exit Function
        End If
        '
        'rbcSelCInclude: SORT BY OPTIONS =Adv, 1 = slsp, 2 = vehicle, 3 = owner , 4 = vehicle/participant, 5 = agy, 6 = vehicle gross/net
        '
        If RptSelCt!rbcSelCInclude(0).Value Then
            If Not gSetFormula("SortBy", "'A'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!rbcSelCInclude(1).Value Then
            If Not gSetFormula("SortBy", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!rbcSelCInclude(2).Value Then      'vehicle
            If Not gSetFormula("SortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
            't-net doesnt cant subsort by slsp
            If RptSelCt!rbcSelC7(2).Value = False Then          'check if t-net selected
                If RptSelCt!ckcSelC10(1).Value = vbChecked Then     'subsort & total slsp with vehicle
                    If Not gSetFormula("L3AOptionSlspThenVeh", "'Y'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("L3AOptionSlspThenVeh", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
            Else                '8-6-10 its tnet option by vehicle, slsp subt disallowed
                If Not gSetFormula("L3AOptionSlspThenVeh", "'N'") Then
                        mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!rbcSelCInclude(4).Value Then    'vehicle/participant
            If Not gSetFormula("SortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
                If Not gSetFormula("L3AOptionSlspThenVeh", "'N'") Then      'no subtotals by slsp within vehicle
                    mCntJob11Plus = -1
                    Exit Function
                End If
        '4-12-02 add agency option
        ElseIf RptSelCt!rbcSelCInclude(5).Value Then     '2-2-03 take out test for vehicle/participant
            If Not gSetFormula("SortBy", "'G'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SortBy", "'O'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        ilRet = mBOBMonthHeader()       'format qtr/year & month heading
        If ilRet <> 0 Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not RptSelCt!rbcSelCInclude(1).Value Then        'not slsp option
            If RptSelCt!rbcSelCInclude(4).Value Then     '8-4-00 vehicle with participant splits
                If Not gSetFormula("TrickIfOwner", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                '4-11-16 is it vehicle with participant vg?
                If RptSelCt!rbcSelCInclude(2).Value = True Then         'vehicle option
                    ' if participant vg selected.  if so, trik report output into thinking its split participants
                    illoop = RptSelCt!cbcSet1.ListIndex             'retrieve the vehicle set selected (applies to vehicle option)
                    If tgVehicleSets1(illoop).iCode = 1 Then        '2-16-16 participant vehicle group selected
                        If Not gSetFormula("TrickIfOwner", "'Y'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("TrickIfOwner", "'N'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    End If
                   
                Else
                    If Not gSetFormula("TrickIfOwner", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
            End If
            If tgUrf(0).iSlfCode > 0 Then           'slsp signed in, alert user $ are not split
                If Not gSetFormula("SplitLegendFlag", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SplitLegendFlag", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        Else                    '8-8-06 slsp option, show amounts reduced to participation if showing splits
'            If RptSelCt!ckcSelC10(0).Value = vbChecked Then     'show splits
'                If Not gSetFormula("TrickIfOwner", "'Y'") Then
'                        mCntJob11Plus = -1
'                        Exit Function
'                End If
'            Else
                If Not gSetFormula("TrickIfOwner", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
'            End If
            'show legend at bottom if the slsp option didnt select show splits
            If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then       'And tgUrf(0).iSlfCode > 0 Then   'show splits option not set with a slsp signed in
                If Not gSetFormula("SplitLegendFlag", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SplitLegendFlag", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If

        End If
        '2-28-01 if net-net option by vehicle option, skip pages by vehicle?
        If RptSelCt!rbcSelC7(2).Value = True And RptSelCt!rbcSelCInclude(6).Value = True Then           'net -net
            If RptSelCt!ckcSelC8(2).Value = vbChecked Then              'skip new page
                If Not gSetFormula("SkipPage", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SkipPage", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If

            ilRet = mBobTotals()        'how to show totals (tots by contr, advt, summary)
            If ilRet <> 0 Then
                mCntJob11Plus = -1
                Exit Function
            End If

        Else
            '10-30-02 all versions will use the SkipPage formula
            'If RptSelCt!rbcSelC4(0).Value Or RptSelCt!rbcSelC4(1).Value Then            'detail or advt totals
            '    If Not RptSelCt!rbcSelCInclude(1).Value Then   ' check if  slsp option
                    If RptSelCt!ckcSelC8(2).Value = vbChecked Then              'skip new page
                        If Not gSetFormula("SkipPage", "'Y'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("SkipPage", "'N'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    End If
               ' Else                '10-30-02 slsp option
               '     If RptSelCt!ckcSelC10(1).Value = vbChecked Then     ' vehicle subtotals by slsP?
               '         If RptSelCt!ckcSelC8(2).Value = vbChecked Then  'skip new page
               '             If Not gSetFormula("SkipPage", "'Y'") Then
               '                 mCntJob11Plus = -1
               '                 Exit Function
               '             End If
               '         Else
               '             If Not gSetFormula("SkipPage", "'N'") Then
               '                 mCntJob11Plus = -1
               '                 Exit Function
               '             End If
               '         End If
               '    End If
               ' End If
            'End If
        End If
        '12-3-00   if option by vehicle or vehicle/participant & summary, allow sub-totals by vehicle to be suppressed
        If RptSelCt!rbcSelC4(2).Value And Not RptSelCt!rbcSelCInclude(1).Value Then      'summary only , but not slsp option
            If Not RptSelCt!ckcSelC8(2).Value = vbChecked Then              'skip new page doesnt suppress the vehicle subtotals
                If (RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(4).Value) And RptSelCt!ckcSelC10(1).Value = vbChecked Then   'veh or veh/part option and show sub-tots by vehicle
                    If Not gSetFormula("ShowVehSubTotal", "'Y'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                Else
                    '12-14-00 if vehicle or vehicle/participant and suppress vehicle subtotals
                    If (RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(4).Value) And Not RptSelCt!ckcSelC10(1).Value = vbChecked Then   'veh or veh/part option and show sub-tots by vehicle

                        If Not gSetFormula("ShowVehSubTotal", "'N'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowVehSubTotal", "'Y'") Then
                            mCntJob11Plus = -1
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        '11-21-05 Send formula indicating a pacing report
        If Trim$(RptSelCt!edcText.Text) <> "" Then
            slDate = RptSelCt!edcText.Text              'get pacing date for header
            slDate = Format$(gDateValue(slDate), "m/d/yy")

            If Not gSetFormula("PacingDate", "'" & slDate & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("PacingDate", "''") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        
        '11-7-16 option implemented to suppress office subtotals when selection by salesperson
        If RptSelCt!rbcSelCInclude(1).Value Then            'slsp option
            'show/hide the office subtotals
            If RptSelCt!ckcSelC10(2).Value = vbChecked Then        'show office subtotals
                If Not gSetFormula("OfcSubTotal", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                                'hide office subtotals
                If Not gSetFormula("OfcSubTotal", "'H'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        Else
            If Not gSetFormula("OfcSubTotal", "'H'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

    ElseIf ilListIndex = CNT_BOBRECAP Then              '4-14-05
        If mBOBCrystal() < 0 Then                                'send Crystl formulas for Header notations (pkg vs hidden),
            mCntJob11Plus = -1                                 'As of Time,  Gross, Net
            Exit Function
        End If
        ilRet = mBOBMonthHeader()       'format qtr/year & month heading
        If ilRet <> 0 Then
            mCntJob11Plus = -1
            Exit Function
        End If

        ilRet = mBobTotals()        'how to show totals (tots by contr, advt, summary)
        If ilRet <> 0 Then
            mCntJob11Plus = -1
            Exit Function
        End If

        If RptSelCt!rbcSelC11(0).Value Then       'sort by vehicle
            If Not gSetFormula("IncludeVehicleTotals", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
            If RptSelCt!ckcSelC10(0).Value = vbChecked Then     'new page per vehicle?
                If Not gSetFormula("NewPage", "'Y'") Then           'cant skip to new page when vehicle subtotals are excluded
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("NewPage", "'N'") Then           'cant skip to new page when vehicle subtotals are excluded
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
        Else
            If Not gSetFormula("IncludeVehicleTotals", "'N'") Then      'sort by sales origin
                mCntJob11Plus = -1
                Exit Function
            End If
            If Not gSetFormula("NewPage", "'N'") Then           'cant skip to new page when vehicle subtotals are excluded
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
    ElseIf ilListIndex = CNT_SALESCOMPARE Then
        ilVehicleGroup = 0              'this will be the index to vehicle group selected, if applicable
        If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then       'not Top down
            'new options for major and minor sorting have different formulas for non-top down reports
            'Top down reports will remain the same
            If RptSelCt!cbcSet1.ListIndex = 0 Then
                If Not gSetFormula("MajorSortBy", "'A'") Then   'advt
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
                If Not gSetFormula("MajorSortBy", "'G'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
                If Not gSetFormula("MajorSortBy", "'B'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
                If Not gSetFormula("MajorSortBy", "'P'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 4 Then          'slsp
                If Not gSetFormula("MajorSortBy", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then            '4-25-06 vehicle option added
                If Not gSetFormula("MajorSortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                            'vehicle group
                If Not gSetFormula("MajorSortBy", "'H'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                illoop = RptSelCt!lbcSelection(12).ListIndex            '3-18-16 chg from lbcselection(4)
                ilVehicleGroup = tgVehicleSets1(illoop).iCode

            End If
            If Not gSetFormula("MajorVehicleGroupHdr", ilVehicleGroup) Then
                mCntJob11Plus = -1
                Exit Function
            End If

            ilNoneExists = True                    'NONE  allowed in this list
            ilMinorGroupHdr = True                 'there is no minor vehicle group hdr to send to crystal for ths report
            If mCBCSet2Test(ilNoneExists, ilMinorGroupHdr) Then
                mCntJob11Plus = -1
                Exit Function
            End If

'            ilVehicleGroup = 0
'            If RptSelCt!cbcSet2.ListIndex = 0 Then              'no minor sort selected
'                If Not gSetFormula("MinorSortBy", "''") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then          'advt
'                If Not gSetFormula("MinorSortBy", "'A'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then          'agency
'                If Not gSetFormula("MinorSortBy", "'G'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then          'bus cat
'                If Not gSetFormula("MinorSortBy", "'B'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then          'prod prot
'                If Not gSetFormula("MinorSortBy", "'P'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 5 Then          'slsp
'                If Not gSetFormula("MinorSortBy", "'S'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then            'vehicle
'                If Not gSetFormula("MinorSortBy", "'V'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'            Else                                            'vehicle group
'                If Not gSetFormula("MinorSortBy", "'H'") Then
'                    mCntJob11Plus = -1
'                    Exit Function
'                End If
'                'get the vehicle group selected for report heading (participant, format, market, etc)
'                ilLoop = RptSelCt!lbcSelection(4).ListIndex
'                ilVehicleGroup = tgVehicleSets1(ilLoop).iCode
'            End If
'            If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If

            'if Advertiser is primary selection, force to show the advt totals; or if the
            'Include advt subtotals is checked because major and minor do not include the advt totals
            'If (RptSelCt!cbcSet1.ListIndex = 0 And RptSelCt!cbcSet1.ListIndex = 0) Or RptSelCt!cbcSet2.ListIndex = 1 Or RptSelCt!ckcSelC13(0).Value = vbChecked Then
            If RptSelCt!ckcSelC13(0).Value = vbChecked Then
                If Not gSetFormula("InclAdvtTotals", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("InclAdvtTotals", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If

            If RptSelCt!ckcSelC13(1).Value = vbChecked Then
                If Not gSetFormula("SeparatePoliticals", "'Y'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SeparatePoliticals", "'N'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
            If mBOBCrystal() < 0 Then                                'send Crystl formulas for Header notations (pkg vs hidden),
                mCntJob11Plus = -1                                 'As of Time,  Gross, Net
                Exit Function
            End If
        Else                        'top down report, retain all original code
            If RptSelCt!cbcSet1.ListIndex = 0 Then
                If Not gSetFormula("SortBy", "'A'") Then   'advt
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
                If Not gSetFormula("SortBy", "'G'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
                If Not gSetFormula("SortBy", "'B'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
                If Not gSetFormula("SortBy", "'P'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 4 Then          'slsp
                If Not gSetFormula("SortBy", "'S'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then            '4-25-06 vehicle option added
                If Not gSetFormula("SortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                            'vehicle group
                If Not gSetFormula("SortBy", "'H'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
                illoop = RptSelCt!lbcSelection(12).ListIndex        '3-18-16 chg from lbcselection(4)
                ilVehicleGroup = tgVehicleSets1(illoop).iCode

            End If

            If mBOBCrystal() < 0 Then                                'send Crystl formulas for Header notations (pkg vs hidden),
                mCntJob11Plus = -1                                 'As of Time,  Gross, Net
                Exit Function
            End If
        End If
        
        '10-13-10 new page each major sort
        If RptSelCt!ckcSelC13(3).Value = vbChecked Then
            If Not gSetFormula("NewPage", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        
        '3-20-18 pass type of calendar selected (std or calendar)
        If RptSelCt!rbcSelC9(0).Value Then          'corporate - unused and hidden for now
            slStr = "O"
        ElseIf RptSelCt!rbcSelC9(1).Value Then      'std
            slStr = "S"
        Else
            slStr = "C"
        End If
        If Not gSetFormula("CalType", "'" & slStr & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        ilRet = mSendBaseCompareToCrystal()        'get the base and compare dates to send to crystal
        If ilRet = -1 Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        'following code made into subroutine (mSendBaseCompareToCrystal
'        slstr = RptSelCt!edcSelCFrom1.Text             'month in text form (jan..dec)
'        gGetMonthNoFromString slstr, ilSaveMonth          'getmonth #
'        If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
'            ilSaveMonth = Val(slstr)
'        End If
'
'        If RptSelCt!rbcSelC9(1).Value Then           '3-21-18 std (implement cal below)
'            'Format the base date Month & year spans to send to Crystal
'            slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(RptSelCt!edcSelCFrom.Text)
'            slEarliest = gObtainEndStd(slEarliest)
'            gObtainYearMonthDayStr slEarliest, True, slSaveYear, slMonth, slDay
'            slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
'            slDate = slDate & " " & right$(Trim$(slSaveYear), 2)    'retrieve the last digits of year (i.e. 97, 98)
'            ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            slstr = slEarliest
'            Do While ilLoop <> 0
'                slLatest = gObtainEndStd(slstr)
'                slstr = gObtainStartStd(slLatest)
'                llDate = gDateValue(slLatest)
'                llDate = llDate + 1
'                slstr = Format$(llDate, "m/d/yy")
'                ilLoop = ilLoop - 1
'            Loop
'            gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
'            slCode = Left$(gMonthName(slLatest), 3)
'            slCode = slCode & " " & right$(Trim$(slYear), 2)
'            If Not gSetFormula("BaseDates", "'" & slDate & "-" & slCode & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'
'            'Format the comparison dates (last year)
'            ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            If RptSelCt!rbcSelC11(1).Value = True Then      '3-23-16 include all last year (vs thru specified month)
'                ilLoop = 12
'                slEarliest = "1/15/" & Trim$(str$(Val(slSaveYear)))
'            Else
'                slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(Val(slSaveYear)))
'            End If
'            'Format the comparison date Month & year spans  to send to crystal
'            slEarliest = gObtainEndStd(slEarliest)  'std end date for start of previous span
'            gObtainYearMonthDayStr slEarliest, True, slYear, slMonth, slDay
'            slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
'            slDate = slDate & " " & Trim$(right$(str$(Val(slSaveYear) - 1), 2))  'retrieve the last digits of year (i.e. 97, 98)
'
'            'ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            slEarliest = slMonth & "/" & "15/" & Trim$(str((Val(slYear) - 1)))  'get previous year
'            slEarliest = gObtainEndStd(slEarliest)
'            slstr = slEarliest
'            Do While ilLoop <> 0
'                slLatest = gObtainEndStd(slstr)
'                slstr = gObtainStartStd(slLatest)
'                llDate = gDateValue(slLatest)
'                llDate = llDate + 1
'                slstr = Format$(llDate, "m/d/yy")
'                ilLoop = ilLoop - 1
'            Loop
'            gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
'            slCode = Left$(gMonthName(slLatest), 3)
'            slCode = slCode & " " & right$(Trim$(slYear), 2)
'            If Not gSetFormula("CompareDates", "'" & slDate & "-" & slCode & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        Else                        '3-20-18 implement calendar month pacing
'            'Format the base date Month & year spans to send to Crystal
'            slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(RptSelCt!edcSelCFrom.Text)
'            slEarliest = gObtainEndCal(slEarliest)
'            gObtainYearMonthDayStr slEarliest, True, slSaveYear, slMonth, slDay
'            slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
'            slDate = slDate & " " & right$(Trim$(slSaveYear), 2)    'retrieve the last digits of year (i.e. 97, 98)
'            ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            slstr = slEarliest
'            Do While ilLoop <> 0
'                slLatest = gObtainEndCal(slstr)
'                slstr = gObtainStartCal(slLatest)
'                llDate = gDateValue(slLatest)
'                llDate = llDate + 1
'                slstr = Format$(llDate, "m/d/yy")
'                ilLoop = ilLoop - 1
'            Loop
'            gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
'            slCode = Left$(gMonthName(slLatest), 3)
'            slCode = slCode & " " & right$(Trim$(slYear), 2)
'            If Not gSetFormula("BaseDates", "'" & slDate & "-" & slCode & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'
'            'Format the comparison dates (last year)
'            ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            If RptSelCt!rbcSelC11(1).Value = True Then      '3-23-16 include all last year (vs thru specified month)
'                ilLoop = 12
'                slEarliest = "1/15/" & Trim$(str$(Val(slSaveYear)))
'            Else
'                slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(Val(slSaveYear)))
'            End If
'            'Format the comparison date Month & year spans  to send to crystal
'            slEarliest = gObtainEndCal(slEarliest)  'cal end date for start of previous span
'            gObtainYearMonthDayStr slEarliest, True, slYear, slMonth, slDay
'            slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
'            slDate = slDate & " " & Trim$(right$(str$(Val(slSaveYear) - 1), 2))  'retrieve the last digits of year (i.e. 97, 98)
'
'            'ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
'            slEarliest = slMonth & "/" & "15/" & Trim$(str((Val(slYear) - 1)))  'get previous year
'            slEarliest = gObtainEndCal(slEarliest)
'            slstr = slEarliest
'            Do While ilLoop <> 0
'                slLatest = gObtainEndCal(slstr)
'                slstr = gObtainStartCal(slLatest)
'                llDate = gDateValue(slLatest)
'                llDate = llDate + 1
'                slstr = Format$(llDate, "m/d/yy")
'                ilLoop = ilLoop - 1
'            Loop
'            gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
'            slCode = Left$(gMonthName(slLatest), 3)
'            slCode = slCode & " " & right$(Trim$(slYear), 2)
'            If Not gSetFormula("CompareDates", "'" & slDate & "-" & slCode & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If
        
        '2-19-16 detrmine this year and last year effective dates if applicable
        If Trim$(RptSelCt!edcText.Text) = "" Then
            slPacingTY = ""
            slPacingLY = ""
        Else
            If RptSelCt!rbcSelC9(1).Value Then          'std
                If RptSelCt!edcText.Text <> "" Then     '9-27-18 test date input validity
                    slDate = RptSelCt!edcText.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCt!edcText.SetFocus
                        Exit Function
                    End If
                End If
                'need to calculate the date of last years effective pacing date; based on the # days difference from start of std year
                slPacingTY = Format$(gDateValue(RptSelCt!edcText.Text), "m/d/yy")
                slPacingTY = gObtainYearStartDate(0, RptSelCt!edcText.Text)      'start of current std year for the pacing date; get the difference of pacing date entered against start date of the std year
                llDate = gDateValue(slPacingTY)
                ilTemp = gDateValue(RptSelCt!edcText.Text) - (gDateValue(slPacingTY))                '# days difference for This year
                'get the year of the pacing date and backup to previous year
                gObtainMonthYear 0, slPacingTY, ilMonth, ilYear
                slPacingLY = gObtainYearStartDate(0, "01/15/" & Trim(str(ilYear - 1)))     'start std date of last year
                ' effective date last year
                slPacingLY = Format$(gDateValue(slPacingLY) + ilTemp, "m/d/yy")
                slPacingTY = Format$(gDateValue(RptSelCt!edcText.Text), "m/d/yy")
            Else                                    '3-20-18 implement pacing calendar
                'need to calculate the date of last years effective pacing date; based on the # days difference from start of cal year
                slPacingTY = Format$(gDateValue(RptSelCt!edcText.Text), "m/d/yy")
                slPacingTY = gObtainYearStartDate(1, RptSelCt!edcText.Text)      'start of current cal year for the pacing date; get the difference of pacing date entered against start date of the std year
                llDate = gDateValue(slPacingTY)
                ilTemp = gDateValue(RptSelCt!edcText.Text) - (gDateValue(slPacingTY))                '# days difference for This year
                'get the year of the pacing date and backup to previous year
                gObtainMonthYear 1, slPacingTY, ilMonth, ilYear
                slPacingLY = gObtainYearStartDate(1, "01/15/" & Trim(str(ilYear - 1)))     'start cal date of last year
                ' effective date last year
                slPacingLY = Format$(gDateValue(slPacingLY) + ilTemp, "m/d/yy")
                slPacingTY = Format$(gDateValue(RptSelCt!edcText.Text), "m/d/yy")
            End If
        End If

        If Not gSetFormula("PacingDateTY", "'" & slPacingTY & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not gSetFormula("PacingDateLY", "'" & slPacingLY & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        slStr = RptSelCt!edcTopHowMany.Text      'validate the number of Top How Many Recs requested
        If slStr <> "" Then
            ilRet = IsNumeric(slStr)
            If Not ilRet Then
                mReset                     'invalid input - not all numeric
                RptSelCt!edcTopHowMany.SetFocus
                Exit Function
            End If
            If Not Val(slStr) > 0 Then
                mReset                     'invalid input - conversion not a positive integer
                RptSelCt!edcSet1.SetFocus
                Exit Function
            End If
        Else
            slStr = "9999"  'default value set high so they get all the records
        End If

        If RptSelCt!ckcSelC10(0).Value = vbChecked Then
            If Not gSetFormula("NumRecsToPrint", Val(slStr)) Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        '10-30-08 use sales source as major sort
        If RptSelCt!ckcSelC13(2).Value = vbChecked Then     'use sales source as major
            If Not gSetFormula("SSAsMajor", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SSAsMajor", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

    End If
slExclude = ""
slInclude = ""
If ilListIndex = CNT_BOB_BYSPOT Then
    'slExclude = ""
    'slInclude = ""
    gIncludeExcludeCkc RptSelCt!ckcSelC6(1), slInclude, slExclude, "AirTime"
    gIncludeExcludeCkc RptSelCt!ckcSelC6(2), slInclude, slExclude, "NTR"
    gIncludeExcludeCkc RptSelCt!ckcSelC6(3), slInclude, slExclude, "HardCost"

    gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Orders"

    gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(5), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(6), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelCt!ckcSelC6(0), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(2), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(3), slInclude, slExclude, "Cancel"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(4), slInclude, slExclude, "Hidden"
    slInclude = "Included: " & slInclude
    slExclude = "Excluded: " & slExclude
End If
If ilListIndex = CNT_BOB_BYCNT Or ilListIndex = CNT_BOB_BYSPOT Then
    'If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    'End If
    'If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    'End If
    'If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    'End If
    'If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    'End If
End If
If ilListIndex = CNT_BOB_BYSPOT Then
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
'Date: 11/13/2019 commented out; duplicate call for CNT_AVG_PRICES
'ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Or (ilListIndex = CNT_AVG_PRICES) Then     'Quarterly avails by min or pct
ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Then     'Quarterly avails by min or pct

    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
ElseIf (ilListIndex = CNT_SALESACTIVITY) Or ilListIndex = CNT_CUMEACTIVITY Or ilListIndex = CNT_DAILY_SALESACTIVITY Then
    If ilListIndex = CNT_DAILY_SALESACTIVITY Then
        If RptSelCt!CSI_CalFrom.Text <> "" Then     'Date: 11/23/2019 added CSI calendar controls for date entries --> edcSelCTo.Text <> "" Then
            slEarliest = RptSelCt!CSI_CalFrom.Text  'edcSelCTo.Text
            If Not gValidDate(slEarliest) Then
                mReset
                RptSelCt!CSI_CalFrom.SetFocus       'edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        If RptSelCt!CSI_CalTo.Text <> "" Then       'Date: 11/23/2019 added CSI calendar controls for date entries --> edcSelCTo1.Text <> "" Then
            'If StrComp(RptSelCt!edcSelCTo1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slLatest = RptSelCt!CSI_CalTo.Text  'edcSelCTo1.Text
                If Not gValidDate(slLatest) Then
                    mReset
                    RptSelCt!CSI_CalTo.SetFocus     'edcSelCTo1.SetFocus
                    Exit Function
                End If
            End If
        End If
        llDate = gDateValue(slEarliest)
        slType = "Activity Dates: " & Format$(llDate, "m/d/yy")
        llDate = gDateValue(slLatest)
        slType = slType & " - " & Format$(llDate, "m/d/yy") & " "
        If Not gSetFormula("EffDate", "'" & slType & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If RptSelCt!rbcSelC4(0).Value Then      'advt sort
            If Not gSetFormula("SortBy", "'A'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                    'sales office sort
            If Not gSetFormula("SortBy", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Not gSetFormula("RptInterval", "'D'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
    End If
    If ilListIndex = CNT_SALESACTIVITY Then
        'Sales Activity by Qtr or Cume Activity
        If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
            mCntJob11Plus = -1
            Exit Function
        End If
        slStr = RptSelCt!CSI_CalFrom.Text       'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
        'insure its a Monday
        llDate = gDateValue(slStr)

        slType = "Activity Dates: " & Format$(llDate, "m/d/yy") & "-"
        slType = slType & Format$(llDate + 6, "m/d/yy") & " "
        If Not gSetFormula("EffDate", "'" & slType & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not gSetFormula("SortBy", "'A'") Then    'force to sort by advt
            mCntJob11Plus = -1
            Exit Function
        End If
        If Not gSetFormula("RptInterval", "'W'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        lgOrigCntrNo = gDateValue(slDate)           'start date of qtr- put in common temporarily
    End If

    If ilListIndex = CNT_SALESACTIVITY Or ilListIndex = CNT_DAILY_SALESACTIVITY Then
        '2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
        If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then   'When Not Separatating Air Time, NTR and HC
            slInclude = ""
            If RptSelCt!ckcSelC13(0).Value = vbChecked Then      'Include AirTime
                slInclude = "Air Time"
            End If
            If RptSelCt!ckcSelC13(1).Value = vbChecked Then      'Include NTR
                If slInclude <> "" Then slInclude = slInclude & ", "
                slInclude = slInclude & "NTR"
            End If
            If RptSelCt!ckcSelC13(2).Value = vbChecked Then      'Include HC
                If slInclude <> "" Then slInclude = slInclude & ", "
                slInclude = slInclude & "HC"
            End If
            If slInclude <> "" Then slInclude = "Includes: " & slInclude
            If Not gSetFormula("RptInclusions", "'" & slInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
    End If


    If ilListIndex = CNT_CUMEACTIVITY Then
        If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
            mCntJob11Plus = -1
            Exit Function
        End If
        slStr = RptSelCt!CSI_CalFrom.Text           'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
        'insure its a Monday
        llDate = gDateValue(slStr)

        slType = "Effective " & Format$(llDate, "m/d/yy") & " "
        If Not gSetFormula("EffDate", "'" & slType & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        lgOrigCntrNo = gDateValue(slDate)           'start date of qtr- put in common temporarily
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
'        If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'            mCntJob11Plus = -1
'            Exit Function
'        End If
        
        
'        If RptSelCt!rbcSelC7(0).Value Then                'Gross
'            If Not gSetFormula("GrossOrNet", "'G'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        Else                                            'for net, if owner option, it's net-net
'            If Not gSetFormula("GrossOrNet", "'N'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        End If
        If RptSelCt!rbcSelCInclude(0).Value Then
            If Not gSetFormula("SortBy", "'A'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
            '1-17-06 add option to show vehicle subtotals rather than always showing them
             If RptSelCt!rbcSelCInclude(0).Value = True Then                  'advt option
                If RptSelCt!rbcSelC4(0).Value And RptSelCt!ckcSelC13(0).Value = vbChecked Then       'detail option, & show vehicle subtotals checked on
                    If Not gSetFormula("ShowVefTots", "'Y'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                ElseIf RptSelCt!rbcSelC4(0).Value And RptSelCt!ckcSelC13(0).Value = vbUnchecked Then       'detail option, & dont show vehicle subtotals checked on
                    If Not gSetFormula("ShowVefTots", "'N'") Then
                        mCntJob11Plus = -1
                        Exit Function
                    End If
                End If
            End If

        ElseIf RptSelCt!rbcSelCInclude(1).Value Then
            If Not gSetFormula("SortBy", "'G'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCt!rbcSelCInclude(2).Value Then
            If Not gSetFormula("SortBy", "'D'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        gIncludeExcludeCkc RptSelCt!ckcSelC12(0), slInclude, slExclude, "New Contracts Only"
        gIncludeExcludeCkc RptSelCt!ckcSelC8(0), slInclude, slExclude, "Air Time"
        gIncludeExcludeCkc RptSelCt!ckcSelC8(1), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelCt!ckcSelC8(2), slInclude, slExclude, "Hard Cost"

        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

    End If
    
    ilRet = mGrossOrNetHdr()            '11-4-13
    If ilRet = -1 Then
        mCntJob11Plus = -1
        Exit Function
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then
    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = RptSelCt!edcSelCTo1.Text             'month in text form (jan..dec)
    gGetMonthNoFromString slStr, igMonthOrQtr          'getmonth #
    If igMonthOrQtr = 0 Then                                 'input isn't text month name, try month #
        igMonthOrQtr = Val(slStr)
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then
            mReset
            RptSelCt!edcSelCTo1.SetFocus                 'invalid # periods
            Exit Function
        End If
    End If
    
    slStr = RptSelCt!edcSelCTo.Text
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelCt!edcSelCTo.SetFocus                 'invalid year
        Exit Function
    End If

    slStr = RptSelCt!edcText.Text            '#periods
    ilMonth = Val(slStr)
    ilRet = gVerifyInt(slStr, 1, 12)
    If ilRet = -1 Then
        mReset
        RptSelCt!edcText.SetFocus
        Exit Function
    End If
    igPeriods = Val(slStr)                      '7-7-14
    If Not gSetFormula("NumberPeriods", ilMonth) Then
        mCntJob11Plus = -1
        Exit Function
    End If


    If ilListIndex = CNT_SALESACTIVITY_SS Then
        If RptSelCt!CSI_CalFrom.Text <> "" Then     'Date: 11/26/2019   added CSI calendar controls for date entries -->  edcSelCFrom.Text <> "" Then
            slEarliest = RptSelCt!CSI_CalFrom.Text  '--> edcSelCFrom.Text
            If Not gValidDate(slEarliest) Then
                mReset
                RptSelCt!CSI_CalFrom.SetFocus       '--> edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        If RptSelCt!CSI_CalTo.Text <> "" Then       'Date: 11/26/2019   added CSI calendar controls for date entries --> edcSelCFrom1.Text <> "" Then
            'If StrComp(RptSelCt!edcSelCFrom1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slLatest = RptSelCt!CSI_CalTo.Text  '--> edcSelCFrom1.Text
                If Not gValidDate(slLatest) Then
                    mReset
                    RptSelCt!CSI_CalTo.SetFocus     '--> edcSelCFrom1.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    

    If ilListIndex = CNT_SALESPLACEMENT Then
        'send the report month headings
'        ilMonth = ilSaveMonth
'        For ilLoop = 1 To 12
'            ilMonth = (ilSaveMonth + ilLoop) - 1
'            If ilMonth > 12 Then
'                ilMonth = ilMonth - 12
'            End If
'            slStr = Mid$(slMonthHdr, (ilMonth - 1) * 3 + 1, 3)
'            If ilLoop = 1 Then          'save for heading to pass to crystal
'                slMonth = slStr
'            End If
'            If Not gSetFormula("P" & Trim$(str$(ilLoop)), "'" & slStr & "'") Then
'                mCntJob11Plus = -1
'                Exit Function
'            End If
'        Next ilLoop
        
'        slType = "for Std " & slMonth & " " & RptSelCt!edcSelCTo.Text
        slType = gGetMonthFromInx(RptSelCt!rbcSelCInclude(0), RptSelCt!rbcSelCInclude(1), RptSelCt!rbcSelCInclude(2))
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
        llDate = gDateValue(slStr)
        slType = slType & " Last Date Billed: " & Format$(llDate, "m/d/yy")
    Else
        slMonth = gGetMonthFromInx(RptSelCt!rbcSelCInclude(0), RptSelCt!rbcSelCInclude(1), RptSelCt!rbcSelCInclude(2))
        llDate = gDateValue(slEarliest)
        If RptSelCt!rbcSelCSelect(0).Value Then
        slType = "Original Entry Dates: " & Format$(llDate, "m/d/yy")
        Else
            slType = "Latest Mod Dates: " & Format$(llDate, "m/d/yy")  '9-11-02 changed from Activity DAtes to Latest Mod Date
        End If
        llDate = gDateValue(slLatest)
        'slType = slType & " - " & Format$(llDate, "m/d/yy") & " for Std " & slMonth & " " & RptSelCt!edcSelCTo.Text
        slType = slType & " - " & Format$(llDate, "m/d/yy") & " for " & slMonth
    End If

    If Not gSetFormula("EffDate", "'" & slType & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If

    'send formula to crystal for report header
    ilRet = mGrossOrNetHdr()            '11-4-13
    If ilRet = -1 Then
        mCntJob11Plus = -1
        Exit Function
    End If

'    If RptSelCt!rbcSelC7(0).Value Then                'Gross
'        If Not gSetFormula("GrossOrNet", "'G'") Then
'            mCntJob11Plus = -1
'            Exit Function
'        End If
'    ElseIf RptSelCt!rbcSelC7(1).Value Then                   'for net, if owner option, it's net-net
'        If Not gSetFormula("GrossOrNet", "'N'") Then
'            mCntJob11Plus = -1
'            Exit Function
'        End If
'    Else                                                    'net-net
'        If Not gSetFormula("GrossOrNet", "'D'") Then
'            mCntJob11Plus = -1
'            Exit Function
'        End If
'    End If

    'send formula to crystal for Detail/Summary
    If RptSelCt!rbcSelC11(0).Value Then                'Detail
        If Not gSetFormula("DetailorSummary", "'D'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else                                            'Sumary
        If Not gSetFormula("DetailOrSummary", "'S'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If


    'skip to new page for each new group
    If RptSelCt!ckcSelC12(0).Value = vbChecked Then         'yes, skip to page each major group
        If Not gSetFormula("SkipPage", "'Y'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("SkipPage", "'N'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If

    If ilListIndex = CNT_SALESPLACEMENT Then
        If Not gSetFormula("WhichReport", "'P'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

        If Not gSetFormula("Included", "' '") Then      'no air time/ntr/hard cost option in this report, send blanks for report inclusion info
            mCntJob11Plus = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("WhichReport", "'A'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

        gIncludeExcludeCkc RptSelCt!ckcSelC13(0), slInclude, slExclude, "Air Time"
        gIncludeExcludeCkc RptSelCt!ckcSelC13(1), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelCt!ckcSelC13(2), slInclude, slExclude, "Hard Cost"
        gIncludeExcludeCkc RptSelCt!ckcSelC12(2), slInclude, slExclude, "Split Slsp"
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

    End If

    '2-18-03 Send to Crystal flag to include slsp subtotals
    If RptSelCt!rbcSelC9(2).Value Then      'option by advt never needs slsp subtotals
        If Not gSetFormula("IncludeSlsp", "'N'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else
        If RptSelCt!ckcSelC12(1).Value = vbChecked Then     'include slsp subtotals
            If Not gSetFormula("IncludeSlsp", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("IncludeSlsp", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
    End If
    ilErr = gGRFSelection(slSelection)      'build date & time filter to send to crystal
    If ilErr <> 0 Then
        mCntJob11Plus = -1
        Exit Function
    End If
ElseIf ilListIndex = CNT_MAKEPLAN Then
    slStr = RptSelCt!edcSelCFrom.Text
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelCt!edcSelCFrom.SetFocus                 'invalid year
        Exit Function
    End If
    slStr = RptSelCt!edcSelCFrom1.Text
    ilRet = gVerifyInt(slStr, 1, 4)
    If ilRet = -1 Then
        mReset
        RptSelCt!edcSelCFrom1.SetFocus                 'invalid qtr
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                           'put qtr in global variable
    If gGetQtrHeader(igYear, igMonthOrQtr) <> 0 Then        'get qtr & year text for Crystal   header
        mCntJob11Plus = -1
    End If
    slStr = RptSelCt!edcSelCTo.Text
    If Val(slStr) > 4 Then
        mReset
        RptSelCt!edcSelCTo.SetFocus                 'invalid # qtr
        Exit Function
    End If
    If RptSelCt!rbcSelC4(1).Value Then                  'send to the quarterly report how many quarters reporting
        If Val(RptSelCt!edcSelCFrom1.Text) + Val(RptSelCt!edcSelCTo.Text) > 5 Then  '4-24-01 prevent going past the end of the year
            mReset
            RptSelCt!edcSelCTo.SetFocus                 'invalid # qtr
            Exit Function
        Else
            slStr = RptSelCt!edcSelCTo.Text
            If Not gSetFormula("NoQtrs", "'" & slStr & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
    End If

    If RptSelCt!rbcSelCSelect(0).Value Then     'corp, more than 1 r/c is necessary
        ilMonth = 11
        If Not gSetFormula("CorpOrStd", "'" & "Corp" & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else
        ilMonth = 12
        If Not gSetFormula("CorpOrStd", "'" & "Std" & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If
    slStr = ""
    For illoop = 0 To RptSelCt!lbcSelection(ilMonth).ListCount - 1 Step 1
        If RptSelCt!lbcSelection(ilMonth).Selected(illoop) Then
            slNameCode = tgRateCardCode(illoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
            If slStr = "" Then
                slStr = slNameCode
            Else                                '
                slStr = slStr & "," & slNameCode
            End If
        End If
    Next illoop
    'slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
    'ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
    If Not gSetFormula("RCName", "'" & slStr & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If

ElseIf ilListIndex = CNT_VEHCPPCPM Then         'verify date entered
    slStr = RptSelCt!CSI_CalFrom.Text           'Date: 12/13/2019 added CSI calendar control for date entry --> edcSelCFrom.Text                 'edit effective date
    If Not gValidDate(slStr) Then
        mReset
        RptSelCt!CSI_CalFrom.SetFocus           'Date: 12/13/2019 added CSI calendar control for date entry --> edcSelCFrom.SetFocus
        Exit Function
    End If
    'Send cpp or cpm formula to crystal
    If RptSelCt!rbcSelC4(0).Value Then                'cpp
        If Not gSetFormula("Cpp", "'P'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else                                            'cpm
        If Not gSetFormula("Cpp", "'M'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If
    
    slStr = "G"
    If RptSelCt!rbcSelC11(1).Value Then
        slStr = "N"
    End If
    If Not gSetFormula("GrossNet", "'" & Trim$(slStr) & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If

ElseIf ilListIndex = CNT_SALESANALYSIS Then
    slDate = RptSelCt!CSI_CalFrom.Text              'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalFrom.SetFocus               'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom.SetFocus
        Exit Function
    End If
    slStr = RptSelCt!edcSelCTo.Text                 'entered year
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelCt!edcSelCTo.SetFocus                 'invalid year
        mCntJob11Plus = -1
        Exit Function
    End If
    slStr = RptSelCt!edcSelCTo1.Text                  'edit qtr
    ilRet = gVerifyInt(slStr, 1, 4)
    If ilRet = -1 Then
        mReset
        RptSelCt!edcSelCTo1.SetFocus                 'invalid qtr
        mCntJob11Plus = -1
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
    If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
        mCntJob11Plus = -1
        Exit Function
    End If
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
       'dan M 7-30-08 NTR/Hard Cost option
    slStr = ""
    If RptSelCt!ckcSelC3(0).Value = 1 Or RptSelCt!ckcSelC3(1).Value = 1 Then
        slStr = slStr & "'With"
        If RptSelCt!ckcSelC3(0).Value = 1 Then
            slStr = slStr & " NTR"
            If RptSelCt!ckcSelC3(1).Value = 1 Then
                slStr = slStr & " and Hard Cost"
            End If
        Else
            slStr = slStr & " Hard Cost"
        End If
    End If
    If slStr <> "" Then
        slStr = slStr & "'"
    End If
    If Not gSetFormula("reporttitle", slStr) Then
        mCntJob11Plus = -1
        Exit Function
    End If
End If
If ilListIndex = CNT_MAKEPLAN Or ilListIndex = CNT_VEHCPPCPM Then
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{ANR_Analysis_Report.anrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({ANR_Analysis_Report.anrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
'    If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'        mCntJob11Plus = -1
'        Exit Function
'    End If
End If

If ilListIndex = CNT_VEH_UNITCOUNT Then
    slDate = RptSelCt!CSI_CalFrom.Text      'Date: added CSI calendar control for date entries --> edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalFrom.SetFocus       'Date: added CSI calendar control for date entries --> edcSelCFrom.SetFocus
        Exit Function
    End If

    slDate = RptSelCt!CSI_CalTo.Text        'Date: added CSI calendar control for date entries --> edcSelCFrom1.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalTo.SetFocus         'Date: added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
        Exit Function
    End If

    slStr = RptSelCt!CSI_CalFrom.Text       'Date: added CSI calendar control for date entries --> edcSelCFrom.Text
    llDate = gDateValue(slStr)
    slStr = RptSelCt!CSI_CalTo.Text         'Date: added CSI calendar control for date entries --> edcSelCFrom1.Text
    llDate2 = gDateValue(slStr)
    If (llDate2 - llDate) + 1 > 7 Or (llDate2 - llDate) + 1 < 0 Then        '3-15-06
        'disallow more than 7 days
        MsgBox "Maximum 7 days exceeded or invalid dates"
        mReset
        RptSelCt!CSI_CalTo.SetFocus         'Date: added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
        Exit Function
    End If
    slStr = Format$(llDate, "m/d/yy") & " - " & Format$(llDate2, "m/d/yy")
    If Not gSetFormula("ReportDates", "'" & slStr & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If

    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Manual"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Web"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(2), slInclude, slExclude, "Marketron"
    If Not gSetFormula("Included", "'" & slInclude & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
ElseIf ilListIndex = CNT_LOCKED Then               '4-5-06
    'Date: 12/5/2019 added CSI calendar control for date entries
    slDate = RptSelCt!CSI_CalFrom.Text  'edcSelCFrom.Text
    If Not gValidDate(slDate) Or gWeekDayStr(slDate) <> 0 Then
        mReset
        MsgBox "Enter a valid Monday start date"
        RptSelCt!CSI_CalFrom.SetFocus   'edcSelCFrom.SetFocus
        Exit Function
    End If
    llDate = gDateValue(slDate)
    slStr = "for " & Format$(llDate, "m/d/yy") & " - "

    slDate = RptSelCt!CSI_CalTo.Text    'edcSelCFrom1.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalTo.SetFocus     'edcSelCFrom1.SetFocus
        Exit Function
    End If
    llDate = gDateValue(slDate)
    slStr = slStr & Format$(llDate, "m/d/yy")
    If Not gSetFormula("DatesRequested", "'" & slStr & "'") Then
        mCntJob11Plus = -1
        Exit Function
    End If
    If RptSelCt!rbcSelC4(0).Value Then          'sort by vehicle
        If Not gSetFormula("SortBy", "'V'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else                                        'sort by date
        If Not gSetFormula("SortBy", "'D'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
ElseIf ilListIndex = CNT_GAMESUMMARY Then               '7-14-06
    slDate = RptSelCt!CSI_CalFrom.Text  'Date: 12/4/2019 added CSI calendar control for date entries -->  edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalFrom.SetFocus   'edcSelCFrom.SetFocus
        Exit Function
    End If
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay

    slDate = RptSelCt!CSI_CalTo.Text    'Date: 12/4/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCt!CSI_CalTo.SetFocus     'edcSelCFrom1.SetFocus
        Exit Function
    End If
    gObtainYearMonthDayStr slDate, True, slYear2, slMonth2, slDay2

    If RptSelCt!rbcSelC4(0).Value Then          'sort by vehicle
        If Not gSetFormula("SortBy", "'V'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else                                        'sort by date
        If Not gSetFormula("SortBy", "'D'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If
    
    '3-9-11 Show Live Log flag for each game if applicable
    If (Asc(tgSpf.sUsingFeatures) And USINGLIVELOG) = USINGLIVELOG Then
        If Not gSetFormula("UsingLiveLog", "'Y'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("UsingLiveLog", "'N'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
    End If



    'Send selected vehicle(s) to Crystal
    slSelection = ""
    slStr = "({GSF_Game_Schd.gsfAirDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") And {GSF_Game_Schd.gsfAirDate} <= Date(" & slYear2 & "," & slMonth2 & "," & slDay2 & ") )"
    If RptSelCt!ckcSelC10(0).Value = 0 Then                 'exclude cancelled games 5/13/2008
     slStr = slStr & " and ({GSF_Game_Schd.gsfGameStatus} <>'C')"
    End If
    If Not (RptSelCt!ckcAll.Value = vbChecked) Then         'selective vehicles
        For illoop = 0 To RptSelCt!lbcSelection(3).ListCount - 1 Step 1
            If RptSelCt!lbcSelection(3).Selected(illoop) Then
                slNameCode = tgVehicle(illoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                slSelection = slSelection & slOr & "{VEF_Vehicles.vefCode} = " & Trim$(slCode)
                slOr = " Or "
            End If
        Next illoop
    End If

    If slSelection = "" Then        'get all vehicles, nothing built in selection formula
        slSelection = Trim$(slStr)
    Else                            'selective vehicles, concatenate with the dates formula
        slSelection = "(" & slSelection & ") and " & slStr
    End If
    If Not gSetSelection(slSelection) Then
        mCntJob11Plus = -1
        Exit Function
    End If
End If

    mCntJob11Plus = 1               'return, all OK
    Exit Function
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
    RptSelCt!frcOutput.Enabled = igOutput
    RptSelCt!frcCopies.Enabled = igCopies
    'RptSelCt!frcWhen.Enabled = igWhen
    RptSelCt!frcFile.Enabled = igFile
    RptSelCt!frcOption.Enabled = igOption
    'RptSelCt!frcRptType.Enabled = igReportType
    Beep
End Sub
'
'
'            mSendCorpStd - Send to Crystal the formula
'               indicating whether the report is run
'               Corp or Std (this string should not
'               be combined with WeekQtrHdr because
'               many of the Crystal reports test the 1st
'               character for 1-4 and determines other report
'               headings.
'
'               <input> Value for Corp (true if corp)
'                       Value for Std (true if std)
'               <output> None, formula sent to Crystal
'
'
Function mSendCorpStd(rbcCorp As Control, rbcStd As Control) As Integer
Dim slStr As String
    mSendCorpStd = True
    If rbcCorp Then
        slStr = "Corp"
    Else
        slStr = "Std"
    End If
    If Not gSetFormula("CorpStd", "'" & slStr & "'") Then
        mSendCorpStd = False
        Exit Function
    End If
End Function
'*******************************************************************
'
'
'               Function mVerifyCntr - verify if all numeric entered
'
'               <input>  slInput   - input string
'               <return>  false if invalid
'
'
'*******************************************************************
'
Function mVerifyNumber(slInput As String) As Integer
    If slInput <> "" Then
        If IsNumeric(slInput) Then
            mVerifyNumber = True                 'valid numeric entered
        Else
            mVerifyNumber = False
        End If
    Else
        mVerifyNumber = True
    End If
End Function
' *********************************************************************
'
'            mWeekQtrHdr()
'           Setup Week and Quarter Header and send to Crystal reports
'           as a formula named WeekQtrHeader
'           Format output as Quarter X Year XXXX
'           <output> slDate - First date of quarter
'           Created:  7/3/96
'
'***********************************************************************
Function mWeekQtrHdr(slDate As String) As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slStr As String

    mWeekQtrHdr = True
    ilYear = RptSelCt!edcSelCTo.Text                'starting year
    If ilYear < 100 Then           'only 2 digit year input ie.  96, 95,
        If ilYear < 50 Then        'adjust for year 1900 or 2000
            ilYear = 2000 + ilYear
        Else
            ilYear = 1900 + ilYear
        End If
    End If

    ilMonth = RptSelCt!edcSelCTo1.Text              'month
    slDate = Trim$(str$(((ilMonth - 1) * 3 + 1))) & "/15/" & Trim$(str$(ilYear))
    slDate = gObtainStartStd(slDate)
    If ilMonth = 1 Then
        slStr = "1st"
    ElseIf ilMonth = 2 Then
        slStr = "2nd"
    ElseIf ilMonth = 3 Then
        slStr = "3rd"
    Else
        slStr = "4th"
    End If
    slStr = slStr & " Qtr" & str$(ilYear)      'add Year
     'Crystal tests this formula for 1, 2, 3, or 4 and uses it to calculate
     'other report headers, etc.
     If Not gSetFormula("WeekQtrHeader", "'" & slStr & "'") Then
         mWeekQtrHdr = False
         Exit Function
     End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportCt                    *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*      11-30-04 allow Sales Commission to be requested*
'*          for more than 1 month
'*******************************************************
Function gCmcGenCt(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'   ilRet = gCmcGenCt(ilListIndex)
'
'   ilRet (O)-  -1= Terminate, error in crystal gsetselectio or gsetformula
'               0 = Crystal input error
'               1 = successful crystal report
'               2 = successful bridge report
'
    Dim illoop As Integer
    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slStatus As String
    Dim slTime As String
    Dim slText As String
    Dim slSortStr As String
    Dim ilPos1 As Integer
    Dim ilYear As Integer

    gCmcGenCt = 0
    Select Case igRptCallType
        Case SLSPCOMMSJOB
            If ilListIndex = COMM_SALESCOMM Then
                slText = RptSelCt!edcSelCFrom.Text                 'pass start month and year for header
                slText = UCase$(Left$(slText, 1))
                slStr = slText & Mid$(RptSelCt!edcSelCFrom.Text, 2)
                slYear = RptSelCt!edcSelCFrom1.Text
                ilYear = Val(slYear)

                '11-30-04 determine if multiple months entered
                slText = RptSelCt!edcSelCTo.Text       'force to 1 month if invalid
                If Val(slText) = 0 Then
                    slText = "1"
                End If

                If Val(slText) <> 1 Then
                    slMonth = "JanFebMarAprMayJunJulAugSepOctNovDec"
                    ilPos1 = InStr(1, slMonth, Trim$(slStr))        'find out the month
                    ilPos1 = ilPos1 / 3 + 1         'get month index
                    ilPos1 = ilPos1 + Val(slText) - 1     'add the number of months
                    If ilPos1 > 12 Then
                        ilPos1 = ilPos1 - 12
                        ilYear = ilYear + 1     'year wraparound
                    End If
                    slStr = slStr & "-"
                    slDay = Mid$(slMonth, (ilPos1 - 1) * 3 + 1, 3) 'get the ending month
                    slStr = slStr & Trim$(slDay)
                End If


                'slStr = slStr & " " & RptSelCt!edcSelCFrom1.Text        'year
                slStr = slStr & " " & str$(ilYear)
                If Not gSetFormula("MonthHdr", "'" & slStr & "'") Then
                    gCmcGenCt = -1
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCFrom.Text             'month in text form (jan..dec)
                gGetMonthNoFromString slStr, ilIndex        'getmonth #
                slStr = Trim$(str$(ilIndex)) & "/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay  'pass 1st date of the current month


                slStr = RptSelCt!edcSelCFrom.Text             'month in text form (jan..dec)
                gGetMonthNoFromString slStr, ilIndex        'getmonth #
                slStr = Trim$(str$(ilIndex)) & "/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)
                slStr = gObtainStartStd(slStr)   '4-20-00 chged to use start of bdcst date instead of end of bdcst month
                                                 'any adjustments entered using the middle of the month were not included as current
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay  'pass 1st date of the current month
                If Not gSetFormula("CommCalc", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenCt = -1
                    Exit Function
                End If


                If RptSelCt!ckcSelC3(0).Value = vbUnchecked Then        'not Bonus commission version with new & increased sales

                    '2-12-02 Subtotals by contrct
                    'If RptSelCt!rbcSelCSelect(0).Value = True Then                'detail (vs sumary)
                    If RptSelCt!rbcSelC7(0).Value = True Then                       '1-10-07 change control for detail vs summary
                        If RptSelCt!ckcSelC8(0).Value = vbChecked Then          'detail, Sub-totals by contr?
                            If Not gSetFormula("ContrTots", "'Y'") Then
                                gCmcGenCt = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("ContrTots", "'N'") Then
                                gCmcGenCt = -1
                                Exit Function
                            End If
                        End If
                    Else            'summary
                        'If (RptSelCt!rbcSelCSelect(1).Value = True And RptSelCt!ckcSelC8(0).Value = True) Then          'if summary and want subtotals by cntr, sende formula to indicate
                        If (RptSelCt!rbcSelC7(1).Value = True And RptSelCt!ckcSelC8(0).Value = True) Then          'if summary and want subtotals by cntr, sende formula to indicate
                            slStr = "S"             'assume to show cnt subtotals
                            If Not RptSelCt!ckcSelC8(0).Value = vbChecked Then
                                slStr = "D"
                            End If
                            If Not gSetFormula("DetOrSum", "'" & slStr & "'") Then         'show all vehicles for contr
                                gCmcGenCt = -1
                                Exit Function
                            End If

                        End If

                        'one summary version, let crystal know if vehicle or slsp major sort
                        If RptSelCt!rbcSelCSelect(0).Value Then     'slsp major sort
                            If Not gSetFormula("MajorByVehicleOrSlsp", "'S'") Then         'major sort by slsp
                                gCmcGenCt = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("MajorByVehicleOrSlsp", "'V'") Then         'major sort by vehicle
                                gCmcGenCt = -1
                                Exit Function
                            End If
                        End If
                    End If


                    If RptSelCt!rbcSelC11(0).Value = True Then     'by advt
                        If Not gSetFormula("SortBy", "'A'") Then         'show all vehicles for contr
                            gCmcGenCt = -1
                            Exit Function
                        End If
                    ElseIf RptSelCt!rbcSelC11(1).Value = True Then      'pct
                        If Not gSetFormula("SortBy", "'P'") Then
                            gCmcGenCt = -1
                            Exit Function
                        End If
                    Else            'owner
                        If Not gSetFormula("SortBy", "'O'") Then        'owner
                            gCmcGenCt = -1
                            Exit Function
                        End If
                    End If
                Else                    'bonus version
                    If (RptSelCt!rbcSelC7(1).Value = True) Then          ' summary
                        slStr = "S"             'assume to show cnt subtotals
                    Else
                        slStr = "D"
                    End If
                    If Not gSetFormula("DetOrSum", "'" & slStr & "'") Then         'show all vehicles for contr
                        gCmcGenCt = -1
                        Exit Function
                    End If
                End If

                 'test for airtime, ntr or both
                If RptSelCt!rbcSelC9(0).Value Then        'air time only
                    slStatus = "Air Time Only"
                ElseIf RptSelCt!rbcSelC9(1).Value Then    'ntr only
                    If RptSelCt!ckcSelC10(0).Value = vbChecked Then
                        slStatus = "NTR & Hard Cost Only"
                    Else
                        slStatus = "NTR Only (excl Hard Cost)"
                    End If
                Else        'both air time & NTR included, is Hard cost included?
                    If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then
                        slStatus = "Air Time & NTR (excl Hard Cost)"
                    Else
                        slStatus = "Air Time, NTR, Hard Cost"
                    End If
                End If
                
'                If Not gSetFormula("AirTimeNTRHdr", "'" & slStatus & "'") Then
'                    gCmcGenCt = -1
'                    Exit Function
'                End If
                
                '4-14-015 test for political, non-polit
                If RptSelCt!ckcSelC13(0).Value = vbChecked And RptSelCt!ckcSelC13(1).Value = vbUnchecked Then        'polit only
                    slStatus = slStatus & ", Political Only"
                ElseIf RptSelCt!ckcSelC13(0).Value = vbUnchecked And RptSelCt!ckcSelC13(1).Value = vbChecked Then        'non-polit only
                    slStatus = slStatus & ", Non-Political Only"
                Else
                    slStatus = slStatus & ", Political & Non-Political"
                End If
                
                If Not gSetFormula("AirTimeNTRHdr", "'" & slStatus & "'") Then      'append Polit/non-polit to airtime/NTR header info
                    gCmcGenCt = -1
                    Exit Function
                End If
                


                '7-16-08 dont show acqusition header on report if excluding them from report
                If RptSelCt!ckcSelC12(0).Value = vbChecked Then
                    If Not gSetFormula("ShowAcq", "'Y'") Then
                        gCmcGenCt = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowAcq", "'N'") Then
                        gCmcGenCt = -1
                        Exit Function
                    End If
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGenCt = -1
                    Exit Function
                End If

            ElseIf ilListIndex = COMM_PROJECTION Then
                If RptSelCt!rbcSelCSelect(0).Value Then           'use package lines
                    slSortStr = "Use Package lines"
                Else
                    slSortStr = "Use Airing lines"
                End If
                If tgSpf.sInvAirOrder <> "S" Then            'bill as ordered, update as ordered, no adjustments at all
                    If RptSelCt!ckcSelC8(0).Value = vbChecked Then                'subt misses
                        slSortStr = slSortStr & "; for standard lines subtract misses"
                    End If                                              'show nothing if ignoring them
                    'Else
                    '    slReserve = slReserve & "exclude misses and makegoods"
                    'End If
                    If RptSelCt!ckcSelC8(1).Value = vbChecked Then                  'count mg when they air?
                        slSortStr = slSortStr & "; for standard lines count MGs"
                    End If
                End If
                If Not gSetFormula("Adjustments", "'" & slSortStr & "'") Then
                    gCmcGenCt = -1
                    Exit Function
                End If

                gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                If Not gSetFormula("LastBilled", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenCt = -1
                    Exit Function
                End If
                slStr = ""
                illoop = Val(RptSelCt!edcSelCFrom.Text)
                If illoop = 1 Then
                    slStr = "1st Quarter "
                ElseIf illoop = 2 Then
                    slStr = "2nd Quarter "
                ElseIf illoop = 3 Then
                    slStr = "3rd Quarter "
                Else
                    slStr = "4th Quarter "
                End If
                slStr = slStr & RptSelCt!edcSelCFrom1.Text
                If Not gSetFormula("QtrHeader", "'" & slStr & "'") Then
                    gCmcGenCt = -1
                    Exit Function
                End If
                If RptSelCt!ckcSelC8(2).Value = vbChecked Then          'skip to new page each slsp
                    If Not gSetFormula("Skip", "'Y'") Then
                        gCmcGenCt = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("Skip", "'N'") Then
                        gCmcGenCt = -1
                        Exit Function
                    End If
                End If

                gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGenCt = -1
                    Exit Function
                End If
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'                gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime        'run time to show on report
'                If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGenCt = -1
'                    Exit Function
'                End If
            End If
        Case CONTRACTSJOB
            'If (igRptType = 0) And (ilListIndex > 1) Then
            '    ilListIndex = ilListIndex + 1
            'End If
            If (ilListIndex < 11) Or (ilListIndex = CNT_INSERTION) Then
                ilRet = mCntJob1_10(ilListIndex, slLogUserCode)
            ElseIf ilListIndex >= 11 And ilListIndex < 38 Then
                If RptSelCt!rbcOutput(4).Value = False Then
                    ilRet = mCntJob11Plus(ilListIndex, slLogUserCode)
                Else
                    ilRet = 1
                End If
            Else
                ilRet = mCntJob38Plus(ilListIndex, slLogUserCode)
            End If
            If ilRet = -1 Then
                gCmcGenCt = -1
                Exit Function
            ElseIf ilRet = 0 Then
                gCmcGenCt = 0
                Exit Function
            ElseIf ilRet = 2 Then
                gCmcGenCt = 2
                Exit Function
            End If
    End Select
    gCmcGenCt = 1
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportCt                      *
'*                                                     *
'*            Created:6/16/93       By:D. LeVine       *
'*            Modified:             By:D. Smith        *
'*                                                     *
'*            D.S. 07/19/00                            *
'*            Added Top Down Reports to the Sales      *
'*            Comparison section                       *
'*                                                     *
'*            D.S. 09/02/00                            *
'*            Fixed formula error when Top Down is     *
'*            used.
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportCt() As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilListIndex As Integer
    Dim slStr As String
    Dim slType As String
    Dim slTemp As String
    Dim blCustomizeLogo As Boolean

    blCustomizeLogo = False                             'most reports will use the general client logo (rptlogo), while proposals/contracts will use custom logos defined with the sales source
                                                        'when opening a crystal report, the default, if not passed or specified, is not to process custom logos defined with sales source
    ilListIndex = RptSelCt!lbcRptType.ListIndex
    Select Case igRptCallType
        Case SLSPCOMMSJOB
            If ilListIndex = COMM_SALESCOMM Then
                'validity check the input dates
                slStr = RptSelCt!edcSelCFrom.Text                 'month text (jan, feb...)
                gGetMonthNoFromString slStr, ilIndex
                If ilIndex = 0 Then
                    mReset
                    RptSelCt!edcSelCFrom.SetFocus                 'invalid month
                    gGenReportCt = False
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCFrom1.Text
                ilRet = gVerifyYear(slStr)
                'If Val(slStr) < 1990 Or Val(slStr) > 2020 Then
                If ilRet = 0 Then
                    mReset
                    RptSelCt!edcSelCFrom1.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If

                slStr = RptSelCt!edcSelCTo.Text            'get # of months
                ilRet = gVerifyInt(slStr, 1, 12)
                If ilRet = -1 Then                          'bad conversion or illegal #
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid # months
                    gGenReportCt = False
                    Exit Function
                End If

                If RptSelCt!ckcSelC3(0).Value = vbChecked Then
                    If Not gOpenPrtJob("commbonus.rpt") Then     'bonus version for new and increased sales
                        gGenReportCt = False
                        Exit Function
                    End If
                Else


                'If (RptSelCt!rbcSelCSelect(0).Value) Then         'detail
                '2-12-02  option to have subtotals by contract


                '5-17-04 Sales Commission prepass & crystal reports re-written for new options.  Old crystal
                'report will be obsolete (commdet.rpt)
                'If RptSelCt!rbcSelC11(0).Value Then             '5-11-04 advt sort within slsp
                '    If (RptSelCt!rbcSelCSelect(0).Value) Or (RptSelCt!rbcSelCSelect(1).Value And RptSelCt!ckcSelC8(0).Value = True) Then         'detail, or summary with contract sub-totals
                '        If Not gOpenPrtJob("commdet.rpt") Then
                '            gGenReportCt = False
                '            Exit Function
                '        End If
                '    Else
                '        If Not gOpenPrtJob("commsum.rpt") Then      'summary
                '            gGenReportCt = False
                '            Exit Function
                '        End If
                '    End If
                'Else                    '5-11-04 summary option uses the same .rpt for all sort versions
                    If (RptSelCt!rbcSelC7(0).Value) Or (RptSelCt!rbcSelC7(1).Value And RptSelCt!ckcSelC8(0).Value = True) Then         'detail, or summary with contract sub-totals
                        If RptSelCt!rbcSelCSelect(0).Value Then         'major sort by slsp
                            If Not gOpenPrtJob("comdetpc.rpt") Then     'detail by pct or vehicle group
                                gGenReportCt = False
                                Exit Function
                            End If
                        Else
                            If Not gOpenPrtJob("comdetvh.rpt") Then     'major sort by vehicle & slsp (detail)
                                gGenReportCt = False
                                Exit Function
                            End If
                        End If
                    Else
                        'If RptSelCt!rbcSelCSelect(0).Value Then      'major sort by slsp (vs vehicle)
                            If Not gOpenPrtJob("commsum.rpt") Then      'summary for vehicle and slsp major sorts
                                gGenReportCt = False
                                Exit Function
                            End If
                        'Else
                        '    If Not gOpenPrtJob("commsmvh.rpt") Then      'summary for vehicle and slsp major sorts
                        ''        gGenReportCt = False
                        '        Exit Function
                        '    End If
                        'End If
                    End If
                End If
                'End If
            ElseIf ilListIndex = COMM_PROJECTION Then
                slStr = RptSelCt!edcSelCFrom1.Text
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSelCt!edcSelCFrom1.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If
                'igYear = Val(RptSel!edcSelCFrom1.Text)
                slStr = RptSelCt!edcSelCFrom.Text
                ilRet = gVerifyInt(slStr, 1, 4)
                If ilRet = -1 Then
                    mReset
                    RptSelCt!edcSelCFrom.SetFocus                 'invalid qtr
                    gGenReportCt = False
                    Exit Function
                End If
                igMonthOrQtr = Val(slStr)                           'put qtr in global variable
                If Not gOpenPrtJob("Commproj.Rpt") Then     'summary only, no detail
                    gGenReportCt = False
                    Exit Function
                End If
            End If
        Case CONTRACTSJOB
            If Not igUsingCrystal Then
                gGenReportCt = True
                Exit Function
            End If
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            'If rbcRptType(0).Value Then
            'remove Spots by Advt & Spots by Date & Time code (see rptselcb)
            If ilListIndex = CNT_BR Then
            'BRs
                blCustomizeLogo = True
                '11-10-03 remove portrait contract (no client uses)
                'If Not RptSelCt!rbcSelCInclude(2).Value Then      'proposals or wide contract                     'proposals/BR contract
                igBRSumZer = False                              'summary - monthly billing
                igBRSum = False                                 'summary with research
                    If igJobRptNo = 1 Then                      'detail pass
                        If RptSelCt!ckcSelC6(1).Value = vbChecked Then        'If including research, show wide with everything if including rates
                            If RptSelCt!ckcSelC6(0).Value = vbChecked Then      'with rates
                                If Not gOpenPrtJob("BR.Rpt", , blCustomizeLogo) Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("BRNoRate.Rpt", , blCustomizeLogo) Then 'with research but without CPP/CPM, and any other rates
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        Else                                    'exclude research, prices optional

                            If Not gOpenPrtJob("BRZer.Rpt", , blCustomizeLogo) Then
                                gGenReportCt = False
                                Exit Function
                            End If
                        End If
                    ElseIf igJobRptNo = 2 Then          'summary pass
                        If Not gOpenPrtJob("BRNTR.Rpt", , blCustomizeLogo) Then 'exclude research, prices optional
                            gGenReportCt = False
                            Exit Function
                        End If
                    ElseIf igJobRptNo = 3 Then              'CPM podcast
                        If Not gOpenPrtJob("BRCPM.Rpt", , blCustomizeLogo) Then '
                            gGenReportCt = False
                            Exit Function
                        End If
                    'Else                                        'summary pass
'                    ElseIf igJobRptNo = 3 Then
                    ElseIf igJobRptNo = 4 Then                      '1-06-21 job #3 added; others adjusted, summary pass
                        If RptSelCt!ckcSelC6(1).Value = vbChecked Then        'include research, assume to show prices
                            If RptSelCt!ckcSelC6(0).Value = vbChecked Then      'with rates   & research
                                igBRSum = True                                  'need to know which summary version is being processed
                                    If Not gOpenPrtJob("BRSum.Rpt", , blCustomizeLogo) Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                igBRSum = True                  '1-06-21
'                                If Not gOpenPrtJob("BRSumnor.Rpt", , blCustomizeLogo) Then   'Research without rates
                                If Not gOpenPrtJob("BRSum.Rpt", , blCustomizeLogo) Then   '1-06-21 Research without rates
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        ElseIf RptSelCt!ckcSelC6(0).Value = vbChecked Then
                            igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
                            If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomizeLogo) Then 'exclude research, prices optional
                                gGenReportCt = False
                                Exit Function
                            End If
                        Else
                            igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
                            If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomizeLogo) Then 'exclude research, prices optional
                                gGenReportCt = False
                                Exit Function
                            End If
                        End If
'                    ElseIf igJobRptNo = 4 Then
                    ElseIf igJobRptNo = 5 Then                      '1-06-21  job #3 added, others adjusted
                        igBRSumZer = True                       '1-28-10 need to know which filter to send to crystal
                        If Not gOpenPrtJob("BRSumZer.Rpt", , blCustomizeLogo) Then 'exclude research, prices optional
                            gGenReportCt = False
                            Exit Function
                        End If
                    'Else
                    '    If Not gOpenPrtJob("BRNTR.Rpt") Then 'exclude research, prices optional
                    '        gGenReportCt = False
                    '        Exit Function
                    '    End If
                    ElseIf igJobRptNo = 3 Then          '1-06-21 this CPM job added
                        If Not gOpenPrtJob("BRCPM.Rpt", , blCustomizeLogo) Then 'exclude research, prices optional
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                'Else                'portrait version of BR
                '    If Not gOpenPrtJob("BRPortr.Rpt") Then
                '        gGenReportCt = False
                '        Exit Function
                '    End If
                '
                'End If
            'ElseIf rbcRptType(1).Value Then
            ElseIf ilListIndex = CNT_INSERTION Then                    'Insertion Notice (details only)
                blCustomizeLogo = True
                'If Not RptSelCt!rbcSelCInclude(2).Value Then      'proposals or wide contract                     'proposals/BR contract
                    'If igJobRptNo = 1 Then                      'detail pass
                        If RptSelCt!ckcSelC6(1).Value = vbChecked Then        'If including research, show wide with everything if including rates
                            If RptSelCt!ckcSelC6(0).Value = vbChecked Then      'with rates
                                If Not gOpenPrtJob("InsBR.Rpt", , blCustomizeLogo) Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("InsBrNor.Rpt", , blCustomizeLogo) Then 'with research but without CPP/CPM, and any other rates
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        Else                                    'exclude research, prices optional

                            If Not gOpenPrtJob("InsBrZer.Rpt", , blCustomizeLogo) Then
                                gGenReportCt = False
                                Exit Function
                            End If
                            
 
                        End If
                    'Else                                        'summary pass
                    'End If
                'End If
            ElseIf ilListIndex = CNT_PAPERWORK Then
                If RptSelCt!rbcSelC7(2).Value Then          '8-14-15 show acq cost only for all lines
                    If Not gOpenPrtJob("PapwkDetAcq.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                Else
                    If RptSelCt!rbcSelCInclude(0).Value Then      'contract summary (vs detail)
                        If Not gOpenPrtJob("Papwksum.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("Papwkdet.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                End If
            ElseIf ilListIndex = CNT_BOB_BYCNT Then 'Projection
                If RptSelCt!rbcSelCInclude(0).Value Then
                        If Not gOpenPrtJob("ProjMAdv.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                ElseIf RptSelCt!rbcSelCInclude(1).Value Then
                    If Not gOpenPrtJob("ProjMSls.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("ProjMVeh.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = CNT_RECAP Then 'Recap
                If Not gOpenPrtJob("ChfRecap.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If

            ElseIf ilListIndex = 8 Then  'MG's
                If Not gOpenPrtJob("MG.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = 9 Then  'Sales Spot Tracking
                If Not gOpenPrtJob("TrakSale.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_COMLCHG Then  'Commercial changes 10-25-00 converted to crystal
                If Not gOpenPrtJob("Cmmlchg.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_HISTORY Then  'Contract History
                If Not gOpenPrtJob("CntHist.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_AFFILTRAK Then  'Affiliate Spot Tracking
                If Not gOpenPrtJob("TrakAff.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf (ilListIndex = CNT_BOB_BYSPOT) Or (ilListIndex = 16) Then    'Spot Projection
                If ilListIndex = CNT_BOB_BYSPOT Then
                    If RptSelCt!rbcSelCSelect(0).Value Then  'set last sunday of first week
                        slType = "Weekly"
                    ElseIf RptSelCt!rbcSelCSelect(1).Value Then  'set last date of 12 standard periods
                        slType = "Standard"
                    ElseIf RptSelCt!rbcSelCSelect(2).Value Then  'set last date of 12 corporate periods
                        slType = "Corporate"
                    ElseIf RptSelCt!rbcSelCSelect(3).Value Then  'set last date of 12 calendar periods
                        slType = "Calendar"
                    End If
                Else    'Get date from combo box
                    slType = RptSelCt!cbcSel.List(RptSelCt!cbcSel.ListIndex)
                    ilRet = gParseItem(slType, 4, " ", slType)    'Get application name
                End If
                If RptSelCt!rbcSelCInclude(0).Value Then
                    If Not RptSelCt!ckcSelC8(0).Value = vbChecked Then    'not summary only
                        If Not gOpenPrtJob("PrjSDAdv.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PrjSSAdv.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                ElseIf RptSelCt!rbcSelCInclude(1).Value Then
                    If Not RptSelCt!ckcSelC8(0).Value = vbChecked Then    'not summary only
                        If Not gOpenPrtJob("PrjSDSls.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PrjSSSls.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                Else
                    '4-09-08 Summary and detail for Vehicle or Agency options now the same rpt module
                    'If Not RptSelCt!ckcSelC8(0).Value = vbChecked Then    'not summary only
                        'If Not gOpenPrtJob("PrjSDVeh.Rpt") Then
                        If Not gOpenPrtJob("PrjVhAg.Rpt") Then

                            gGenReportCt = False
                            Exit Function
                        End If
                    'Else
                    '    If Not gOpenPrtJob("PrjSSVeh.Rpt") Then
                    '        gGenReportCt = False
                    '        Exit Function
                    '    End If
                    'End If
                End If
                    'Report!crcReport.ReportFileName = sgRptPath & "ProjWeek.Rpt"
            ElseIf (ilListIndex = CNT_QTRLY_AVAILS) Then          'Quarterly Avails
                If RptSelCt!rbcSelC4(0).Value Then          'qtrly Summary vs qtrly booked
                    If RptSelCt!rbcSelCInclude(0).Value Then  'DP (vs days within DP or DP within days)
                    '10-20-11 avrsumpc, avrdpdypc & avrdydppc were never called since rbcSelC11(0).value was defaulted true and now allowed to be altered.
                    'remove to call them as they are currently not in-use
'                        If RptSelCt!rbcSelCSelect(3).Value = True And RptSelCt!rbcSelC11(1).Value = True Then    '2-4-05  sellout % and separate 30/60 % values?
'                            If Not gOpenPrtJob("AvrSumPC.Rpt") Then
'                                gGenReportCt = False
'                                Exit Function
'                            End If
'                        Else
                            'for sellout % and different 30/60 percent sellout
                            If Not gOpenPrtJob("AvrSum.Rpt") Then
                                gGenReportCt = False
                                Exit Function
                            End If
'                        End If
                    ElseIf RptSelCt!rbcSelCInclude(1).Value Then  'Days within daypart
'                        If RptSelCt!rbcSelCSelect(3).Value = True And RptSelCt!rbcSelC11(1).Value = True Then    '2-4-05  sellout % and separate 30/60 % values?
'                            If Not gOpenPrtJob("AvrDyDPPC.Rpt") Then
'                                gGenReportCt = False
'                                Exit Function
'                            End If
'                        Else
                            If Not gOpenPrtJob("AvrDyDP.Rpt") Then
                                gGenReportCt = False
                                Exit Function
                            End If
'                        End If
                    ElseIf RptSelCt!rbcSelCInclude(2).Value Then  'Daypart within days
'                        If RptSelCt!rbcSelCSelect(3).Value = True And RptSelCt!rbcSelC11(1).Value = True Then    '2-4-05  sellout % and separate 30/60 % values?
'                            If Not gOpenPrtJob("AvrDPDyPC.Rpt") Then
'                                gGenReportCt = False
'                                Exit Function
'                            End If
'                        Else
                        If Not gOpenPrtJob("AvrDPDy.Rpt") Then
                                gGenReportCt = False
                                Exit Function
                            End If
'                        End If
                    End If
                ElseIf RptSelCt!rbcSelC4(1).Value Then            'qtrly detail
                    If Not gOpenPrtJob("AvrQdet.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = CNT_AVG_PRICES Then    'avg spot prices (weekly & monthly)
                '12-24-08 make prepass for Avg Spot Price report
                If Not gOpenPrtJob("AvgSpotPrice.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                End If
'                If RptSelCt!rbcSelCInclude(0).Value Then    'slsp option
'                    If Not gOpenPrtJob("Avgslsp.Rpt") Then
'                        gGenReportCt = False
'                        Exit Function
'                    End If
'                Else                                        'vehicle option
'                    If Not gOpenPrtJob("Avgveh.Rpt") Then
'                        gGenReportCt = False
'                        Exit Function
'                    End If
'                End If
            ElseIf ilListIndex = CNT_ADVT_UNITS Then    'Advertiser Units Ordered
                If Not gOpenPrtJob("AdvtUnitsOrd.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_SALES_CPPCPM Then     'sales by CPP CPM
                If Not gOpenPrtJob("slscpppm.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_AVGRATE Then           'AVerage Rate  (detail or summary)
                slStr = RptSelCt!edcSelCTo.Text
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If
                
                slStr = RptSelCt!edcSelCTo1.Text
                If RptSelCt!rbcSelCSelect(0).Value Then     'week
                    ilRet = gVerifyInt(slStr, 1, 4)         'quarter selection
                    If ilRet = -1 Then
                        mReset
                        RptSelCt!edcSelCTo1.SetFocus                 'invalid # periods
                        gGenReportCt = False
                        Exit Function
                    End If
                Else
                    ilRet = gVerifyInt(slStr, 1, 12)         'starting month
                    If ilRet = -1 Then
                        mReset
                        RptSelCt!edcSelCTo1.SetFocus                 'invalid starting month
                        gGenReportCt = False
                        Exit Function
                    End If
                    
                    slTemp = (RptSelCt!edcSelCFrom1.Text)      '# months to print
                    ilRet = gVerifyInt(slTemp, 1, 12)         'verify # months to print
                    If ilRet = -1 Then
                        mReset
                        RptSelCt!edcSelCFrom1.SetFocus
                        gGenReportCt = False
                        Exit Function
                    End If
                End If
                igMonthOrQtr = Val(slStr)
                
                'TTP 10601 - Average 30" Unit Rate Report by Summary is showing as Detail
                If RptSelCt!rbcOutput(4).Value = False Then          'Not Exporting
                    If RptSelCt!rbcSelC4(0).Value Then          'Detail
                        If Not gOpenPrtJob("avgrate.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else                                        'summary
                        If Not gOpenPrtJob("avgrtsm.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                End If
            ElseIf ilListIndex = CNT_TIEOUT Then           'Tie Out
                slStr = RptSelCt!CSI_CalFrom.Text           'Date: 1/8/2020 added CSI calendar control for date entry --> edcSelCFrom.Text
                If Not gValidDate(slStr) Then
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus           'Date: 1/8/2020 added CSI calendar control for date entry --> edcSelCFrom.SetFocus
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCFrom1.Text
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If
                If RptSelCt!rbcSelCInclude(0).Value Then         'office option
                    If Not gOpenPrtJob("TieOutOf.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                Else                                        'vehicle option
                    If Not gOpenPrtJob("TieOutVh.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                End If
                'igYear = Val(slStr)
            ElseIf ilListIndex = CNT_BOB Then               'Billed & Booked
                slStr = RptSelCt!edcSelCTo.Text             'check for presence of selective contract #
                If slStr = "" Then
                    slStr = "0"
                End If
                If Not IsNumeric(slStr) Then
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid contract #
                    gGenReportCt = False
                    Exit Function
                End If

                ilRet = mVerifyMonthYrPeriods(RptSelCt, ilListIndex, RptSelCt!rbcSelC9(0))      'rbcSelc9(0) = true if corp month selection
                If ilRet = True Then            'got an error in conversion of input
                    gGenReportCt = False
                    Exit Function
                End If

                'net-net option for vehicle.  this version shows the owners share only of the revenue
                'If RptSelCt!rbcSelC7(2).Value And RptSelCt!rbcSelCInclude(2).Value = True Then            '2-28-01  Net-Net version
                '11-17-06 change to use list box for unique version of B & B net-net option
                '11-17-06 Screen selectivity changed to list box because another sort option has been added and
                'it didnt fit on the screen.  The list box answer has been converted into the original radio button answers:
                'cbcset2 : 0 = adv, 1 = agy, 2 = owner, 3 = slsp, 4 =vehicle, 5 = vehicle grossnet, 6 = vehicle / participant
                
                'Totals by Detail (no slsp sort):
                '   Bobdet.rpt - all sort options except vehicle gross net & slsp, with/without skip to new page
                'Totals by Detail by Slsp:
                '   Bobdtskp.rpt - Detail by slsp only, new page skip, no vehicle subtotal
                '   Bobvehdt - detail by slsp with vehicle subtotals only, skip/no skip to new page
                'Totals by Advt (no Slsp sort):
                '   BobAdv.rpt - totals by advt for all sorts , plus page skip, except vehicle gross/net and slsp without vehicle subt (no skip)
                'Totals by ADvt slsp sort only:
                '   Bobadskp.rpt - totals by advt, skip to new page, Slsp sort ; no vehicle subtotals.  No skip  goes to bobadv.rpt
                '   bobvehad.rpt - totals by advt, skip or no skip, slsp sort with vehicle subtotals
                'Totals by Summary -
                '   bobsum - all sorts except vehicle gross/net and slsp with  vehicle subtotals
                'Totals by Summary for Slsp sort only-
                '   bobvehsm - sort by slsp with vehicle subtotals
                '
                If RptSelCt!cbcSet2.ListIndex = 5 Then          'its the net net version of billed & booked, showing owners share plus gross line, net line & net net
                    If Not gOpenPrtJob("BobNN.Rpt") Then
                        gGenReportCt = False
                        Exit Function
                    End If
                Else
                    If RptSelCt!ckcSelC8(2).Value = vbChecked Then            'Skip to new pages each new group
                        If RptSelCt!rbcSelC4(0).Value Then            'detail version
                            If RptSelCt!rbcSelCInclude(1).Value Then  'slsp option, a unique version to skip to new page on slsp (not office which is the generalized level to skip)
                                If RptSelCt!ckcSelC10(1).Value = vbChecked Then     'add vehicle subtotals by slsP
                                     If Not gOpenPrtJob("Bobvehdt.Rpt") Then        'combined Skip to New Page in this module, remove bobvhdtk
                                        gGenReportCt = False
                                        Exit Function
                                    End If
                                Else
                                    'detail, Page skip for slsp sort, no veh subtotals
                                    If Not gOpenPrtJob("BobDtskp.Rpt") Then
                                        gGenReportCt = False
                                        Exit Function
                                    End If
                                End If
                            Else        'Detail Page skipping and NOT by slsp
                                'Skip page, detail, all other sorts except slsp and vehicle gross/net (bobnn.rpt)
                                If Not gOpenPrtJob("BobDet.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If

                            End If
                        ElseIf RptSelCt!rbcSelC4(1).Value Then        'Summay with advt level included
                            If RptSelCt!rbcSelCInclude(1).Value Then    'slsp option, skip to new page eachslsp (not office, which is the generalized level to skip on)
                                If RptSelCt!ckcSelC10(1).Value = vbChecked Then      'slsp option with vehicle subtotals?
                                    'If Not gOpenPrtJob("Bobvhadk.Rpt") Then
                                     If Not gOpenPrtJob("Bobvehad.Rpt") Then    '10-30-02 combine Skip to New Page in this module, remove bobvhadk

                                        gGenReportCt = False
                                        Exit Function
                                    End If
                                Else
                                    If Not gOpenPrtJob("BobAdskp.Rpt") Then
                                        gGenReportCt = False
                                        Exit Function
                                    End If
                                End If
                            Else        'subtotals by advt (vs contract), NOT slsp option

                                'If Not gOpenPrtJob("BobAdvSk.Rpt") Then
                                '    gGenReportCt = False
                                '    Exit Function
                                'End If
                                '8-4-00 Combine bobadvsk.rpt with bobadv.  Send flag to skip to new page each group

                                If Not gOpenPrtJob("BobAdv.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If

                            End If
                        Else
                            If Not RptSelCt!rbcSelCInclude(1).Value Then        '8-4-00 not slsp
                                If Not gOpenPrtJob("BobSum.Rpt") Then   'summary without advt level
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(1).Value And RptSelCt!ckcSelC10(1).Value = vbChecked Then     'slsp with veh subtotals?
                                'If Not gOpenPrtJob("Bobvhsmk.Rpt") Then
                                 If Not gOpenPrtJob("Bobvehsm.Rpt") Then        '10-30-02 Combine Skip to New Page in thismodule, remove bobvhsmk

                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("BobSumSk.Rpt") Then   'slsp summary, no veh subtotals without advt level
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Else                                            'no page skips
                        If RptSelCt!rbcSelC4(0).Value Then            'detail version
                            If RptSelCt!rbcSelCInclude(1).Value And RptSelCt!ckcSelC10(1).Value = vbChecked Then        'slsp with veh option
                                If Not gOpenPrtJob("Bobvehdt.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("BobDet.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        ElseIf RptSelCt!rbcSelC4(1).Value Then        'Summay with advt level included
                            If RptSelCt!rbcSelCInclude(1).Value And RptSelCt!ckcSelC10(1).Value = vbChecked Then        'slsp with veh option
                                If Not gOpenPrtJob("Bobvehad.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("BobAdv.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        Else
                            If RptSelCt!rbcSelCInclude(1).Value And RptSelCt!ckcSelC10(1).Value = vbChecked Then        'slsp with veh option
                                If Not gOpenPrtJob("Bobvehsm.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            Else
                                If Not gOpenPrtJob("BobSum.Rpt") Then   'summary without advt level
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If

                'move the validity test prior to opening crystal reports
'                ilRet = mVerifyMonthYrPeriods(RptSelCt, ilListIndex)
'                If ilRet = True Then            'got an error in conversion of input
'                    gGenReportCt = False
'                    Exit Function
'                End If


        ElseIf ilListIndex = CNT_BOBRECAP Then              '4-14-05
            'ilRet = mVerifyBOBInput()        'verify # qtr, year and # periods input
            ilRet = mVerifyMonthYrPeriods(RptSelCt, ilListIndex, RptSelCt!rbcSelC9(0))        '7-3-08 allow starting month # vs starting qtr
            If ilRet = True Then            'got an error in conversion of input
                gGenReportCt = False
                Exit Function
            End If
            If Not gOpenPrtJob("BobRecap.Rpt") Then
                gGenReportCt = False
                Exit Function
            End If

        ElseIf ilListIndex = CNT_SALESACTIVITY Then         'Sales Activity Increase Decrease
                slStr = RptSelCt!CSI_CalFrom.Text   'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.Text                 'edit effective date
                If Not gValidDate(slStr) Then
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus   'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCTo.Text                   'edit year
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCTo1.Text                  'edit qtr
                ilRet = gVerifyInt(slStr, 1, 4)
                If ilRet = -1 Then
                    mReset
                    RptSelCt!edcSelCTo1.SetFocus                 'invalid qtr
                    gGenReportCt = False
                    Exit Function
                End If
                igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
                If Not gOpenPrtJob("SalesAct.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_DAILY_SALESACTIVITY Then        '6-5-01
                If Not gOpenPrtJob("SalesAct.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
'***********************************
            ElseIf ilListIndex = CNT_SALESCOMPARE Then
                'validity check the input dates
'                slStr = RptSelCt!edcSelCFrom.Text
'                ilRet = gVerifyYear(slStr)
'                If ilRet = 0 Then
'                    mReset
'                    RptSelCt!edcSelCFrom.SetFocus                 'invalid year
'                    gGenReportCt = False
'                    Exit Function
'                End If
'                slStr = RptSelCt!edcSelCFrom1.Text                'month text (jan, feb...)
'                gGetMonthNoFromString slStr, ilIndex
'                If ilIndex = 0 Then                               'input isn't text month name, try month #
'                    If Val(slStr) > 0 And Val(slStr) < 13 Then
'                        ilIndex = Val(slStr)
'                    Else
'                        mReset
'                        RptSelCt!edcSelCFrom1.SetFocus            'invalid month
'                        gGenReportCt = False
'                        Exit Function
'                    End If
'                End If
'
'                slStr = RptSelCt!edcSelCTo.Text                   'no months
'                If Val(slStr) < 1 Or Val(slStr) > 12 Then
'                    mReset
'                    RptSelCt!edcSelCTo.SetFocus
'                    gGenReportCt = False
'                    Exit Function
'                End If
'                igPeriods = Val(slStr)                      '7-7-14

                'this report has only calendar and standard, corporate is hidden as index 0 (and is never set)
                ilRet = mVerifyMonthYrPeriods(RptSelCt, ilListIndex, RptSelCt!rbcSelC9(0))         '3-20-18 use common verify routine
                If ilRet = True Then            'got an error in conversion of input
                    gGenReportCt = False
                    Exit Function
                End If
                
                'cbcSet1 and cbcSet2 indicate the major and minor selectivity.  The major sort selection as been converted to
                'radio button selectivity due to retrofitting into previous code.

                'rbcSelc4(0) = detail ; rbcSelC4(1) = Summary
                'rbcSelCInclude(0) = advt
                'rbcSelCInclude(1) = Slsp
                'rbcSelCInclude(2) = Agency
                'rbcSelCInclude(3) = business category
                'rbcSelCInclude(4) = product protection
                'rbcselCInclude(5) = vehicle
                'rbcSelCInclude(6) = vehicle group    this is new 9-04-07

                'new as of 09-04-07
                'cbcSet1.listIndex = 0 Advt; 1 = agy, 2 = bus cat, 3 = prod prot, 4 = slsp, 5 = vehicle, 6 = vehicle group
                'cbcSet2.ListIndex = 0 = none; 1 Advt; 2 = agy, 3 = bus cat, 4 = prod prot, 5 = slsp, 6 = vehicle, 7 = vehicle group'
                '
                If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then          'NOT top down report
                    If Trim$(RptSelCt!edcText.Text) = "" Then               'no pace
                        If Not gOpenPrtJob("SlsCompare.Rpt") Then
                            gGenReportCt = False
                        Exit Function
                        End If
                    Else                                                    'pacing
                        If Not gOpenPrtJob("SlsComparePace.Rpt") Then
                            gGenReportCt = False
                        Exit Function
                        End If
                    End If
                Else
                    'this section is for detail TOP DOWN only
                    If RptSelCt!rbcSelC4(0).Value Then              'detail version
                        If Trim$(RptSelCt!edcText.Text) = "" Then               'no pace
                            If RptSelCt!rbcSelCInclude(0).Value Then    'advt
                                If Not gOpenPrtJob("TDAdvDt.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(1).Value Then    'Slsp
                                If Not gOpenPrtJob("TDSlsDt.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(3).Value Or RptSelCt!rbcSelCInclude(4).Value Or RptSelCt!rbcSelCInclude(5).Value Or RptSelCt!rbcSelCInclude(6).Value Then      'Agy, Prod Protection, Bus Category, vehicle or vehicle group
                                If Not gOpenPrtJob("TDDt.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        Else                                        'pacing detail
                            If RptSelCt!rbcSelCInclude(0).Value Then    'advt
                                If Not gOpenPrtJob("TDAdvDtPace.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(1).Value Then    'Slsp
                                If Not gOpenPrtJob("TDSlsDtPace.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(3).Value Or RptSelCt!rbcSelCInclude(4).Value Or RptSelCt!rbcSelCInclude(5).Value Or RptSelCt!rbcSelCInclude(6).Value Then      'Agy, Prod Protection, Bus Category, vehicle or vehicle group
                                If Not gOpenPrtJob("TDDtPace.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        End If
                    Else            'Summary only, TOP DOWN
                        If Trim$(RptSelCt!edcText.Text) = "" Then               'no pace

                            If RptSelCt!rbcSelCInclude(0).Value Then    'advt
                                If Not gOpenPrtJob("TDAdvDt.Rpt") Then   'detail & summary are the same for advt
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(1).Value Then    'Slsp
                                If Not gOpenPrtJob("TDSlsSm.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                                '4-25-06 add vehicle option
                            ElseIf RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(3).Value Or RptSelCt!rbcSelCInclude(4).Value Or RptSelCt!rbcSelCInclude(5).Value Or RptSelCt!rbcSelCInclude(6).Value Then    'Agy, Prod Protection, Bus Category, vehicle or vehicle group
                                If Not gOpenPrtJob("TDSm.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                        Else                        'pace
                             If RptSelCt!rbcSelCInclude(0).Value Then    'advt
                                If Not gOpenPrtJob("TDAdvDtPace.Rpt") Then   'detail & summary are the same for advt
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            ElseIf RptSelCt!rbcSelCInclude(1).Value Then    'Slsp
                                If Not gOpenPrtJob("TDSlsSmPace.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                                '4-25-06 add vehicle option
                            ElseIf RptSelCt!rbcSelCInclude(2).Value Or RptSelCt!rbcSelCInclude(3).Value Or RptSelCt!rbcSelCInclude(4).Value Or RptSelCt!rbcSelCInclude(5).Value Or RptSelCt!rbcSelCInclude(6).Value Then    'Agy, Prod Protection, Bus Category, vehicle or vehicle group
                                If Not gOpenPrtJob("TDSmPace.Rpt") Then
                                    gGenReportCt = False
                                    Exit Function
                                End If
                            End If
                       
                        End If
                    End If
                End If          'If RptSelCt!ckcSelC10(0).Value = vbUnchecked Then          'NOT top down report
'^^^^^^^^^^^^^^^^^^^^^^^^^
            ElseIf ilListIndex = CNT_CUMEACTIVITY Then
                slStr = RptSelCt!edcSelCTo.Text                 'entered year
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSelCt!edcSelCTo.SetFocus                 'invalid year
                    gGenReportCt = False
                    Exit Function
                End If
                slStr = RptSelCt!edcSelCTo1.Text                  'edit qtr
                ilRet = gVerifyInt(slStr, 1, 4)
                If ilRet = -1 Then
                    mReset
                    RptSelCt!edcSelCTo1.SetFocus                 'invalid qtr
                    gGenReportCt = False
                    Exit Function
                End If
                igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
                If RptSelCt!rbcSelC4(0).Value Then              'detail
                    If RptSelCt!rbcSelCInclude(0).Value Then    'advt option
                        If Not gOpenPrtJob("slsactad.Rpt") Then    'contract within advt option
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else                                        'vehicle, agency, & demo option
                        If Not gOpenPrtJob("slsactmo.Rpt") Then     'was cumactdt (chgd 4/28/98) for advt option
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                Else                                            'summary
                    If RptSelCt!rbcSelCInclude(0).Value Then    'advt option, summary
                        If Not gOpenPrtJob("slactasm.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else                                        'vehicle, agency, & demo option : summary
                        If Not gOpenPrtJob("slactmsm.Rpt") Then     'was cumactsm (chgd 4/28/98) for advt option
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                    'If Not gOpenPrtJob("CumActSm.Rpt") Then
                    '    gGenReportCt = False
                    '    Exit Function
                    'End If
                End If
            ElseIf ilListIndex = CNT_MAKEPLAN Then              'avg prices need to make plan
                If RptSelCt!rbcSelC4(0).Value Then              'weekly
                    If RptSelCt!rbcSelC7(0).Value Then          'detail weekly
                        If Not gOpenPrtJob("PlanWkDt.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PlanWkSm.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                Else                                            'monthly/quarterly
                    If RptSelCt!rbcSelC7(0).Value Then          'detail monthly
                        If Not gOpenPrtJob("PlanMnDt.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PlanMnSm.Rpt") Then
                            gGenReportCt = False
                            Exit Function
                        End If
                    End If
                End If
            ElseIf ilListIndex = CNT_VEHCPPCPM Then
                If Not gOpenPrtJob("Cppmveh.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_SALESANALYSIS Then
                If Not gOpenPrtJob("SalesAna.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
             ElseIf ilListIndex = CNT_SALESACTIVITY_SS Or ilListIndex = CNT_SALESPLACEMENT Then '7-25-02
                If Not gOpenPrtJob("SlsACTSS.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_VEH_UNITCOUNT Then     '7-15-03
                If Not gOpenPrtJob("VhUnitCt.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_LOCKED Then     '4-5-06
                If Not gOpenPrtJob("LockAvail.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
             ElseIf ilListIndex = CNT_GAMESUMMARY Then     '7-14-06
                If Not gOpenPrtJob("GameSum.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_PAPERWORKTAX Then     '4-9-07
                If Not gOpenPrtJob("PapWkTax.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_BOBCOMPARE Then        '9-13-07
                If Not gOpenPrtJob("BOBCompare.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_CONTRACTVERIFY Then        '4-8-13
                If Not gOpenPrtJob("ContrVerify.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then        '10-6-15
                If Not gOpenPrtJob("InsertionActivity.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
                ElseIf ilListIndex = CNT_XML_ACTIVITY Then        '10-6-15
                If Not gOpenPrtJob("XMLActivity.Rpt") Then
                    gGenReportCt = False
                    Exit Function
                End If
            End If
    End Select
    gGenReportCt = True
End Function
'
'
'                   mBOBCrystal - Pass formulas to Crystal reports
'                               for Billed & Booked and Sales
'                               Comparisons reports
'
Function mBOBCrystal() As Integer
Dim slReserve As String
Dim slStr As String
Dim slSelection As String
Dim slTime As String
Dim slDate As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim ilListIndex As Integer
Dim slInclude As String
Dim slExclude As String

    slInclude = ""
    slExclude = ""

    ilListIndex = RptSelCt!lbcRptType.ListIndex         'report selected
    If RptSelCt!rbcSelCSelect(0).Value Then           'use package lines
        slReserve = "Use Package lines "
    Else
        slReserve = "Use Airing lines "         '7-12-01
        If tgSpf.iPkageGenMeth = 1 Then         'calc $ method for ordered packages by Line? Only gather from lines, no receivables
            slReserve = Trim$(slReserve) & ", ignore billing "
        End If
    End If
    If tgSpf.sInvAirOrder <> "S" Then            'bill as ordered, update as ordered, no adjustments at all
        If RptSelCt!ckcSelC8(0).Value = vbChecked Then                'subt misses
            slReserve = slReserve & "; for standard lines subtract misses"
        End If                                              'show nothing if ignoring them
        'Else
        '    slReserve = slReserve & "exclude misses and makegoods"
        'End If
        If RptSelCt!ckcSelC8(1).Value = vbChecked Then                  'count mg when they air?
            slReserve = slReserve & "; for standard lines count MGs"
        End If
    End If

    'if running Billed & Booked by slsp and only primary slsp to be used, show in report header
    If igRptCallType = CONTRACTSJOB And (ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBRECAP Or ilListIndex = CNT_SALESCOMPARE Or ilListIndex = CNT_BOBCOMPARE) Then  '4-14-05

        If (ilListIndex <> CNT_SALESCOMPARE And ilListIndex <> CNT_BOBCOMPARE) Then   'Sales Comparison and B & B Comparisons always split slsp
            If RptSelCt!rbcSelCInclude(1).Value And RptSelCt!ckcSelC10(0).Value = vbUnchecked Then      '9-18-06 test for splits changed
                slReserve = slReserve & "; without slsp splits "
            End If
        End If

'        If RptSelCt!ckcSelC6(1).Value = vbChecked Then    'use air time
'            slReserve = slReserve & "; Air Time"
'        End If
'        If RptSelCt!ckcSelC6(2).Value = vbChecked Then      'use ntr
'            If RptSelCt!ckcSelC6(1).Value = vbChecked Then
'                slReserve = slReserve & ", NTR"
'            Else
'                slReserve = slReserve & "; NTR"
'            End If
'        End If
'        If RptSelCt!ckcSelC6(3).Value = vbChecked Then      'use hard cost?
'            If RptSelCt!ckcSelC6(1).Value = vbChecked Or RptSelCt!ckcSelC6(2).Value = vbChecked Then
'                slReserve = slReserve & ", Hard Cost"
'            Else
'                slReserve = slReserve & "; Hard Cost"
'            End If
'        End If


        If ilListIndex = CNT_BOB And RptSelCt!ckcSelC13(0).Value = vbChecked Then            'billed & booked only has option to process the acquisition costs only
                                                'from rvf/phf and/or schd lines
            slReserve = slReserve & "; Acq. Costs Only"
        End If

        gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Orders"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(5), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(6), slInclude, slExclude, "Promo"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(0), slInclude, slExclude, "Trades"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(1), slInclude, slExclude, "Air Time"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(10), slInclude, slExclude, "Digital"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(2), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelCt!ckcSelC6(3), slInclude, slExclude, "Hard Cost"
        gIncludeExcludeCkc RptSelCt!ckcSelC12(0), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelCt!ckcSelC12(1), slInclude, slExclude, "Non-Polit"
        '9/25/20 - TTP 9952 - show Adj on CNT_SALESCOMPARE
        If ilListIndex = CNT_SALESCOMPARE Then gIncludeExcludeCkc RptSelCt!ckcInclRevAdj, slInclude, slExclude, "Adj"
        slReserve = slReserve & ", " & Trim$(slInclude)
    End If
    If Not gSetFormula("Adjustments", "'" & slReserve & "'") Then
        mBOBCrystal = -1
        Exit Function
    End If

    If RptSelCt!rbcSelC7(0).Value Then                'Gross
        If Not gSetFormula("GrossNet", "'G'") Then
            mBOBCrystal = -1
            Exit Function
        End If
    'Slsp option (rbcselcInclude(1)) & sub-totals by vehicle (ckcSelc10(1)) and triple-net option (rbcselc7(2))
   ' ElseIf (RptSelCt!rbcSelCInclude(1).Value = True And RptSelCt!ckcSelC10(1).Value = vbChecked And RptSelCt!rbcSelC7(2).Value = True) Then
    '    If Not gSetFormula("GrossNet", "'T'") Then
   '         mBOBCrystal = -1
    '        Exit Function
    '    End If
    ' Vehicle/Participant option (rbcselcInclude(4)) & net-net (rbcSelC7(2))
    'Billed & Booked Recap and net-net option selected
    'ElseIf (RptSelCt!rbcSelCInclude(4).Value = True And RptSelCt!rbcSelC7(2).Value = True) Or (ilListIndex = CNT_BOBRECAP And igRptCallType = CONTRACTSJOB And RptSelCt!rbcSelC7(2).Value = True) Then
    '    If Not gSetFormula("GrossNet", "'D'") Then
    '        mBOBCrystal = -1
    '        Exit Function
    '    End If
    'all version of B & B can have T-Net option
    ElseIf RptSelCt!rbcSelC7(2).Value = True Then           'triple net
        If Not gSetFormula("GrossNet", "'T'") Then
            mBOBCrystal = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("GrossNet", "'N'") Then
            mBOBCrystal = -1
            Exit Function
        End If
    End If

    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
    If Trim$(slStr) = "" Then               'no last billed date yet, system start up
        slStr = "1/1/1975"
    End If
    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
    If Not gSetFormula("LastBilled", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        mBOBCrystal = -1
        Exit Function
    End If
    
    If RptSelCt!rbcSelC9(4).Value = True Then           '1-25-21 if bill method, show the last billed for calendar
        gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slStr      'convert last bdcst billing date to string
        If Trim$(slStr) = "" Then               'no last billed date yet, system start up
            slStr = "1/1/1975"
        End If
        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
        If Not gSetFormula("LastBilledCal", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            mBOBCrystal = -1
            Exit Function
        End If
    End If

    'gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
    gRandomDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If ilListIndex = CNT_BOBCOMPARE Then
        slSelection = slSelection & " and ({GRF_Generic_Report.grfPer13Genl} = 0)"   'budget records are retrieved via subreports
    End If
    If Not gSetSelection(slSelection) Then
        mBOBCrystal = -1
        Exit Function
    End If
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime        'run time to show on report
'    If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'        mBOBCrystal = -1
'        Exit Function
'    End If
End Function
'****************************************************************
'*                                                              *
'*      Procedure Name:mCntJob1_10                              *
'*                                                              *
'*             Created:6/16/93       By:D. LeVine               *
'*            Modified:              By:                        *
'*                                                              *
'*            Comments: Initialize Contract reports             *
'*      3/29/99 show user entered Start/end Active &            *
'               entered dates in paperwork summary heading      *
'*                                                              *
'*      6-24-04 ignore altered schedule lines (schstatus = 'A') *
'*               Business Booked by Contract                    *
'****************************************************************
Function mCntJob1_10(ilListIndex As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slInclStatus                                                                          *
'******************************************************************************************

    Dim slDate As String
    Dim slTime As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slSelection As String
    Dim illoop As Integer
    Dim slStr As String
    Dim slOr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilAllSelected As Integer
    Dim slInclude As String
    Dim slExclude As String
    Dim slMissedStart As String
    Dim slMissedEnd As String
    Dim slMGStart As String
    Dim slMGEnd As String
    Dim slActive As String
    Dim slEntered As String
    Dim llSingleCntr As Long
    Dim slLocalFeed As String
    Dim slUserID As String             '2-16-13

    mCntJob1_10 = 0
    'removed spots by Advt code--see rptselcb
'Exit Function
    'removed spots by date & time code--see rptselcb

    If (ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION) Then 'Contract BR or Insertions
        'If (Not igUsingCrystal) Then
        'Date: 1/13/2020 added CSI calendar control for date entries
        'If (RptSelCt!edcSelCFrom.Text <> "") And (RptSelCt!edcSelCFrom1.Text <> "") Then
        If (RptSelCt!CSI_CalFrom.Text <> "") And (RptSelCt!CSI_CalTo.Text <> "") Then
            'If StrComp(RptSelCt!edcSelCFrom1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_CalFrom.Text      'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
                If gValidDate(slDate) Then
                    slDate = RptSelCt!CSI_CalTo.Text    'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom1.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCt!CSI_CalTo.SetFocus     'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus       'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                    Exit Function
                End If
            Else
                slDate = RptSelCt!CSI_CalFrom.Text      'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus       'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!CSI_CalFrom.Text <> "" Then     'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.Text <> "" Then
            slDate = RptSelCt!CSI_CalFrom.Text          'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalFrom.SetFocus           'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                Exit Function
            End If
        ElseIf RptSelCt!CSI_CalTo.Text <> "" Then     'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom1.Text <> "" Then
            'If StrComp(RptSelCt!edcSelCFrom1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_CalTo.Text        'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom1.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_CalTo.SetFocus         'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
                    Exit Function
                End If
            End If
        End If
        'Date: 1/13/2020 added CSI calendar control for date entries
        'If (RptSelCt!edcSelCTo.Text <> "") And (RptSelCt!edcSelCTo1.Text <> "") Then
        If (RptSelCt!CSI_From1.Text <> "") And (RptSelCt!CSI_To1.Text <> "") Then
            'If StrComp(RptSelCt!edcSelCTo1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_To1.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_From1.Text          'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.Text
                If gValidDate(slDate) Then
                    slDate = RptSelCt!CSI_To1.Text      'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo1.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCt!CSI_To1.SetFocus       'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo1.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSelCt!CSI_From1.SetFocus         'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.SetFocus
                    Exit Function
                End If
            Else
                slDate = RptSelCt!CSI_From1.Text        'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_From1.SetFocus         'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.SetFocus
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!CSI_From1.Text <> "" Then       'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.Text <> "" Then
            slDate = RptSelCt!CSI_From1.Text            'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.Text
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_From1.SetFocus             'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo.SetFocus
                Exit Function
            End If
        ElseIf RptSelCt!CSI_To1.Text <> "" Then         'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo1.Text <> "" Then
            'If StrComp(RptSelCt!edcSelCTo1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_To1.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_To1.Text          'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo1.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!CSI_To1.SetFocus           'Date: 1/13/2020 added CSI calendar control for date entries --> edcSelCTo1.SetFocus
                    Exit Function
                End If
            End If
        End If

        'check validity of selective contract # entered
        If RptSelCt!edcTopHowMany.Text <> "" Then
            'open cnt hdr
            hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCHF)
                btrDestroy hmCHF
            End If
            llSingleCntr = Val(RptSelCt!edcTopHowMany.Text)     '10-20-06 avoid type mismatch error with Clng

            tmChfSrchKey1.lCntrNo = llSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCHF)
                btrDestroy hmCHF
                mReset
                RptSelCt!edcTopHowMany.SetFocus
                Exit Function
            End If
            ilRet = btrClose(hmCHF)
            btrDestroy hmCHF
        End If
        
        If (ilListIndex = CNT_INSERTION) Then           'cannot combine NTR if differences only since NTR doesnt do differences
            If RptSelCt!rbcOutput(3).Value = True Then
                If RptSelCt!edcResponse.Text <> "" Then
                    slDate = RptSelCt!edcResponse.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCt!edcResponse.SetFocus
                        Exit Function
                    End If
                End If
            End If
            
            If RptSelCt!ckcSelC5(0).Value = vbChecked And RptSelCt!ckcSelC12(0).Value = vbChecked Then      'differences only with NTR option?  disallow it
                RptSelCt!ckcSelC12(0).Value = vbUnchecked
                MsgBox "Combined NTR features has been disabled for Differences only option"
            End If
        End If

        'If RptSelCt!rbcSelCInclude(2).Value Then     'narrow contract
        '    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        '    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        '    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        '    If Not gSetSelection(slSelection) Then
        '        mCntJob1_10 = -1
        '        Exit Function
        '    End If
        'Else                                    'wide contract/proposal
        'End If
            'If (igJobRptNo = 1) Or (igJobRptNo = 2 And RptSelCt!rbcSelC4(1).Value) Then 'Detail or sumary pass & user requested summary
            '2-10-02
            'If (igJobRptNo = 1) Or (igJobRptNo = 2 And RptSelCt!rbcSelC4(1).Value) Then 'Detail or sumary pass & user requested summary
            '10-30-07


'   Columns indicate the job # for the specific task(left column).  The numbers below each job# indicate the flag in
'   cbfExtra2Byte to filter records in the .rpt
'
'   Task                    igJobRptNo#1    igJobRptNo2     igJobRptNo3-->4     igJobRptNo4-->5     igJobRptNo5-->3
'                           Detail          NTR             ResearchSum         Billing Sum         CPM Line IDs
'                                                                               0 for filter is used as a record in main .rpt to link the subreports to
'   Detail/Sum,Combine
'     AT & NTR w/Resrch     0               4               <> 4,5,-1,8         0
'   Detail Only Combine
'     AT & NTR w/Resrch     0
'   Summ Only Combine
'     AT & NTR w/Resrch                     4              <> 4,5,-1,8         0
'

    '12-21-20   Add CPM to show all line IDs, CPM Research Summary, and CPM billing summary
'   CPM Line IDs will add
'   another job (#5)

    
'   Detail/Sum,Combine
'     AT & NTR wo/Resrch     0              4                                   0
'   Detail Only Combine
'     AT & NTR wo/Resrch     0
'   Summ Only Combine
'     AT & NTR wo/Resrch                    4                                   0
'


'   These versions do not have a billing summary for NTR items; they are only shown on combined
'   Detail/Sum,Separate
'     AT & NTR w/Resrch     0               4               <> 4,5,-1,8,9    0
'   Detail Only Separate
'     AT & NTR w/Resrch     0
'   Summ Only Separate
'     AT & NTR w/Resrch                     4               <> 4,5,-1,8,9    0

'   Detail/Sum,Separate
'     AT & NTR wo/Resrch     0              4                               0
'   Detail Only Separate
'     AT & NTR wo/Resrch     0
'   Summ Only Separate
'     AT & NTR wo/Resrch                    4                               0

'       CBF Record types :  cbfExtra2Byte = 0:  detail
'                                           2 = vehicle summary
'                                           3 = contract tots
'                                           4 = NTR
'                                           5 = sports (games)
'                                           6 = for Installment contracts only for those vehicles that
'                                               were not on the air time schedule (i.e. NTRs).  These
'                                               vehicles need to show on the Installment Summary version
'                                           -1 = Key record for Insertion Order
'                                           8 = NTR billing summary (non-installment)
'                                           9 = CPM detail IDs
'                                           10 = CPM vehicle summary (for Research page)
'                                           11 = CPM billing summary by vehicle

'            If (igJobRptNo = 1) Or (igJobRptNo = 3 And RptSelCt!rbcSelC4(1).Value) Then 'Detail or sumary pass & user requested summary
            If (igJobRptNo = 1) Or (igJobRptNo = 4 And igDetSumBoth = 1) Then '1-06-21 Detail or sumary pass & user requested summary
                'gCurrDateTime slDate, slTime, slMonth, slDay, slYear       'can only get the current date & time once for all versions of the contract reports
                gUnpackDate igNowDate(0), igNowDate(1), slDate
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
                '9-14-09 no longer using lgNowTime for the time, getting milliseconds in time
                slTime = Trim$(str(lgNowTime))

                slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & slTime          '9-14-09
                
                '2-16-13 filter on gen date, time & urfcode
                slUserID = Trim$(str(tgUrf(0).iCode))
                slSelection = slSelection & " and {CBF_Contract_BR.cbfurfCode} = " & slUserID
                
                If igJobRptNo = 1 Then              'for pass 1 of BR (Detail), filter out summary records
                    If ilListIndex = CNT_INSERTION Then
                        slSelection = slSelection & " AND ({CBF_Contract_BR.cbfExtra2Byte} = -1 and {VEF_Vehicles.vefType} <> 'P' ) "
                        sgSelection = Trim$(slSelection)            '5-15-15
                    Else
                        slSelection = slSelection & " AND ({CBF_Contract_BR.cbfExtra2Byte} = 0)"
                    End If
                Else
                    'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} <> 4 and {CBF_Contract_BR.cbfExtra2Byte} <> 5 )"        '10-20-05 send  everything but NTR & the sports comments
                    'Research is included, need to know which version of the summary will be printed
                    'Both Research version and the billing summary come thru here
                    If igBRSumZer Then              '2-2-10 its the billing summary
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = -1 or {CBF_Contract_BR.cbfExtra2Byte} = 0)  "        'send only the record so that the airtime and NTR subreports can link to it,
                                        '7-12-10 plus the detail records which contain the monthly summaries
                    Else
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} <> 4 and {CBF_Contract_BR.cbfExtra2Byte} <> 5 and {CBF_Contract_BR.cbfExtra2Byte} <> -1 and {CBF_Contract_BR.cbfExtra2Byte} <> 8 and {CBF_Contract_BR.cbfExtra2Byte} <> 9 ))"        '10-20-05 send  detail records & total research data reqd except ntr & sports comments
                    End If
                End If
            Else                                'BR summary
                gUnpackDate igNowDate(0), igNowDate(1), slDate
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
                '9-14-09 no longer using lgNowTime for the time, getting milliseconds in time
                slTime = Trim$(str(lgNowTime))

                slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                
                'slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & slTime          '9-14-09

                '2-16-13 filter on gen date, time & urfcode
                slUserID = Trim$(str(tgUrf(0).iCode))
                slSelection = slSelection & " and {CBF_Contract_BR.cbfurfCode} = " & slUserID

                'If igJobRptNo = 4 Then              'NTR summary
                If igJobRptNo = 2 Then                  '10-30-07 NTR summary
                    slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 4)"        'send only NTR records
                ElseIf igJobRptNo = 3 Then         '1-06-21 CPM line IDs
                    slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 9)"        'send only CPM records
                Else                                'any summary other than NTR
                    If Not RptSelCt!ckcSelC6(1).Value = vbChecked Then          'dont include research
                                   'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0 or {CBF_Contract_BR.cbfExtra2Byte} = 6)"        'send only detail records or NTR installment, no total research data reqd
                        '5-2-14 remove including cbfextra2byte = -1
                        'billing summary
                        slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = 0) "        'send only the record so that the airtime and NTR subreports can link to it
                               '7-12-10 plus the detail records which contain the monthly summaries
                    Else
                        'Research is included, need to know which version of the summary will be printed
                        'Both Research version and the billing summary come thru here
                        If igBRSumZer Then              'its the billing summary
                          'slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} = -1 or {CBF_Contract_BR.cbfExtra2Byte} = 0) "        'send only the record so that the airtime and NTR subreports can link to it
                          '5-2-14 remove inclusion of cbfexra2byte = -1
                          slSelection = slSelection & " And ( {CBF_Contract_BR.cbfExtra2Byte} = 0) "        'send only the record so that the airtime and NTR subreports can link to it
                                        '7-12-10 plus the detail records which contain the monthly summaries
                        Else
                            slSelection = slSelection & " And ({CBF_Contract_BR.cbfExtra2Byte} <> 4 and {CBF_Contract_BR.cbfExtra2Byte} <> 5 and {CBF_Contract_BR.cbfExtra2Byte} <> -1 and {CBF_Contract_BR.cbfExtra2Byte} <> 8 and {CBF_Contract_BR.cbfExtra2Byte} <> 9 and {CBF_Contract_BR.cbfExtra2Byte} <> 10 and  {CBF_Contract_BR.cbfExtra2Byte} <> 11 )"        '10-20-05 send  detail records & total research data reqd except ntr & sports comments
                        End If
                    End If
                End If
            End If
            
            '8410 - on br.rpt ignore cbfLineTypes of X and Y (which are flagged hidden vehicles on a Installement contract)
            If ilListIndex = CNT_BR Then
                If igJobRptNo = 1 Then
                    slSelection = slSelection & " And ({CBF_Contract_BR.cbfLineType} <> 'X' and {CBF_Contract_BR.cbfLineType} <> 'Y') "
                End If
                If igJobRptNo = 4 Then
                    slSelection = slSelection & " And {CBF_Contract_BR.cbfLineType} <> 'X' "
                End If
                If igJobRptNo = 5 Then
                    slSelection = slSelection & " And {CBF_Contract_BR.cbfLineType} <> 'X' "
                End If
            End If
            If igJobRptNo <> 1 Then     '2-13-04 make sure all summaries get the splits if requested
                If ilListIndex = CNT_BR Then                'this doesnt apply for Insertion Orders
                    '2-2-10  Summary, is the version to merge the NTR billing?
'                    If igBRSumZer Or igBRSum Then           'bill summary and research summary have options to combine the NTR with air time totals
                    If igJobRptNo = 5 Or igJobRptNo = 4 Then                      'billing summary or research summary                        If RptSelCt!ckcSelC10(1).Value = vbChecked Then
                        If RptSelCt!ckcSelC10(1).Value = vbChecked Then                  '1-6-21
                            If Not gSetFormula("ShowNTRSummary", "'Y'") Then
                                mCntJob1_10 = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("ShowNTRSummary", "'N'") Then
                                mCntJob1_10 = -1
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If RptSelCt!ckcSelC10(0).Value = vbChecked Then           'show Slsp Commission Splits
                        If Not gSetFormula("ShowSplits", "'Y'") Then  'show the slsp comm splits on summary
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowSplits", "'N'") Then  'show the slsp comm splits on summary
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                    'on proposals, show the agy comm and net, previously only showed the gross $
                    If RptSelCt!ckcSelC13(0).Value = vbChecked Then       'show net amt on props
                        If Not gSetFormula("UserWantsNet", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("UserWantsNet", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                    
                    '10-31-19 show EDI agy client and adv product code.  show them on summary versions only
                    If ((Asc(tgSaf(0).sFeatures6) And EDIAGYCODES) = EDIAGYCODES) Then    'using agy client code?
                        If Not gSetFormula("ShowEDICodes", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowEDICodes", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                End If
            Else                        'detail, if contracts/proposal (not Insertion order), see if comments should be printed
                If ilListIndex = CNT_BR Then
                    '4-30-13  Show flight rates on packages (vs just show the total line
                    If (Asc(tgSpf.sUsingFeatures10) And PKGLNRATEONBR) = PKGLNRATEONBR Then     'show flight rates with package lines
                        If Not gSetFormula("ShowRateForPkgFlight", "'Y'") Then
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowRateForPkgFlight", "'N'") Then
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                
                    If ((Asc(tgSpf.sUsingFeatures8) And SHOWCMMTONDETAILPAGE) = SHOWCMMTONDETAILPAGE) Then      'yes, show comments on detail
                        If Not gSetFormula("ShowComments", "'Y'") Then  'show comments:  other,system site, chg reason, cancellations reason)
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowComments", "'N'") Then  'hide comments:  other,system site, chg reason, cancellations reason)
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If

                    'on proposals, show the agy comm and net, previously only showed the gross $
                    If RptSelCt!ckcSelC13(0).Value = vbChecked Then       'show net amt on props
                        If Not gSetFormula("UserWantsNet", "'Y'") Then  'show the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("UserWantsNet", "'N'") Then  'omit the net and agy comm along with gross $ on any proposals
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                    
                End If
            End If
            
            'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
            If ilListIndex = CNT_BR Then
                If RptSelCt!ckcSuppressNTRDetails.Value = vbChecked Then
                    If Not gSetFormula("SuppressNTRDetail", "true", False) Then
                        'Not going to fail if the formula is not present to set
                        'mCntJob1_10 = -1
                        'Exit Function
                    End If
                Else
                    If Not gSetFormula("SuppressNTRDetail", "false", False) Then
                        'Not going to fail if the formula is not present to set
                        'mCntJob1_10 = -1
                        'Exit Function
                    End If
                End If
            End If
            
            '8-25-15 move to common with contracts/proposals
            '5-23-13 Show Prod Protection code
            If RptSelCt!ckcSelC13(1).Value = vbChecked Then
                If Not gSetFormula("ShowProdProt", "'Y'") Then  'show product protection under Sales Office box in header
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowProdProt", "'N'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If
            'Fix v81 TTP 10745 testing results, Issue 4 - It looks like this change might have re-opened TTP 10537
            'If Not RptSelCt!ckcSelC6(1).Value = vbChecked Or igJobRptNo = 3 Or igJobRptNo = 4 Then                     'exclude research, do form that may or maynot include rates.  if CPM version, may or may not need to show the ratesThen
            If Not RptSelCt!ckcSelC6(1).Value = vbChecked Or igJobRptNo = 2 Or igJobRptNo = 3 Or igJobRptNo = 4 Then                     'exclude research, do form that may or maynot include rates.  if CPM version, may or may not need to show the ratesThen
                If RptSelCt!ckcSelC6(0).Value = vbChecked Then            'include rates
                    If Not gSetFormula("ShowRates", "'Y'") Then  'include rates with research (which includes weekly totals)
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowRates", "'N'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
                                    
            ElseIf ilListIndex = CNT_INSERTION Then                         'Research version
                If RptSelCt!ckcSelC6(0).Value = vbChecked Then            'include rates
                    If Not gSetFormula("ShowRates", "'Y'") Then  'include rates with research (which includes weekly totals)
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowRates", "'N'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
            
                            
            End If

            If RptSelCt!ckcSelC6(2).Value = vbChecked Then                'proof?
                If Not gSetFormula("Proof", "'Y'") Then     'Requesting hidden lines
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("Proof", "'N'") Then     'Normal BR
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If
            
            '4-26-12 all insertion orders and contracts to test for vehicle word wrap
            If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
                If Not gSetFormula("WordWrapVehicle", "'Y'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("WordWrapVehicle", "'N'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If

            If ilListIndex = CNT_INSERTION Then                'this doesnt apply for BR, only Insrtion

                If Not RptSelCt!ckcSelC6(1).Value = vbChecked Then            'exclude research, do form that may or maynot include rates
                    If RptSelCt!ckcSelC6(0).Value = vbChecked Then            'include rates
                        If Not gSetFormula("ShowRates", "'Y'") Then  'include rates with research (which includes weekly totals)
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                        
                        '9-22-15
                        If (Asc(tgSaf(0).sFeatures3) And SUPPRESSNETCOMM) = SUPPRESSNETCOMM Then 'Suppress Net and Commision on insertion orders
                            If Not gSetFormula("SuppressNet+Comm", "'Y'") Then  'suppress the comm + Net values on Insertions that show rates
                                mCntJob1_10 = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("SuppressNet+Comm", "'N'") Then  'suppress the comm + Net values on Insertions that show rates
                                mCntJob1_10 = -1
                                Exit Function
                            End If
                        End If

                    Else
                        If Not gSetFormula("ShowRates", "'N'") Then
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                End If
                If RptSelCt!ckcSelC10(0).Value = vbChecked Then           'show Slsp Commission Splits
                    If Not gSetFormula("NetNet", "'Y'") Then  'show net net values and change text to Participant for total lines
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("NetNet", "'N'") Then  'dont show net net values
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
                '10-24-08 option to include/exclude NTR
                If RptSelCt!ckcSelC12(0).Value = vbChecked Then           'show NTRs
                    If Not gSetFormula("ShowNTR", "'Y'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ShowNTR", "'N'") Then  'exclude NTR
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If

                '6/9/15: replaced acquisition from site override with Barter in system options
                '9-22-05 pass site options using acquisition.  Parse out bit
                'if using acquisition and research not requested, tell Crystal not to calculate any agy commissions

                'If ((Asc(tgSpf.sOverrideOptions) And SPACQUISITION) = SPACQUISITION) And (RptSelCt!ckcSelC6(1).Value = vbUnchecked) Then
                If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) And (RptSelCt!ckcSelC6(1).Value = vbUnchecked) Then
                    'using acquisitions, override with spot price if $0 acq?
                    If ((Asc(tgSaf(0).sFeatures1) And SHOWPRICEONINSERTIONWITHACQUISTION) = SHOWPRICEONINSERTIONWITHACQUISTION) Then ' Show Spot Prices on Insertion Orders if $0 Acquistion Exist
'                        'show $0 acq, if it exists
'                        If Not gSetFormula("CalcAgyComm", "'Y'") Then  'Override with spot rate if $0 acq; need to calc comm on those (if applicable). using acquistion costs
'                                                                        'and  stations are paid the amount on the Insertion order; no comm taken out
'                            mCntJob1_10 = -1
'                            Exit Function
'                        End If
'                    Else
                        '9-22-15 for Insertion orders, commission calculated internal in code because of varying rep commission structure
                        If Not gSetFormula("CalcAgyComm", "'N'") Then  'no agy comm for this version of the insertion order.  using acquistion costs
                                                                        'and  stations are paid the amount on the Insertion order; no comm taken out
                            mCntJob1_10 = -1
                            Exit Function
                        End If
                    End If
                Else            'using acq and research version; calc comm as required
                    If Not gSetFormula("CalcAgyComm", "'Y'") Then  'Use appropirate comm for this version of the insertion order.  No acquistion
                                                                'costs are shown
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
                
'                'move to common with contracts/proposals
'                '5-23-13 Show Prod Protection code
'                If RptSelCt!ckcSelC13(1).Value = vbChecked Then
'                    If Not gSetFormula("ShowProdProt", "'Y'") Then  'show product protection under Sales Office box in header
'                        mCntJob1_10 = -1
'                        Exit Function
'                    End If
'                Else
'                    If Not gSetFormula("ShowProdProt", "'N'") Then
'                        mCntJob1_10 = -1
'                        Exit Function
'                    End If
'                End If
                
                '5-23-13 Replace Agency phone # with Agency Name if using N & A from Site
                'if using site, it has been defaulted to "No", unchecked
                 If RptSelCt!ckcInclZero.Value = vbChecked Then
                    If Not gSetFormula("ReplAgyPhone", "'Y'") Then  'Replace agency phone # w/agy name
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ReplAgyPhone", "'N'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
                
                If RptSelCt!ckcSelC5(0).Value = vbChecked Then
                    If Not gSetFormula("DiffPlusCurrent", "'Y'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("DiffPlusCurrent", "'N'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
                
                '1-7-19 billing summary was implemented on the Insertion Order for a client; some didnt want them so use Site feature
                If (Asc(tgSaf(0).sFeatures6) And BILLINGONINSERTIONS) = BILLINGONINSERTIONS Then 'Insertion Order include Monthly Billed Summary
                    'show billing info
                    If Not gSetFormula("ShowHideBilling", "'S'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Else
                    'hide billing info
                    If Not gSetFormula("ShowHideBilling", "'H'") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                End If
            End If
        'End If

    ElseIf (ilListIndex = CNT_PAPERWORK) Then   'Contract Summary  (paperwork)
        slActive = ""
        slEntered = ""
        'Date: 12/19/2019 added CSI calendar control for date entries
        'If (RptSelCt!edcSelCFrom.Text <> "") And (RptSelCt!edcSelCFrom1.Text <> "") Then   'dates entered in both start & end
        If (RptSelCt!CSI_CalFrom.Text <> "") And (RptSelCt!CSI_CalTo.Text <> "") Then    'dates entered in both start & end
            If StrComp(RptSelCt!edcSelCFrom1.Text, "TFN", 1) <> 0 Then  'tfn entered for end date?
                slDate = RptSelCt!CSI_CalFrom.Text          ' edcSelCFrom.Text
                If gValidDate(slDate) Then
                    slDate = RptSelCt!CSI_CalTo.Text        ' edcSelCFrom1.Text
                    If gValidDate(slDate) Then
                        slDate = RptSelCt!CSI_CalTo.Text    ' edcSelCFrom1.Text
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slStr = Format$(gDateValue(slDate), "m/d/yy")
                        'slSelection = "{CHF_Contract_Header.chfStartDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        slDate = RptSelCt!CSI_CalFrom.Text  ' edcSelCFrom.Text
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slActive = Format$(gDateValue(slDate), "m/d/yy") & " - " & slStr
                        'slSelection = slSelection & " And {CHF_Contract_Header.chfEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    Else
                        mReset
                        RptSelCt!CSI_CalTo.SetFocus         ' edcSelCFrom1.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus           ' edcSelCFrom.SetFocus
                    Exit Function
                End If
            Else
                slDate = RptSelCt!CSI_CalFrom.Text          ' edcSelCFrom.Text
                If Not gValidDate(slDate) Then
                   ' gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    'slSelection = "{CHF_Contract_Header.chfEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'Else
                    mReset
                    RptSelCt!CSI_CalFrom.SetFocus           ' edcSelCFrom.SetFocus
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!CSI_CalFrom.Text <> "" Then         ' edcSelCFrom.Text <> "" Then     'active start date entered, no active end date
            slDate = RptSelCt!CSI_CalFrom.Text              ' edcSelCFrom.Text
            If gValidDate(slDate) Then
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slActive = "from " & Format$(gDateValue(slDate), "m/d/yy")
                'slSelection = "{CHF_Contract_Header.chfEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            Else
                mReset
                RptSelCt!CSI_CalFrom.SetFocus               ' edcSelCFrom.SetFocus
                Exit Function
            End If
        ElseIf RptSelCt!CSI_CalTo.Text <> "" Then           ' edcSelCFrom1.Text <> "" Then    'only active end date entered
            'If StrComp(RptSelCt!edcSelCFrom1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_CalTo.Text            ' edcSelCFrom1.Text
                If gValidDate(slDate) Then
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slActive = "thru " & Format$(gDateValue(slDate), "m/d/yy")
                    'slSelection = "{CHF_Contract_Header.chfStartDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    mReset
                    RptSelCt!CSI_CalTo.SetFocus             ' edcSelCFrom1.SetFocus
                    Exit Function
                End If
            End If
        Else
            slActive = "all dates"
        End If
        'If (RptSelCt!edcSelCTo.Text <> "") And (RptSelCt!edcSelCTo1.Text <> "") Then
        If (RptSelCt!CSI_From1.Text <> "") And (RptSelCt!CSI_To1.Text <> "") Then
            'If StrComp(RptSelCt!edcSelCTo1.Text, "TFN", 1) <> 0 Then
            If StrComp(RptSelCt!CSI_From1.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_From1.Text                ' edcSelCTo.Text
                If gValidDate(slDate) Then
                    slDate = RptSelCt!CSI_From1.Text            ' edcSelCTo1.Text
                    If gValidDate(slDate) Then
                        slDate = RptSelCt!CSI_From1.Text        ' edcSelCTo.Text
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slEntered = Format$(gDateValue(slDate), "m/d/yy")
                        'If slSelection = "" Then
                        '    slSelection = "{CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                       ' Else
                        '    slSelection = slSelection & " And {CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        'End If
                        slDate = RptSelCt!CSI_To1.Text          ' edcSelCTo1.Text
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slEntered = slEntered & " - " & Format$(gDateValue(slDate), "m/d/yy")
                        'slSelection = slSelection & " And {CHF_Contract_Header.chfOHDDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    Else
                        mReset
                        RptSelCt!CSI_To1.SetFocus               ' edcSelCTo1.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSelCt!CSI_From1.SetFocus                 ' edcSelCTo.SetFocus
                    Exit Function
                End If
            Else
                slDate = RptSelCt!CSI_From1.Text                ' edcSelCTo.Text
                If Not gValidDate(slDate) Then
                    'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    'If slSelection = "" Then
                    '    slSelection = "{CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    'Else
                    '    slSelection = slSelection & "And {CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    'End If
                'Else
                    mReset
                    RptSelCt!CSI_From1.SetFocus                 ' edcSelCTo.SetFocus
                    Exit Function
                End If
            End If
        ElseIf RptSelCt!CSI_From1.Text <> "" Then               ' edcSelCTo.Text <> "" Then
            slDate = RptSelCt!CSI_From1.Text                    ' edcSelCTo.Text
            If gValidDate(slDate) Then
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slEntered = "from " & Format$(gDateValue(slDate), "m/d/yy")
                'If slSelection = "" Then
                '    slSelection = "{CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'Else
                '    slSelection = slSelection & " And {CHF_Contract_Header.chfOHDDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'End If
            Else
                mReset
                RptSelCt!CSI_From1.SetFocus                     ' edcSelCTo.SetFocus
                Exit Function
            End If
        ElseIf RptSelCt!CSI_To1.Text <> "" Then                 ' edcSelCTo1.Text <> "" Then
            If StrComp(RptSelCt!edcSelCTo1.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCt!CSI_To1.Text                  '  edcSelCTo1.Text
                If gValidDate(slDate) Then
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slEntered = "thru " & Format$(gDateValue(slDate), "m/d/yy")
                    'If slSelection = "" Then
                    '    slSelection = "{CHF_Contract_Header.chfOHDDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    'Else
                    '    slSelection = slSelection & " And {CHF_Contract_Header.chfOHDDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    'End If
                Else
                    mReset
                    RptSelCt!CSI_To1.SetFocus                   ' edcSelCTo1.SetFocus
                    Exit Function
                End If
            End If
        Else
            slEntered = "all dates"
        End If
'Date: 9/10/2018    radio buttons replaced with dropdown list
'        If RptSelCt!rbcSelC9(0).Value Then
'            If Not gSetFormula("SortBy", "'A'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!rbcSelC9(1).Value Then
'            If Not gSetFormula("SortBy", "'G'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!rbcSelC9(2).Value Then
'            If Not gSetFormula("SortBy", "'S'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        Else
'            If Not gSetFormula("SortBy", "'V'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        End If

        'TTP 10520 - Paperwork Summary report by vehicle may show incorrect gross when run for selective vehicles
        If RptSelCt!rbcSelCSelect(3).Value = True Then
            'By Vehicle (Set so that Hidden Lines will be totaled)
            If Not gSetFormula("TotalsInclHidden", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else
            'Any other option (Allow Hidden lines to not be totaled)
            If Not gSetFormula("TotalsInclHidden", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        
        'Sort #1
        Select Case RptSelCt!cbcSort1.ListIndex
        Case 0
            If Not gSetFormula("SortBy", "'D'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 1
            If Not gSetFormula("SortBy", "'A'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 2
            'ElseIf RptSelCt!rbcSelC9(1).Value Then
            If Not gSetFormula("SortBy", "'G'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 3
            'ElseIf RptSelCt!rbcSelC9(2).Value Then
            If Not gSetFormula("SortBy", "'S'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 4
            If Not gSetFormula("SortBy", "'V'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End Select

        'Sort #2
        Select Case RptSelCt!cbcSort2.ListIndex
        Case 0
            If Not gSetFormula("SortBy2", "'None'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 1
            If Not gSetFormula("SortBy2", "'D'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 2
            If Not gSetFormula("SortBy2", "'A'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 3
            If Not gSetFormula("SortBy2", "'G'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 4
            If Not gSetFormula("SortBy2", "'S'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Case 5
            If Not gSetFormula("SortBy2", "'V'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End Select
        
        'Date: 9/14/2018    added SHOW: GRP or ORDERHOLD DATE
        If RptSelCt!rbcShow(0).Value = vbTrue Then            'show GRP
            If Not gSetFormula("ShowGRP", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
            If Not gSetFormula("ShowExtCntrNo", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        
        ElseIf RptSelCt!rbcShow(1).Value = vbTrue Then       'show OrderHold Date
            If Not gSetFormula("ShowGRP", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
            If Not gSetFormula("ShowExtCntrNo", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
            
        'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
        ElseIf RptSelCt!rbcShow(2).Value = vbTrue Then       'TTP 10721 - Paperwork Summary report: add radio button next to "Show GRP/Prop Date" to show the External Contract number
            If Not gSetFormula("ShowExtCntrNo", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        
            If Not gSetFormula("ShowGRP", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        
        
        End If

        If RptSelCt!ckcSelC10(0).Value = vbChecked Then         'show rates
            If Not gSetFormula("ShowRates", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else                                        'dont show any rates
            If Not gSetFormula("ShowRates", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If

        If Not gSetFormula("ActiveDates", "'" & slActive & "'") Then    'added 3/29/99
            mCntJob1_10 = -1
            Exit Function
        End If
        If Not gSetFormula("EnteredDates", "'" & slEntered & "'") Then
            mCntJob1_10 = -1
            Exit Function
        End If

        If RptSelCt!ckcSelC12(0).Value = vbChecked Then
            If Not gSetFormula("SkipPage", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipPage", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If


        ilRet = mGrossOrNetHdr()            '11-4-13
        If ilRet = -1 Then
            mCntJob1_10 = -1
            Exit Function
        End If

'        If RptSelCt!rbcSelC7(0).Value Then
'            If Not gSetFormula("GrossOrNet", "'G'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        Else
'            If Not gSetFormula("GrossOrNet", "'N'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        End If

        If RptSelCt!rbcSelCInclude(0).Value Then        'Contract (vs line)
            If RptSelCt!ckcSelC13(0).Value = vbChecked Then         '6-14-02
                If Not gSetFormula("ShowComm%", "'Y'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowComm%", "'N'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If
        End If


        slExclude = ""
        slInclude = ""
        gIncludeExcludeCkc RptSelCt!ckcSelC8(0), slInclude, slExclude, "Discrep Only"
        gIncludeExcludeCkc RptSelCt!ckcSelC8(1), slInclude, slExclude, "Credit Check Only"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Orders"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(6), slInclude, slExclude, "Rev"            '11-7-16
        gIncludeExcludeCkc RptSelCt!ckcSelC3(2), slInclude, slExclude, "Rejected"       '4-29-09 chged from dead
        gIncludeExcludeCkc RptSelCt!ckcSelC3(3), slInclude, slExclude, "Working"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(4), slInclude, slExclude, "Unapproved"
        gIncludeExcludeCkc RptSelCt!ckcSelC3(5), slInclude, slExclude, "Complete"

        gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(5), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelCt!ckcSelC5(6), slInclude, slExclude, "Promo"
        '12-7-04 Cancel before Start
        gIncludeExcludeCkc RptSelCt!ckcSelC6(4), slInclude, slExclude, "CBS"
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If

    'ElseIf (rbcRptType(4).Value) Then
    ElseIf (ilListIndex = CNT_BOB_BYCNT) Then   'Contract Business Booked by contract (formerly projection)
        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/16/2019 added CSI calendar control for date entry --> edcSelCFrom.Text
        If gValidDate(slDate) Then
            If RptSelCt!rbcSelCSelect(0).Value Then  'set last sunday of first week
                slDate = gObtainNextSunday(slDate)
                slDate = gDecOneWeek(slDate)   'get previous end week
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If

                For illoop = 1 To 12 Step 1
                    slDate = gIncOneWeek(slDate)
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                Next illoop
                If Not gSetFormula("Type", "'Weekly,'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If

                'Report!crcReport.Formulas(0) = "Wk 1= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
            ElseIf RptSelCt!rbcSelCSelect(1).Value Then  'set last date of 12 standard periods
                If Not gSetFormula("Type", "'Standard Month,'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(0) = "Type= Standard Month,"
                slDate = gObtainStartStd(slDate)
                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'End date of previous month
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(1) = "P0= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                For illoop = 1 To 12 Step 1
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    slDate = gObtainEndStd(slDate)
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            ElseIf RptSelCt!rbcSelCSelect(2).Value Then  'set last date of 12 corporate periods
                If Not gSetFormula("Type", "'Corporate Month,'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(0) = "Type= Corporate Month,"
                slDate = gObtainStartCorp(slDate, True)
                slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'End date of previous month
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(1) = "P0= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                For illoop = 1 To 12 Step 1
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    slDate = gObtainEndCorp(slDate, True)
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mCntJob1_10 = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            End If
            If RptSelCt!rbcSelC4(0).Value Then        'gross
                If Not gSetFormula("GrossNetOrProd", "'G'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            ElseIf RptSelCt!rbcSelC4(1).Value Then    'net
                If Not gSetFormula("GrossNetOrProd", "'N'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            ElseIf RptSelCt!rbcSelC4(2).Value Then    'netnet
                If Not gSetFormula("GrossNetOrProd", "'P'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If
        Else
            mReset
            RptSelCt!CSI_CalFrom.SetFocus       'Date: 12/16/2019 added CSI calendar control for date entry --> edcSelCFrom.SetFocus
            Exit Function
        End If
    ElseIf ilListIndex = CNT_RECAP Then 'Recap
        If RptSelCt!edcSelCFrom.Text <> "" Then
            slStr = RptSelCt!edcSelCFrom.Text
        Else
            slStr = "0"
        End If
        slSelection = "({CLF_Contract_Line.clfType} <> 'H') and {CHF_Contract_Header.chfCntrNo} >= " & slStr
        If RptSelCt!edcSelCTo.Text <> "" Then
            slStr = RptSelCt!edcSelCTo.Text
        Else
            slStr = "2147483647"
        End If
        slSelection = slSelection & " And " & "{CHF_Contract_Header.chfCntrNo} <= " & slStr
    'Spot Placement code removed--see rptselcb
    'Spot Discrepancies code removed--see rptselcb

    ElseIf ilListIndex = CNT_MG Then 'MG's

        slLocalFeed = "({SMF_Spot_MG_Specs.smfchfCode}> 0)"


''**CCCCC

'

'        If slMissedStart <> "" Then
'            If gValidDate(slMissedStart) Then
'                gObtainYearMonthDayStr slMissedStart, True, slYear, slMonth, slDay
'                slSelection = "{SMF_Spot_MG_Specs.smfMissedDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'            Else
'                mReset
'                RptSelCt!edcSelCFrom.SetFocus
'                Exit Function
'            End If
'        End If

'        If slMissedEnd <> "" Then
'            If gValidDate(slMissedEnd) Then
'                gObtainYearMonthDayStr slMissedEnd, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{SMF_Spot_MG_Specs.smfMissedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfMissedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCFrom1.SetFocus
'                Exit Function
'            End If
'        End If
'    '    Test entered makegood dates

'        If slMGStart <> "" Then
'            If gValidDate(slMGStart) Then
'                gObtainYearMonthDayStr slMGStart, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{SMF_Spot_MG_Specs.smfActualDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfActualDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCTo.SetFocus
'                Exit Function
'            End If
'        End If

'        If slMGEnd <> "" Then
'            If gValidDate(slMGEnd) Then
'                gObtainYearMonthDayStr slMGEnd, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{SMF_Spot_MG_Specs.smfActualDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfActualDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCTo1.SetFocus
'                Exit Function
'            End If
'        End If
'        If slSelection = "" Then
'            slSelection = "{SDF_Spot_Detail.sdfSpotType} <> 'X' and  {SDF_Spot_Detail.sdfPriceType} <> 'N'"
'        Else
'            slSelection = slSelection & " And " & "{SDF_Spot_Detail.sdfSpotType} <> 'X' and {SDF_Spot_Detail.sdfPriceType} <> 'N'"
'        End If
'        If slSelection <> "" Then
'            slInclStatus = " and ("
'        Else
'            slInclStatus = "("
'        End If
'        slOr = ""
'        slStr = ""
'        If RptSelCt!ckcSelC3(0).Value = vbChecked Then          'include makegoods
'            slInclStatus = slInclStatus & "{SMF_Spot_MG_Specs.smfSchStatus} = 'G'"
'            slOr = " or "
'            slStr = "Makegoods"
'        End If
'        If RptSelCt!ckcSelC3(1).Value = vbChecked Then          'include outsides
'            slInclStatus = slInclStatus & slOr & " {SMF_Spot_MG_Specs.smfSchStatus} = 'O'"
'            If slStr = "" Then
'                slStr = "Outsides"
'            Else
'                slStr = slStr & " and Outsides"
'            End If
'        End If
'        If Not gSetFormula("MG&Out", "'" & slStr & "'") Then
'            mCntJob1_10 = -1
'            Exit Function
'        End If
'
'        If RptSelCt!rbcSelCSelect(0).Value Then           'select vehicles for missed
'            If Not gSetFormula("MissMGVeh", "'Vehicle selection for Missed Vehicles'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!rbcSelCSelect(1).Value Then           'select vehicles for mg/out
'            If Not gSetFormula("MissMGVeh", "'Vehicle selection for MG/Outside Vehicles'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!rbcSelCSelect(2).Value Then           'select vehicles for missed and mg/out
'            If Not gSetFormula("MissMGVeh", "'Vehicle selection for either Missed or MG/Outside Vehicles'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        Else                                                'select vehicles for either missed or mg/out
'            If Not gSetFormula("MissMGVeh", "'Vehicle selection for both Missed and MG/Outside Vehicles'") Then
'                mCntJob1_10 = -1
'                Exit Function
'            End If
'        End If
'
'        slInclStatus = slInclStatus & ")"
'
'
'        slSelection = slSelection & slInclStatus
'        If Not RptSelCt!ckcAll.Value = vbChecked Then         'not all vehicles selected
'            If slSelection <> "" Then
'                slSelection = "(" & slSelection & ") " & " and ("
'                slOr = ""
'            Else
'                slSelection = "("
'                slOr = ""
'            End If
'            'setup selective vehicles
'            For ilLoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
'                If RptSelCt!lbcSelection(6).Selected(ilLoop) Then
'                    slNameCode = tgCSVNameCode(ilLoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
'                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                    If RptSelCt!rbcSelCSelect(0).Value Then             'check missed vehicle
'                        slSelection = slSelection & slOr & "{SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
'                    ElseIf RptSelCt!rbcSelCSelect(1).Value Then           'match on mg vehicle                                              'check mg/out vehicle
'                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode)
'                    ElseIf RptSelCt!rbcSelCSelect(2).Value Then           'match on either missed or mg vehicle                                              'check mg/out vehicle
'                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode) & " or " & " {SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
'                    Else                                                'match on both missed or mg vehicle
'                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode) & " and " & " {SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
'                    End If
'                    slOr = " Or "
'                End If
'            Next ilLoop
'            slSelection = slSelection & ")"
'
'        End If


        slMissedStart = RptSelCt!CSI_CalFrom.Text           'Date: 1/7/2020 added CSI controls for date entries --> edcSelCFrom.Text
        slMissedEnd = RptSelCt!CSI_CalTo.Text               'Date: 1/7/2020 added CSI controls for date entries --> edcSelCFrom1.Text
        slMGStart = RptSelCt!CSI_From1.Text                 'Date: 1/7/2020 added CSI controls for date entries --> edcSelCTo.Text
        slMGEnd = RptSelCt!CSI_To1.Text                     'Date: 1/7/2020 added CSI controls for date entries --> edcSelCTo1.Text

        gObtainYearMonthDayStr Now, True, slYear, slMonth, slDay
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

        slSelection = slSelection & " and " & slLocalFeed           '7-20-04 exclude network (feed spots)
        'send formulas for report dates
        If slMissedStart = "" And slMissedEnd = "" Then
            If Not gSetFormula("MissedDates", "'Missed Dates: All'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf slMissedStart = "" Then          'start date blank
            slInclude = "Missed Dates: thru "
            slMissedEnd = Format$(gDateValue(slMissedEnd), "m/d/yy")
            If Not gSetFormula("MissedDates", "'" & slInclude & slMissedEnd & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf slMissedEnd = "" Then            'end date blank
            slInclude = "Missed Dates: from "
            slMissedStart = Format$(gDateValue(slMissedStart), "m/d/yy")
            If Not gSetFormula("MissedDates", "'" & slInclude & slMissedStart & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else
            slInclude = "Missed Dates: "
            slMissedStart = Format$(gDateValue(slMissedStart), "m/d/yy")
            slMissedEnd = Format$(gDateValue(slMissedEnd), "m/d/yy")
            If Not gSetFormula("MissedDates", "'" & slInclude & slMissedStart & "-" & slMissedEnd & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        If slMGStart = "" And slMGEnd = "" Then
            If Not gSetFormula("MGDates", "'MG Dates: All'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf slMGStart = "" Then          'start date blank
            slInclude = "MG Dates: thru "
            slMGEnd = Format$(gDateValue(slMGEnd), "m/d/yy")
            If Not gSetFormula("MGDates", "'" & slInclude & slMGEnd & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf slMGEnd = "" Then            'end date blank
            slInclude = "MG Dates: from "
            slMGStart = Format$(gDateValue(slMGStart), "m/d/yy")
            If Not gSetFormula("MGDates", "'" & slInclude & slMGStart & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else
            slInclude = "MG Dates: "
            slMGStart = Format$(gDateValue(slMGStart), "m/d/yy")
            slMGEnd = Format$(gDateValue(slMGEnd), "m/d/yy")
            If Not gSetFormula("MGDates", "'" & slInclude & slMGStart & "-" & slMGEnd & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
    ElseIf ilListIndex = CNT_SPOTTRAK Then 'Sales Spot Tracking

'        slDate = RptSelCt!edcSelCFrom.Text
'        If slDate <> "" Then
'            If gValidDate(slDate) Then
'                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                slSelection = "{STF_Spot_Tracking.stfCreateDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'            Else
'                mReset
'                RptSelCt!edcSelCFrom.SetFocus
'                Exit Function
'            End If
'        End If
'        slDate = RptSelCt!edcSelCFrom1.Text
'        If slDate <> "" Then
'            If gValidDate(slDate) Then
'                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfCreateDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfCreateDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCFrom1.SetFocus
'                Exit Function
'            End If
'        End If
'        slDate = RptSelCt!edcSelCTo.Text
'        If slDate <> "" Then
'            If gValidDate(slDate) Then
'                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfLogDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfLogDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCTo.SetFocus
'                Exit Function
'            End If
'        End If
'        slDate = RptSelCt!edcSelCTo1.Text
'        If slDate <> "" Then
'            If gValidDate(slDate) Then
'                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfLogDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfLogDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                End If
'            Else
'                mReset
'                RptSelCt!edcSelCTo1.SetFocus
'                Exit Function
'            End If
'        End If
'        'Bypass deleted spots
'        If (RptSelCt!ckcSelC3(0).Value = vbUnchecked) Or (RptSelCt!ckcSelC3(1).Value = vbUnchecked) Or (RptSelCt!ckcSelC3(2).Value = vbUnchecked) Then
'            If (RptSelCt!ckcSelC3(0).Value = vbChecked) Then
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfPrint} = 'R'"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfPrint} = 'R'"
'                End If
'            End If
'            If (RptSelCt!ckcSelC3(1).Value = vbChecked) Then
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfPrint} = 'P'"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfPrint} = 'P'"
'                End If
'            End If
'            If (RptSelCt!ckcSelC3(2).Value = vbChecked) Then
'                If slSelection = "" Then
'                    slSelection = "{STF_Spot_Tracking.stfPrint} = 'D'"
'                Else
'                    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfPrint} = 'D'"
'                End If
'            End If
'        End If

        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCFrom.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_CalFrom.SetFocus   'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_CalTo.Text        'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCFrom1.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalTo.SetFocus     'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCFrom1.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_From1.Text        'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCTo.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_From1.SetFocus     'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_To1.Text          'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCTo1.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_To1.SetFocus       'Date: 12/9/2019 added CSI calendar controls for date entries --> edcSelCTo1.SetFocus
                Exit Function
            End If
        End If
        ilRet = gGRFSelection(slSelection)      'build date & time filter to send to crystal
        If ilRet <> 0 Then
            mCntJob1_10 = -1
            Exit Function
        End If

        ilRet = mTrackAndComlChgDates(ilListIndex)     'send dates requested for report headings
        If ilRet <> 0 Then
            mCntJob1_10 = -1
            Exit Function
            End If
            
    ElseIf ilListIndex = CNT_COMLCHG Then 'Commercial Changes
        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalFrom.SetFocus   'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_CalTo.Text        'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom1.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_CalTo.SetFocus     'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCFrom1.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_From1.Text        'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_From1.SetFocus     'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCt!CSI_To1.Text          'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo1.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_To1.SetFocus       'Date: 12/9/2019 added CSI calendar controls for date entry --> edcSelCTo1.SetFocus
                Exit Function
            End If
        End If
        ilRet = gGRFSelection(slSelection)      'build date & time filter to send to crystal
        If ilRet <> 0 Then
            mCntJob1_10 = -1
            Exit Function
        End If
        
        ilRet = mTrackAndComlChgDates(ilListIndex)     'send dates requested for report headings
        If ilRet <> 0 Then
            mCntJob1_10 = -1
            Exit Function
        End If
    End If
    'paperwork, business booked (projection)
    'If (ilListIndex <> 8) And (ilListIndex <> 9) And (ilListIndex < 15) And (ilListIndex <> 0) Then
    If (ilListIndex = CNT_BOB_BYCNT) Then
        If igRptType = 0 Then                       'coming from proposals
            If slSelection = "" Then
                slSelection = "({CHF_Contract_Header.chfStatus} = 'W' Or {CHF_Contract_Header.chfStatus} = 'D' Or {CHF_Contract_Header.chfStatus} = 'C' Or {CHF_Contract_Header.chfStatus} = 'I')"
            Else
                slSelection = "(" & slSelection & ") " & " or "
                slSelection = slSelection & "({CHF_Contract_Header.chfStatus} = 'W' Or {CHF_Contract_Header.chfStatus} = 'D' Or {CHF_Contract_Header.chfStatus} = 'C' Or {CHF_Contract_Header.chfStatus} = 'I')"
            End If
        End If
    End If


    If (ilListIndex = CNT_BOB_BYCNT) Then   'Business Booked (formerly Projection)
        'ignore header, lines, flights that are deleted, any cancel before start, or hidden line.  Also ignore
        'flights whose end date is prior to the earliest date to include (P0)
        '6-24-04 ignore altered and interrupted schedule lines
        slSelection = "({CLF_Contract_Line.clfType} <> 'H' and {CLF_Contract_Line.clfSchStatus} <> 'A' and {CLF_Contract_Line.clfSchStatus} <> 'I' and {CHF_Contract_Header.chfDelete} <> 'Y' and {CLF_Contract_Line.clfDelete} <> 'Y' and {CFF_Contract_Flight.cffDelete} <> 'Y' and {CFF_Contract_Flight.cffEndDate} >= {CFF_Contract_Flight.cffStartDate} and {CFF_Contract_Flight.cffEndDate} >{@P0}) "
        slOr = ""
        If RptSelCt!ckcSelC5(1).Value = vbChecked Then            'includes orders
            slOr = "And ({CHF_Contract_Header.chfStatus} = 'O' or {CHF_Contract_Header.chfStatus} = 'N'"
        End If

        If RptSelCt!ckcSelC5(0).Value = vbChecked Then            'includes holds
            If slOr <> "" Then
                slOr = slOr & " Or {CHF_Contract_Header.chfStatus} = 'H' or {CHF_Contract_Header.chfStatus} = 'G'"
            Else
                slOr = " And ({CHF_Contract_Header.chfStatus} = 'H' or {CHF_Contract_Header.chfStatus} = 'G'"
            End If
        End If
        slSelection = slSelection & slOr & ")"
        If Not RptSelCt!ckcAll.Value = vbChecked Then
            If slSelection <> "" Then
                slSelection = "(" & slSelection & ") " & " And ("
                slOr = ""
            Else
                slSelection = "("
                slOr = ""
            End If
            If RptSelCt!rbcSelCInclude(0).Value Then 'Advertiser/Contracts
                ilAllSelected = True
                For illoop = 0 To RptSelCt!lbcSelection(0).ListCount - 1 Step 1
                    If Not RptSelCt!lbcSelection(0).Selected(illoop) Then
                        ilAllSelected = False
                        Exit For
                    End If
                Next illoop
                If ilAllSelected Then
                    For illoop = 0 To RptSelCt!lbcSelection(5).ListCount - 1 Step 1
                        If RptSelCt!lbcSelection(5).Selected(illoop) Then
                            slNameCode = tgAdvertiser(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            slSelection = slSelection & slOr & "{CHF_Contract_Header.chfadfCode} = " & Trim$(slCode)
                            slOr = " Or "
                        End If
                    Next illoop
                Else
                    For illoop = 0 To RptSelCt!lbcSelection(0).ListCount - 1 Step 1
                        If RptSelCt!lbcSelection(0).Selected(illoop) Then
                            slNameCode = RptSelCt!lbcCntrCode.List(illoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            slSelection = slSelection & slOr & "{CHF_Contract_Header.chfCode} = " & Trim$(slCode)
                            slOr = " Or "
                        End If
                    Next illoop
                End If
            ElseIf RptSelCt!rbcSelCInclude(1).Value Then 'Salesperson
                For illoop = 0 To RptSelCt!lbcSelection(2).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(2).Selected(illoop) Then
                        slNameCode = tgSalesperson(illoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                        slSelection = slSelection & slOr & "{CHF_Contract_Header.chfslfCode1} = " & Trim$(slCode)
                        slOr = " Or "
                    End If
                Next illoop
            ElseIf RptSelCt!rbcSelCInclude(2).Value Then 'Vehicle
                For illoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(6).Selected(illoop) Then
                        slNameCode = tgCSVNameCode(illoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                        slSelection = slSelection & slOr & "{CLF_Contract_Line.clfvefCode} = " & Trim$(slCode)
                        slOr = " Or "
                    End If
                Next illoop
            End If
            slSelection = slSelection & ")"
        End If
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    End If
    'If rbcRptType(0).Value Then
    If ilListIndex = CNT_BR Or ilListIndex = CNT_INSERTION Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    'ElseIf rbcRptType(1).Value Then
    ElseIf ilListIndex = CNT_PAPERWORK Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If

    ElseIf ilListIndex = CNT_RECAP Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If

    ElseIf ilListIndex = CNT_MG Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_SPOTTRAK Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    End If
    mCntJob1_10 = 1
    Exit Function
End Function
'
'
'           gGRFSelection - send filter selection to Crystal.
'           Filter with date & time generated in GRF.btr
'
'           <input> slSelection - string to build selection
'           <return> True = error
Function gGRFSelection(slSelection As String) As Integer
Dim slDate As String
Dim slTime As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String


    gGRFSelection = 0
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gGRFSelection = -1                      'send error return
    End If
End Function
'
'           mBobMonthheader - Determine the month/year and type of
'           month (std/corp) the Billed & Booked has been requested for
'           Report headings in Crystal
'
'           7-2-08 Change Billed and Booked, B & B Commissions, B & B Recap
'                  to enter year, starting month and # months instead of
'                  year, Qtr #, # periods
'
'       BILLED AND BOOKED
'       BILLED AND BOOKED RECAP
'       BILLED AND BOOKED COMPARISONS
'
Public Function mBOBMonthHeader()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLoop                        ilRet                     *
'*  ilTemp                        ilWhichQtr                                              *
'******************************************************************************************

Dim ilSaveMonth As Integer
Dim slMonthInYear As String * 36
Dim slMonth As String
Dim ilStartQtr As Integer
Dim ilYear As Integer



    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
    mBOBMonthHeader = 0

    ilStartQtr = gGetQtrForColumns(igMonthOrQtr)   'column headings for quarter totals

    If RptSelCt!rbcSelC9(0).Value Then          'corp option
        slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year

        ilYear = gGetYearofCorpMonth(igMonthOrQtr, igYear)
        'changed from starting qtr to allow for starting month and to show in report header
        'If Not gSetFormula("MonthHeader", "'" & slMonth & " " & Str$(ilYear) & "'") Then
        '12-21-12 Need to show the requested user entry, not the actual year of the month the corporate year starts in
        If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(igYear) & "'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If

        If Not gSetFormula("CorpStd", "'C'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If

        gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)

        If Not gSetFormula("StartingMonth", ilSaveMonth) Then       'pass starting month of the starting corp qtr for report column headings
            mBOBMonthHeader = -1
            Exit Function
        End If
    ElseIf RptSelCt!rbcSelC9(1).Value Then     'std
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month #
        'changed from starting qtr to allow for starting month and to show in report header
        If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(igYear) & "'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If

        If Not gSetFormula("CorpStd", "'S'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If

        'If Not gSetFormula("StartingMonth", ((ilLoop - 1) * 3 + 1)) Then         'pass starting month of the starting std qtr for report column headings
        If Not gSetFormula("StartingMonth", igMonthOrQtr) Then         'pass starting month of the starting std qtr for report column headings
            mBOBMonthHeader = -1
            Exit Function
        End If
    ElseIf RptSelCt!rbcSelC9(4).Value Then     'bill method
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month #
        'changed from starting qtr to allow for starting month and to show in report header
        If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(igYear) & "'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If
        If Not gSetFormula("CorpStd", "'B'") Then       'cal by cnt
            mBOBMonthHeader = -1
            Exit Function
        End If
        gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)

        If Not gSetFormula("StartingMonth", ilSaveMonth) Then       'pass starting month of the starting corp qtr for report column headings
            mBOBMonthHeader = -1
            Exit Function
        End If

    Else                                'calendar (calc by day)
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month #
        'changed from starting qtr to allow for starting month and to show in report header
        If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(igYear) & "'") Then
            mBOBMonthHeader = -1
            Exit Function
        End If

        If RptSelCt!rbcSelC9(2).Value Then
            If Not gSetFormula("CorpStd", "'D'") Then       'cal by cnt
                mBOBMonthHeader = -1
                Exit Function
            End If
        ElseIf RptSelCt!rbcSelC9(3).Value Then              'cal by spots
            If Not gSetFormula("CorpStd", "'P'") Then
                mBOBMonthHeader = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("CorpStd", "'B'") Then       'by Bill Cycle
                mBOBMonthHeader = -1
            Exit Function
            End If
        End If

        'If Not gSetFormula("StartingMonth", ((ilLoop - 1) * 3 + 1)) Then         'pass starting month of the starting std qtr for report column headings
        If Not gSetFormula("StartingMonth", igMonthOrQtr) Then        'pass starting month of the starting std qtr for report column headings
            mBOBMonthHeader = -1
            Exit Function
        End If
    End If
    If Not gSetFormula("QtrHeader", "'" & Trim$(str(ilStartQtr)) & "'") Then       'pass starting qtr for column headings
        mBOBMonthHeader = -1
        Exit Function
    End If
End Function

'               mBobTotals - determine what level of totals should
'               be shown on different versions of Billed & Booked
'               and send formula to Crystal report
Public Function mBobTotals()

    mBobTotals = 0
    'How to show total levels?
    If RptSelCt!rbcSelC4(0).Value Then              'totals by contract
        If Not gSetFormula("TotalLevel", "'C'") Then
            mBobTotals = -1
            Exit Function
        End If
    ElseIf RptSelCt!rbcSelC4(1).Value Then          'totals by advt
        If Not gSetFormula("TotalLevel", "'A'") Then
            mBobTotals = -1
            Exit Function
        End If
    Else                                            'totals by summary
        If Not gSetFormula("TotalLevel", "'S'") Then
            mBobTotals = -1
            Exit Function
        End If
    End If
End Function
'
'           Verify input parameters for Contract index starting at 38 (Paperwork Tax Summary)
'
Function mCntJob38Plus(ilListIndex As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilSaveMonth                                                                           *
'******************************************************************************************

Dim slDate As String
Dim slTime As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim slSelection As String
Dim llStartDate As Long
Dim llEndDate As Long
Dim slStr As String
Dim ilNoneExists As Integer
Dim ilMinorGroupHdr As Integer
Dim slMonthInYear As String * 36
Dim ilRet As Integer
Dim ilTemp As Integer
Dim ilVehicleGroup As Integer
Dim llUserStartDateSent As Long
Dim llUserEndDateSent As Long
Dim llUserActiveStartDate As Long
Dim llUserActiveEndDate As Long

    mCntJob38Plus = 0
    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"

    If ilListIndex = CNT_PAPERWORKTAX Then          '4-9-07
        slDate = RptSelCt!CSI_CalFrom.Text          'Date: 12/20/2019 added CSI calendar control for date entries --> edcSelCFrom.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
            Else
                mReset
                RptSelCt!CSI_CalFrom.SetFocus       'Date: 12/20/2019 added CSI calendar control for date entries --> edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        llStartDate = gDateValue(slDate)
        slDate = RptSelCt!CSI_CalTo.Text            'Date: 12/20/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelCt!CSI_CalTo.SetFocus         'Date: 12/20/2019 added CSI calendar control for date entries --> edcSelCFrom1.SetFocus
                Exit Function
            End If
        End If
        llEndDate = gDateValue(slDate)

        slStr = Format$(llStartDate, "m/d/yy") + " - " + Format$(llEndDate, "m/d/yy")
        If Not gSetFormula("DatesRequested", "'" & slStr & "'") Then
            mCntJob38Plus = -1
            Exit Function
        End If

        If RptSelCt!ckcSelC3(0).Value = vbChecked Then              'skip new page
            If Not gSetFormula("SkipPage", "'Y'") Then
                mCntJob38Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipPage", "'N'") Then
                mCntJob38Plus = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            mCntJob38Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_BOBCOMPARE Then
        'verify user input dates
        ilRet = mVerifyMonthYrPeriods(RptSelCt, ilListIndex, RptSelCt!rbcSelC9(0))        '7-3-08 allow starting month # vs starting qtr
        If ilRet = True Then            'got an error in conversion of input
            mCntJob38Plus = -1
            Exit Function
        End If

        ilRet = mBOBMonthHeader()       'format qtr/year & month heading
        If ilRet <> 0 Then
            mCntJob38Plus = -1
            Exit Function
        End If

        If mBOBCrystal() < 0 Then                                'send Crystl formulas for Header notations (pkg vs hidden),
            mCntJob38Plus = -1                                 'As of Time,  Gross, Net
            Exit Function
        End If

        'see if vehicle group selected, pass to put type in header (mkt, participants, formats, etc)
        ilTemp = RptSelCt!cbcSet1.ListIndex
        ilVehicleGroup = tgVehicleSets1(ilTemp).iCode

        If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
            mCntJob38Plus = -1
            Exit Function
        End If

        ilNoneExists = False                    'NONE not allowed in this list
        ilMinorGroupHdr = False                 'there is no minor vehicle group hdr to send to crystal for ths report
        If mCBCSet2Test(ilNoneExists, ilMinorGroupHdr) Then
            mCntJob38Plus = -1
            Exit Function
        End If

        If RptSelCt!ckcSelC13(1).Value = vbChecked Then
                If Not gSetFormula("SeparatePoliticals", "'Y'") Then
                    mCntJob38Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SeparatePoliticals", "'N'") Then
                    mCntJob38Plus = -1
                    Exit Function
                End If
            End If
        'End If

            If RptSelCt!ckcSelC13(2).Value = vbChecked Then         ' use sales source as major sort
                If Not gSetFormula("UseSSInSort", "'Y'") Then
                    mCntJob38Plus = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("UseSSInSort", "'N'") Then
                    mCntJob38Plus = -1
                    Exit Function
                End If
            End If
        ElseIf ilListIndex = CNT_CONTRACTVERIFY Then            '4-8-13
            slDate = RptSelCt!edcSelCFrom.Text
            If slDate <> "" Then
                If gValidDate(slDate) Then
                Else
                    mReset
                    RptSelCt!edcSelCFrom.SetFocus
                    Exit Function
                End If
            End If
            llStartDate = gDateValue(slDate)
            slDate = RptSelCt!edcSelCFrom1.Text
            If slDate <> "" Then
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCt!edcSelCFrom1.SetFocus
                    Exit Function
                End If
            End If
            llEndDate = gDateValue(slDate)
    
            slStr = Format$(llStartDate, "m/d/yy") + " - " + Format$(llEndDate, "m/d/yy")
            If Not gSetFormula("DatesRequested", "'" & slStr & "'") Then
                mCntJob38Plus = -1
                Exit Function
            End If

            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob38Plus = -1
                Exit Function
            End If
        ElseIf ilListIndex = CNT_INSERTION_ACTIVITY Then        '10-6-15
            'verify all dates entered
            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!edcSelCFrom)
            If Not mCntJob38Plus Then
                RptSelCt!edcSelCFrom.SetFocus
                Exit Function
            End If

            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!edcSelCFrom1)
            If Not mCntJob38Plus Then
                RptSelCt!edcsecfrom1.SetFocus
                Exit Function
            End If
            
            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!edcSelCTo)
            If Not mCntJob38Plus Then
                RptSelCt!edcSelCTo.SetFocus
                Exit Function
            End If

            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!edcSelCTo1)
            If Not mCntJob38Plus Then
                RptSelCt!edcSelCTo1.SetFocus
                Exit Function
            End If
            
            'get the earliest and latest dates from user requests
            slDate = RptSelCt!CSI_CalFrom.Text              'Date: 12/18/2019 added CSI calendar control for date entries --> edcSelCFrom.Text              'user sent start date
            llUserStartDateSent = gDateValue(slDate)            'user entered start date sent
            slDate = RptSelCt!CSI_CalTo.Text                'Date: 12/18/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text             'user sent end date
            llUserEndDateSent = gDateValue(slDate)                  'user entered end date sent
                        
            slDate = RptSelCt!CSI_From1.Text                'Date: 12/18/2019 added CSI calendar control for date entries --> edcSelCTo.Text            'user active start date
            llUserActiveStartDate = gDateValue(slDate)         'user entered active start date
            slDate = RptSelCt!CSI_To1.Text                  'Date: 12/18/2019 added CSI calendar control for date entries -->  edcSelCTo1.Text           'user active end date
            llUserActiveEndDate = gDateValue(slDate)                    'user entered active end date

            slStr = "Sent Dates " & Format$(llUserStartDateSent, "m/d/yy") & "-" & Format$(llUserEndDateSent, "m/d/yy") & "; Active Dates " & Format$(llUserActiveStartDate, "m/d/yy") & "-" & Format$(llUserActiveEndDate, "m/d/yy")
            If Not gSetFormula("DatesRequested", "'" & slStr & "'") Then
                mCntJob38Plus = -1
                Exit Function
            End If
            
            ilTemp = RptSelCt!cbcSet1.ListIndex
            If Not gSetFormula("SortBy", ilTemp) Then
                mCntJob38Plus = -1
                Exit Function
            End If
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob38Plus = -1
                Exit Function
            End If
        ElseIf ilListIndex = CNT_XML_ACTIVITY Then        '3-25-16
            'verify all dates entered
            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!CSI_CalFrom)        'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCFrom)
            If Not mCntJob38Plus Then
                RptSelCt!CSI_CalFrom.SetFocus   ' edcSelCFrom.SetFocus
                Exit Function
            End If

            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!CSI_CalTo)          'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCFrom1)
            If Not mCntJob38Plus Then
                RptSelCt!CSI_CalTo.SetFocus     ' edcsecfrom1.SetFocus
                Exit Function
            End If
            
            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!CSI_From1)          'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCTo)
            If Not mCntJob38Plus Then
                RptSelCt!CSI_From1.SetFocus     ' edcSelCTo.SetFocus
                Exit Function
            End If

            mCntJob38Plus = gVerifyDate(RptSelCt, RptSelCt!CSI_To1)            'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCTo1)
            If Not mCntJob38Plus Then
                RptSelCt!CSI_To1.SetFocus       ' edcSelCTo1.SetFocus
                Exit Function
            End If
            
            'get the earliest and latest dates from user requests
            slDate = RptSelCt!CSI_CalFrom.Text      'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCFrom.Text              'user sent start date
            llUserStartDateSent = gDateValue(slDate)            'user entered start date sent
            slDate = RptSelCt!CSI_CalTo.Text        'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text             'user sent end date
            llUserEndDateSent = gDateValue(slDate)                  'user entered end date sent
                        
            slDate = RptSelCt!CSI_From1.Text        'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCTo.Text            'user active start date
            llUserActiveStartDate = gDateValue(slDate)         'user entered active start date
            slDate = RptSelCt!CSI_To1.Text          'Date: 12/26/2019 added CSI calendar control for date entries --> edcSelCTo1.Text           'user active end date
            llUserActiveEndDate = gDateValue(slDate)                    'user entered active end date

            slStr = "Sent Dates " & Format$(llUserStartDateSent, "m/d/yy") & "-" & Format$(llUserEndDateSent, "m/d/yy") & "; Active Dates " & Format$(llUserActiveStartDate, "m/d/yy") & "-" & Format$(llUserActiveEndDate, "m/d/yy")
            If Not gSetFormula("DatesRequested", "'" & slStr & "'") Then
                mCntJob38Plus = -1
                Exit Function
            End If
            
            ilTemp = RptSelCt!cbcSet1.ListIndex
            If Not gSetFormula("SortBy", ilTemp) Then
                mCntJob38Plus = -1
                Exit Function
            End If
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob38Plus = -1
                Exit Function
            End If
        End If
        


    mCntJob38Plus = 1               'return, all OK
    Exit Function

End Function
'
'           mCBCSet2Test - determine the option selected in the list and
'           send Crystal the formula to sort by
'           <input>  ilNoneExists : true if option NONE exists
'                    ilMinorGroupHdr :  true if there is a minor vehicle group hdr to send to crystal
'           return - -1 if error
'
'       called from Sales Comparison or Billed & Booked Comparisons

Public Function mCBCSet2Test(ilNoneExists As Integer, ilMinorGroupHdr As Integer) As Integer
Dim ilVehicleGroup As Integer
Dim ilAdjust As Integer
Dim ilTemp As Integer
Dim slMinorSortStr As String * 8
Dim slPassToCrystal As String * 1
Dim ilListIndex As Integer

        slMinorSortStr = " AGBPSVH"  'none, adv, agy, bus cat, prod prot, slsp, veh, veh grp


        mCBCSet2Test = 0
        If ilNoneExists Then            'NONE exists in the list, offset all the selections by 1
            ilAdjust = 1
        Else
            ilAdjust = 2
        End If
        ilVehicleGroup = 0

        ilListIndex = RptSelCt!cbcSet2.ListIndex
        slPassToCrystal = Mid(slMinorSortStr, ilListIndex + ilAdjust, 1)

            If Not gSetFormula("MinorSortBy", "'" & slPassToCrystal & "'") Then
                mCBCSet2Test = -1
                Exit Function
            End If

            If slPassToCrystal = "H" Then     'H = Vehicle group selected for minor sort selection;
                                                                'or ilMinorGroupHdr set to true for B&B Comparisons
                'get the vehicle group selected for report heading (participant, format, market, etc)
                ilTemp = RptSelCt!lbcSelection(12).ListIndex    '3-18-16 chg from lbcselection(4)
                ilVehicleGroup = tgVehicleSets1(ilTemp).iCode
            End If

            If ilMinorGroupHdr Then
                If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
                    mCBCSet2Test = -1
                    Exit Function
                End If
            End If
        Exit Function


'        If RptSelCt!cbcSet2.ListIndex = 0 Then               'no minor sort selected
'            If Not gSetFormula("MinorSortBy", "''") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then         'advt
'            If Not gSetFormula("MinorSortBy", "'A'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then           'agency
'            If Not gSetFormula("MinorSortBy", "'G'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then           'bus cat
'            If Not gSetFormula("MinorSortBy", "'B'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then           'prod prot
'            If Not gSetFormula("MinorSortBy", "'P'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 5 Then          'slsp
'            If Not gSetFormula("MinorSortBy", "'S'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then             'vehicle
'            If Not gSetFormula("MinorSortBy", "'V'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        ElseIf RptSelCt!cbcSet2.ListIndex = 7 Then                                         'vehicle group
'            If Not gSetFormula("MinorSortBy", "'H'") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'            'get the vehicle group selected for report heading (participant, format, market, etc)
'            ilTemp = RptSelCt!lbcSelection(4).ListIndex
'            ilVehicleGroup = tgVehicleSets1(ilTemp).iCode
'        Else                'NONE selected
'            If Not gSetFormula("MinorSortBy", "''") Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        End If
'
'        If ilMinorGroupHdr Then
'            If Not gSetFormula("MinorVehicleGroupHdr", ilVehicleGroup) Then
'                mCBCSet2Test = -1
'                Exit Function
'            End If
'        End If
End Function
'
'        verify input parameters for Year, Month and # periods
'        <input> Form in case error
'                ilListIndex = report name
'                ilCorpSelection - radio button containing the corp selection (true if corp selected); otherwise
'                its standard or calendar selection
'
Public Function mVerifyMonthYrPeriods(Form As Form, ilListIndex As Integer, ilCorpSelection As Control) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slMonth                                                                               *
'******************************************************************************************

    Dim slStr As String
    Dim ilSaveMonth As Integer
    Dim ilRet As Integer
    Dim slMonthInYear As String * 36

    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
    mVerifyMonthYrPeriods = 0
    'verify user input dates
    slStr = Form!edcSelCFrom.Text
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        Form!edcSelCFrom.SetFocus                 'invalid year
        mVerifyMonthYrPeriods = -1
        Exit Function
    End If

    slStr = Form!edcSelCFrom1.Text                 'month in text form (jan..dec), or just a month # could have been entered
                                             'standard or calendar months
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
    Else
        slStr = str$(ilSaveMonth)
    End If

    ilRet = gVerifyInt(slStr, 1, 12)
    If ilRet = -1 Then
        mReset
        Form!edcSelCFrom1.SetFocus                 'invalid month #
        mVerifyMonthYrPeriods = -1
        Exit Function
    End If

    If ilCorpSelection Then                         'corporate months
        'convert the month name to the correct relative month # of the corp calendar
        'i.e. if 10 entered and corp calendar starts with oct, the result will be july (10th month of corp cal)
        slStr = Form!edcSelCFrom1.Text
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth <> 0 Then                           'input is text month name,
            slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
            igMonthOrQtr = gGetCorpMonthNoFromMonthName(slMonthInYear, slStr)         'getmonth # relative to start of corp cal
        Else
            igMonthOrQtr = Val(slStr)
        End If

    Else
        igMonthOrQtr = Val(slStr)                       'put month entered in global variable
    End If

    '3-2-02 verify input # period
    If (ilListIndex = CNT_BOB Or ilListIndex = CNT_BOBRECAP) And igRptCallType = CONTRACTSJOB Then
        slStr = Form!edcSelCTo1.Text                  'edit # periods
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then
            mReset
            Form!edcSelCTo1.SetFocus                 'invalid # periods
            mVerifyMonthYrPeriods = -1
            Exit Function
        End If
        igPeriods = Val(slStr)              '7-7-14
    Else                                            'BOBCOMPARE or BOB_SALESCOMPARE
        slStr = Form!edcSelCTo.Text                  'edit # periods
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then
            mReset
            Form!edcSelCTo.SetFocus                 'invalid # periods
            mVerifyMonthYrPeriods = -1
            Exit Function
        End If
        igPeriods = Val(slStr)              '7-7-14
    End If
    Exit Function
End Function

'       mAvgReptOptions - AVg Rate and Average Spot Price Report options
'                         to show in the report headers
Function mAvgReptOptions(slSelection As String, slInclude As String, slExclude As String) As Integer
    mAvgReptOptions = 0
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelCt!ckcSelC3(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelCt!ckcSelC3(1), slInclude, slExclude, "Orders"
    
    '6-15-11
    gIncludeExcludeCkc RptSelCt!ckcSelC10(0), slInclude, slExclude, "AirTime"
    gIncludeExcludeCkc RptSelCt!ckcSelC10(1), slInclude, slExclude, "Rep"
    
    gIncludeExcludeCkc RptSelCt!ckcSelC5(0), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(1), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(2), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(3), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelCt!ckcSelC5(4), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelCt!ckcSelC6(0), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSelCt!ckcSelC6(2), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelCt!ckcSelC12(0), slInclude, slExclude, "Polit"
    gIncludeExcludeCkc RptSelCt!ckcSelC12(1), slInclude, slExclude, "Non-Polit"
    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mAvgReptOptions = -1
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            mAvgReptOptions = -1
        End If
    End If
    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mAvgReptOptions = -1
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            mAvgReptOptions = -1
        End If
    End If
    
    '10-29-10  gross, net tnet option for Avg Rate, Advt units Ordered & Avg Spot price reports
    If RptSelCt!rbcSelC9(0).Value = True Then           'gross
        If Not gSetFormula("GrossNetTNet", "'G'") Then
            mAvgReptOptions = -1
            Exit Function
        End If
    ElseIf RptSelCt!rbcSelC9(1).Value = True Then           'net
        If Not gSetFormula("GrossNetTNet", "'N'") Then
            mAvgReptOptions = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("GrossNetTNet", "'T'") Then      't-net
            mAvgReptOptions = -1
            Exit Function
        End If
    End If

    Exit Function
End Function

'
'       Obtain date headers for Affiliate Spot Tracking, Commercial Changes,
'       and Sales Spot Tracking reports
'       Show Create DAte and Aired Date
'
Private Function mTrackAndComlChgDates(ByVal ilListIndex As Integer) As Integer
    Dim slReptDates As String
    Dim llDate As Long
    Dim slCreated As String
    Dim slAired As String
    Dim slStart As String
    Dim slEnd As String

    mTrackAndComlChgDates = 0
    slCreated = "Created: "
    slAired = "Aired: "
    'Date: 12/9/2019 added CSI calendar control for date entries
    If (ilListIndex = CNT_AFFILTRAK) Or (ilListIndex = CNT_COMLCHG) Or (ilListIndex = CNT_SPOTTRAK) Then
        slStart = RptSelCt!CSI_CalFrom.Text
        slEnd = RptSelCt!CSI_CalTo.Text
    Else
        slStart = RptSelCt!edcSelCFrom.Text
        slEnd = RptSelCt!edcSelCFrom1.Text
    End If
    If Trim$(slStart) = "" And Trim$(slEnd) = "" Then
        slCreated = slCreated & "All Dates"
    ElseIf Trim$(slStart) <> "" And Trim$(slEnd) <> "" Then
        llDate = gDateValue(slStart)
        slCreated = slCreated & Format$(slStart) & " - "
        llDate = gDateValue(slEnd)
        slCreated = slCreated & Format$(slEnd)
    Else
        If Trim$(slStart) = "" Then     'start date blank, from beginning thru indicated end date
            slCreated = slCreated & "Thru "
            llDate = gDateValue(slEnd)
            slCreated = slCreated & Format$(llDate, "m/d/yy")
        Else                        'end date is blank, get from indicated start date thru end of file
            slCreated = slCreated & "From "
            llDate = gDateValue(slStart)
            slCreated = slCreated & Format$(llDate, "m/d/yy")
        End If
    End If
    
    'Date: 12/9/2019 added CSI calendar control for date entries
    If (ilListIndex = CNT_AFFILTRAK) Or (ilListIndex = CNT_COMLCHG) Or (ilListIndex = CNT_SPOTTRAK) Then
        slStart = RptSelCt!CSI_From1.Text
        slEnd = RptSelCt!CSI_To1.Text
    Else
        slStart = RptSelCt!edcSelCTo.Text
        slEnd = RptSelCt!edcSelCTo1.Text
    End If
    If Trim$(slStart) = "" And Trim$(slEnd) = "" Then
        slAired = slAired & "All Dates"
    ElseIf Trim$(slStart) <> "" And Trim$(slEnd) <> "" Then
        llDate = gDateValue(slStart)
        slAired = slAired & Format$(slStart) & " - "
        llDate = gDateValue(slEnd)
        slAired = slAired & Format$(slEnd)
    Else
        If Trim$(slStart) = "" Then     'start date blank, from beginning thru indicated end date
            slAired = slAired & "Thru "
            llDate = gDateValue(slEnd)
            slAired = slAired & Format$(llDate, "m/d/yy")
        Else                        'end date is blank, get from indicated start date thru end of file
            slAired = slAired & "From "
            llDate = gDateValue(slStart)
            slAired = slAired & Format$(llDate, "m/d/yy")
        End If
    End If
    If Not gSetFormula("DateCreated", "'" & slCreated & "'") Then
        mTrackAndComlChgDates = -1
        Exit Function
    End If
    If Not gSetFormula("DateAired", "'" & slAired & "'") Then
        mTrackAndComlChgDates = -1
        Exit Function
    End If
    Exit Function
End Function

'
'           Send formula to .rpt to indicating gross or net for report heading
'
Function mGrossOrNetHdr() As Integer
    If RptSelCt!rbcSelC7(0).Value Then                'Gross
        If Not gSetFormula("GrossOrNet", "'G'") Then
            mGrossOrNetHdr = -1
            Exit Function
        End If
    ElseIf RptSelCt!rbcSelC7(1).Value Then                'Net
        If Not gSetFormula("GrossOrNet", "'N'") Then
            mGrossOrNetHdr = -1
            Exit Function
        End If
     Else                                                    'net-net/ t-net
        If Not gSetFormula("GrossOrNet", "'D'") Then
            mGrossOrNetHdr = -1
            Exit Function
        End If
    End If
    Exit Function
End Function

'
'   Get the base and compare dates to show on report for either standard bdcst month
'   or calendar month procesing
'
Private Function mSendBaseCompareToCrystal() As Integer
    Dim slStr As String
    Dim ilSaveMonth As Integer
    Dim slEarliest As String
    Dim slLatest As String
    Dim llDate As Long
    Dim illoop As Integer
    Dim slCode As String
    Dim slSaveYear As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String

    slStr = RptSelCt!edcSelCFrom1.Text             'month in text form (jan..dec)
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
    End If
    
    'Format the base date Month & year spans to send to Crystal
    slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(RptSelCt!edcSelCFrom.Text)
    If RptSelCt!rbcSelC9(1).Value Then      'std
        slEarliest = gObtainEndStd(slEarliest)
    Else                                    'cal month
        slEarliest = gObtainEndCal(slEarliest)
    End If
    gObtainYearMonthDayStr slEarliest, True, slSaveYear, slMonth, slDay
    slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
    slDate = slDate & " " & right$(Trim$(slSaveYear), 2)    'retrieve the last digits of year (i.e. 97, 98)
    illoop = Val(RptSelCt!edcSelCTo.Text)           '#months
    slStr = slEarliest
    Do While illoop <> 0
        If RptSelCt!rbcSelC9(1).Value Then      'std
            slLatest = gObtainEndStd(slStr)
            slStr = gObtainStartStd(slLatest)
        Else
            slLatest = gObtainEndCal(slStr)
            slStr = gObtainStartCal(slLatest)
        End If
        llDate = gDateValue(slLatest)
        llDate = llDate + 1
        slStr = Format$(llDate, "m/d/yy")
        illoop = illoop - 1
    Loop
    gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
    slCode = Left$(gMonthName(slLatest), 3)
    slCode = slCode & " " & right$(Trim$(slYear), 2)
    If Not gSetFormula("BaseDates", "'" & slDate & "-" & slCode & "'") Then
        mSendBaseCompareToCrystal = -1
        Exit Function
    End If

    'Format the comparison dates (last year)
    illoop = Val(RptSelCt!edcSelCTo.Text)           '#months
    If RptSelCt!rbcSelC11(1).Value = True Then      '3-23-16 include all last year (vs thru specified month)
        illoop = 12
        slEarliest = "1/15/" & Trim$(str$(Val(slSaveYear)))
    Else
        slEarliest = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(Val(slSaveYear)))
    End If
    'Format the comparison date Month & year spans  to send to crystal
    If RptSelCt!rbcSelC9(1).Value Then      'std
        slEarliest = gObtainEndStd(slEarliest)  'std end date for start of previous span
    Else
        slEarliest = gObtainEndCal(slEarliest)  'cal end date for start of previous span
    End If
    gObtainYearMonthDayStr slEarliest, True, slYear, slMonth, slDay
    slDate = Left$(gMonthName(slEarliest), 3)       'retrieve only first 3 char of month name
    slDate = slDate & " " & Trim$(right$(str$(Val(slSaveYear) - 1), 2))  'retrieve the last digits of year (i.e. 97, 98)

    'ilLoop = Val(RptSelCt!edcSelCTo.Text)           '#months
    slEarliest = slMonth & "/" & "15/" & Trim$(str((Val(slYear) - 1)))  'get previous year
    If RptSelCt!rbcSelC9(1).Value Then
        slEarliest = gObtainEndStd(slEarliest)
    Else
        slEarliest = gObtainEndCal(slEarliest)      'cal month
    End If
    slStr = slEarliest
    Do While illoop <> 0
        If RptSelCt!rbcSelC9(1).Value Then
            slLatest = gObtainEndStd(slStr)
            slStr = gObtainStartStd(slLatest)
        Else
            slLatest = gObtainEndCal(slStr)
            slStr = gObtainStartCal(slLatest)
        End If
        llDate = gDateValue(slLatest)
        llDate = llDate + 1
        slStr = Format$(llDate, "m/d/yy")
        illoop = illoop - 1
    Loop
    gObtainYearMonthDayStr slLatest, True, slYear, slMonth, slDay
    slCode = Left$(gMonthName(slLatest), 3)
    slCode = slCode & " " & right$(Trim$(slYear), 2)
    If Not gSetFormula("CompareDates", "'" & slDate & "-" & slCode & "'") Then
        mSendBaseCompareToCrystal = -1
        Exit Function
    End If
End Function
