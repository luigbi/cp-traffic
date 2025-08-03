Attribute VB_Name = "RPTVFYPC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfypc.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelPC.Bas
'
' Release: 4.5
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Public tgProcessProp() As SELECTPROP

Type SELECTPROP
    lCntrNo As Long
    lCode As Long
End Type
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
'Public igRCSelectedIndex As Integer         'selected r/c index
'Public igBSelectedIndex As Integer          'selected budget Plan index
'Public igBFCSelectedIndex As Integer        'selected budget forecast index
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
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenPC(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    Dim slSelection As String
    Dim slNameCode As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slReserve As String
    gCmcGenPC = 0
    slDate = RptSelPC!edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelPC!edcSelCFrom.SetFocus
        Exit Function
    End If

    slStr = RptSelPC!edcSelCFrom1.Text                  'edit qtr

    ilRet = gVerifyInt(slStr, 1, 4)
    If ilRet = -1 Then
        mReset
        RptSelPC!edcSelCFrom1.SetFocus                 'invalid qtr
        gCmcGenPC = -1
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable

    slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
    If Not gSetFormula("RateCard", "'" & slNameCode & "'") Then
        gCmcGenPC = -1
        Exit Function
    End If
        slExclude = ""
        slInclude = ""
        gIncludeExcludeCkc RptSelPC!ckcSelC1(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(1), slInclude, slExclude, "Orders"
        gIncludeExcludeCkc RptSelPC!ckcPropType(0), slInclude, slExclude, "Working"
        gIncludeExcludeCkc RptSelPC!ckcPropType(1), slInclude, slExclude, "Complete"
        gIncludeExcludeCkc RptSelPC!ckcPropType(2), slInclude, slExclude, "Unapproved"
        If tgSpf.sSystemType = "R" Then     'only Radio stations allowed to have feed spots
            gIncludeExcludeCkc RptSelPC!ckcSelC1(13), slInclude, slExclude, "Feed"
        End If

        gIncludeExcludeCkc RptSelPC!ckcSelC1(2), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(3), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(4), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(5), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(6), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(7), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(8), slInclude, slExclude, "Promo"

        gIncludeExcludeCkc RptSelPC!ckcSelC1(9), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(10), slInclude, slExclude, "Missed"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(11), slInclude, slExclude, "N/C"
        gIncludeExcludeCkc RptSelPC!ckcSelC1(12), slInclude, slExclude, "Fill"

        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenPC = -1
                Exit Function
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenPC = -1
                Exit Function
            End If
        End If

        'One or two column avail report
        If RptSelPC!rbcSelC2(1).Value Then         'two column (30/60)
            If Not gSetFormula("Columns1or2", "'2'") Then
                gCmcGenPC = -1
                Exit Function
            End If
        Else        'one column (units)
            If Not gSetFormula("Columns1or2", "'1'") Then
                gCmcGenPC = -1
                Exit Function
            End If
        End If

        If RptSelPC!rbcSelC5(0).Value Then           ' detail
            'Decide what to do with the reserves
            slReserve = "H"
            If RptSelPC!rbcSelC6(1).Value Then           'show the reserves
                slReserve = "S"
            End If

            If Not gSetFormula("Reserve", "'" & slReserve & "'") Then
                gCmcGenPC = -1
                Exit Function
            End If
        End If
        slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
        If Not gSetFormula("RateCard", "'" & slNameCode & "'") Then
            gCmcGenPC = -1
            Exit Function
        End If


    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{AVR_Quarterly_Avails.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({AVR_Quarterly_Avails.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenPC = -1
        Exit Function
    End If
    gCmcGenPC = 1         'ok
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
    RptSelPC!frcOutput.Enabled = igOutput
    RptSelPC!frcCopies.Enabled = igCopies
    'RptSelPC!frcWhen.Enabled = igWhen
    RptSelPC!frcFile.Enabled = igFile
    RptSelPC!frcOption.Enabled = igOption
    'RptSelPC!frcRptType.Enabled = igReportType
    Beep
End Sub
