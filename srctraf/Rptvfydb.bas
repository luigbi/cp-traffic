Attribute VB_Name = "RPTVFYDB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfydb.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelDB.Bas
'
' Release: 4.5
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
'Public tgRptSelDBAgencyCode() As SORTCODE
'Public sgRptSelDBAgencyCodeTag As String
'Public tgRptSelDBSalespersonCode() As SORTCODE
'Public sgRptSelDBSalespersonCodeTag As String
'Public tgRptSelDBAdvertiserCode() As SORTCODE
'Public sgRptSelDBAdvertiserCodeTag As String
'Public tgRptSelDBNameCode() As SORTCODE
'Public sgRptSelDBNameCodeTag As String
'Public tgRptSelDBBudgetCode() As SORTCODE
'Public sgRptSelDBBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelDBDemoCode() As SORTCODE
'Public sgRptSelDBDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
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
''Global Const PRJ_SCENARIO = 4
'Public Const PRJ_POTENTIAL = 4
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
'Public Const COLL_ACCTHIST = 10             'Account History
'Public Const COLL_MERCHANT = 11             'Merchandising & Promotions
'Public Const COLL_MERCHRECAP = 12            'Merchandising & Promotions Recap
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
'Public igOutputTo As Integer        '0 = display , 1 = print
'Public sgInclRates As String           'Includes Rates on contract = Y/N
'Public igSummaryID As Integer           'summary id # : (5-9)
'Public igNowDate(0 To 1)  As Integer     'generation date of pre-pass file
'Public igNowTime(0 To 1) As Integer       'generation time of pre-pass file
'Public lgNowTime As Long            '10-20-01
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
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenReportDB                              *
'*                                                             *
'*             Created:6/16/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                             *
'*         Comments: Formula setups for Crystal                *
'*                                                             *
'*          Return : 0 =  either error in input, stay in       *
'*                   -1 = error in Crystal, return to          *
'*                        calling program                      *
''*                       failure of gSetformula or another    *
'*                    1 = Crystal successfully completed       *
'*                    2 = successful Bridge                    *
'***************************************************************
Function gCmcGenDB(ilListIndex As Integer, ilGenShiftKey As Integer) As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slTime As String
    gCmcGenDB = 0
    slSelection = ""
    gUnpackDate igNowDate(0), igNowDate(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenDB = -1
        Exit Function
    End If
    If Not gSetFormula("SummaryID", igSummaryID) Then  '
        gCmcGenDB = -1
        Exit Function
    End If

    gCmcGenDB = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportDB                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportDB() As Integer
    If Not igUsingCrystal Then
        gGenReportDB = True
        Exit Function
    End If

    If igSummaryID = 5 Then                      'Line summary (vs qtr, week, vehicle, dp)
        If sgInclRates = "Y" Then             'with rates
            If Not gOpenPrtJob("DBLnRate.Rpt") Then
                gGenReportDB = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("DBLnNor.Rpt") Then 'with research but without CPP/CPM, and any other rates
                gGenReportDB = False
                Exit Function
            End If
        End If
    Else                                    'qtr,week vehicle, dp summary
        If sgInclRates = "Y" Then      'with rates
            If Not gOpenPrtJob("DBRate.Rpt") Then
                gGenReportDB = False
                Exit Function
            End If
        Else
            If Not gOpenPrtJob("DBNoRate.Rpt") Then     'Research without rates
                gGenReportDB = False
                Exit Function
            End If
        End If
    End If

    gGenReportDB = True
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
    RptSelDB!frcOutput.Enabled = igOutput
    RptSelDB!frcCopies.Enabled = igCopies
    'RptSelDB!frcWhen.Enabled = igWhen
    RptSelDB!frcFile.Enabled = igFile
    RptSelDB!frcOption.Enabled = igOption
    'RptSelDB!frcRptType.Enabled = igReportType
    Beep
End Sub
