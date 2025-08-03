Attribute VB_Name = "RPTVFYAC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyac.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSel.Bas
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
Function gCmcGenAC(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    Dim ilLoop As Integer
    Dim ilInx As Integer
    Dim ilYearInx As Integer
    gCmcGenAC = 0
    slStr = RptSelAc!edcSelCFrom.Text                 'entered year
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelAc!edcSelCFrom.SetFocus                 'invalid year
        gCmcGenAC = -1
        Exit Function
    End If
    slStr = RptSelAc!edcSelCFrom1.Text                  'edit # years to compare, 0 indicate show base year actuals only (max 3 years back)
    ilRet = gVerifyInt(slStr, 0, 3)
    If ilRet = -1 Then
        mReset
        RptSelAc!edcSelCFrom1.SetFocus                 'invalid # for years to compare
        gCmcGenAC = -1
        Exit Function
    End If
    For ilLoop = 1 To Val(slStr) + 1                 '0 indicates show base year only
        If RptSelAc!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
            ilYearInx = igYear - ilLoop + 1
                ilInx = gGetCorpCalIndex(ilYearInx)
            If ilInx < 0 Then
                'error, show message and return
                MsgBox "Corporate Year Missing for" & str$(ilYearInx), vbOKOnly + vbExclamation, "RptSelAc"
                gCmcGenAC = -1
                Exit Function
            End If
        End If
    Next ilLoop
    slSelection = ""
    For ilLoop = 1 To ilRet + 1
        If slSelection = "" Then
            slSelection = str$(igYear) & " vs "
        Else
            slSelection = slSelection & " " & Trim$(str$(igYear - ilLoop + 1))
        End If
    Next ilLoop
    If Not gSetFormula("CompareYears", "'" & slSelection & "'") Then
        gCmcGenAC = -1
        Exit Function
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenAC = -1
        Exit Function
    End If
    gCmcGenAC = 1         'ok
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
    RptSelAc!frcOutput.Enabled = igOutput
    RptSelAc!frcCopies.Enabled = igCopies
    'RptSelAc!frcWhen.Enabled = igWhen
    RptSelAc!frcFile.Enabled = igFile
    RptSelAc!frcOption.Enabled = igOption
    'RptSelAc!frcRptType.Enabled = igReportType
    Beep
End Sub
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
        ilYear = RptSelAc!edcSelCTo.Text                'starting year
        ilMonth = RptSelAc!edcSelCTo1.Text              'month
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
        slStr = slStr & " Quarter" & str$(ilYear)      'add Year
         If Not gSetFormula("WeekQtrHeader", "'" & slStr & "'") Then
             mWeekQtrHdr = False
             Exit Function
         End If
End Function
