Attribute VB_Name = "RPTVFYAP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyap.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' Created 10/31/97 by W.Bjerke  Actuals/Projection Comparison
'
' Release: 1.0
'
' Description:
'   This file contains the verification code for the Actual/Projection report
'****************************************************************************
Option Explicit
Option Compare Text
'Public sgTime As String
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
''Global Const COMM_SALESCOMM = 0         'sales commission
''Global Const COMM_PROJECTION = 1        'projection report
''Projections job report constants
''Global Const PRJ_SALESPERSON = 0
''Global Const PRJ_VEHICLE = 1
''Global Const PRJ_OFFICE = 2
''Global Const PRJ_CATEGORY = 3
''Global Const PRJ_SCENARIO = 4
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
''Global Const CNT_SALESCOMPARE = 25          'Sales Comparison
''Global Const CNT_CUMEACTIVITY = 26          'Cumulative Activity
''Global Const CNT_MAKEPLAN = 27              'Avg prices needed to make plan
''Global Const CNT_VEHCPPCPM = 28             'Current CPP & CPM by vehicle
''Global Const CNT_SALESANALYSIS = 29         'Sales Analysis Summary
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
''
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
'Global spot types for Spots by Advt & spots by Date & Time
'bit selectivity for charged and different types of no charge spots
'bits defined right to left (0 to 9)
'Global Const SPOT_CHARGE = &H1         'charged
'Global Const SPOT_00 = &H2          '0.00
'Global Const SPOT_ADU = &H4         'ADU
'Global Const SPOT_BONUS = &H8       'bonus
'Global Const SPOT_EXTRA = &H10      'Extra
'Global Const SPOT_FILL = &H20       'Fill
'Global Const SPOT_NC = &H40         'no charge
'Global Const SPOT_MG = &H80         'mg
'Global Const SPOT_RECAP = &H100     'recapturable
'Global Const SPOT_SPINOFF = &H200   'spinoff
'Library calendar file- used to obtain post log date status
'Dim hmLcf As Integer            'Library calendar file handle
'Dim tmLcf As LCF                'LCF record image
'Dim tmLcfSrchKey As LCFKEY0            'LCF record image
'Dim imLcfRecLen As Integer        'LCF record length
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
Function gCmcGenAp(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'   ilRet = gCmcGenAp(ilListIndex)
'
'   ilRet (O)-  -1= Terminate, error in crystal gsetselectio or gsetformula
'               0 = Crystal input error
'               1 = successful crystal report
'               2 = successful bridge report
'
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilYear As Integer
    Dim slStr As String
    Dim ilMonth As Integer
    Dim slHeader As String
    Dim slLYHeader As String
    gCmcGenAp = 0

       'dan M 7-02-08 NTR/Hard Cost option
    slStr = ""
    If RptSelAp!ckcSelC9(0).Value = 1 Or RptSelAp!ckcSelC9(1).Value = 1 Then
        slStr = slStr & "'With"
        If RptSelAp!ckcSelC9(0).Value = 1 Then
            slStr = slStr & " NTR"
            If RptSelAp!ckcSelC9(1).Value = 1 Then
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
        gCmcGenAp = False
        Exit Function
    End If
    gCurrDateTime slStr, sgTime, slMonth, slDay, slYear
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'    'Send report generation time to Crystal
'    If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(sgTime, False))))) Then
'        gCmcGenAp = -1
'        Exit Function
'    End If

    'Send the effective date to Crystal
'    slStr = RptSelAp!edcSelCFrom.Text
    slStr = RptSelAp!CSI_CalFrom.Text
    If Not gSetFormula("EffDate", "'" & slStr & "'") Then
        gCmcGenAp = -1
        Exit Function
    End If

    'Send year and quarter to Crystal subheader
    slStr = RptSelAp!edcSelCTo1.Text
    ilMonth = Val(slStr)
    If ilMonth = 1 Then
        slHeader = slStr + "st Quarter " + RptSelAp!edcSelCTo.Text
    ElseIf ilMonth = 2 Then
        slHeader = slStr + "nd Quarter " + RptSelAp!edcSelCTo.Text
    ElseIf ilMonth = 3 Then
        slHeader = slStr + "rd Quarter " + RptSelAp!edcSelCTo.Text
    ElseIf ilMonth = 4 Then
        slHeader = slStr + "th Quarter " + RptSelAp!edcSelCTo.Text
    End If
    If Not gSetFormula("WeekQtrHeader", "'" & slHeader & "'") Then
        gCmcGenAp = -1
        Exit Function
    End If

    'Setup last year's qtr column heading for Crystal report
    ilYear = Val(RptSelAp!edcSelCTo.Text)          'year
    'slStr = RptSelAP!edcSelCFrom.Text              'effective date
    'If Mid$(slStr, 3, 1) = "/" Then
    '    ilMonth = Val(Left$(slStr, 2))
    'Else
    '    ilMonth = Val(Left$(slStr, 1))
    'End If
    'slDate = Trim$(Str$(((ilMonth - 1) * 3 + 1))) & "/15/" & Trim$(Str$(ilYear))
    'slDate = gObtainStartStd(slDate)
    If ilMonth = 1 Then
        slLYHeader = "1st"
    ElseIf ilMonth = 2 Then
        slLYHeader = "2nd"
    ElseIf ilMonth = 3 Then
        slLYHeader = "3rd"
    Else
        slLYHeader = "4th"
    End If
    slLYHeader = Trim$(slLYHeader) & " Qtr" & str$(ilYear - 1)    'add Year
    If Not gSetFormula("LYWeekQtrHeader", "'" & slLYHeader & "'") Then
        gCmcGenAp = -1
        Exit Function
    End If

    'Send calander type to Crystal
    If RptSelAp!rbcSelC7(0).Value = True Then
        slStr = "Corp"
    Else
        slStr = "Std"
    End If
    If Not gSetFormula("CalType", "'" & slStr & "'") Then
        gCmcGenAp = -1
       Exit Function
    End If

    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(sgTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenAp = -1
        Exit Function
    End If

     gCmcGenAp = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
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
Function gGenReportAp() As Integer
    If Not gOpenPrtJob("actproj.Rpt") Then
        gGenReportAp = False
        Exit Function
    End If
    gGenReportAp = True
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
    RptSelAp!frcOutput.Enabled = igOutput
    RptSelAp!frcCopies.Enabled = igCopies
    'RptSelAp!frcWhen.Enabled = igWhen
    RptSelAp!frcFile.Enabled = igFile
    RptSelAp!frcOption.Enabled = igOption
    'RptSelAp!frcRptType.Enabled = igReportType
    Beep
End Sub
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
