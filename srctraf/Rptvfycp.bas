Attribute VB_Name = "RPTVFYCP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfycp.bas on Wed 6/17/09 @ 12:56 P
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
'Library calendar file- used to obtain post log date status
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmcGenCp                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenCp(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'   ilRet = gCmcGenCp()
'
'   ilRet (O)-  -1= Terminate, error in crystal gsetselectio or gsetformula
'               0 = Crystal input error
'               1 = successful crystal report
'               2 = successful bridge report
'
    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    
    gCmcGenCp = 0
'    slDate = RptSelCp!edcSelCFrom.Text
    slDate = RptSelCp!CSI_CalFrom.Text          '12-13-19 change to use csi calendar control
    If Not gValidDate(slDate) Then
        mReset
        RptSelCp!CSI_CalFrom.SetFocus
        Exit Function
    End If
    slStr = RptSelCp!edcSelCTo.Text                 'entered year
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelCp!edcSelCTo.SetFocus                 'invalid year
        'gCmcGenCp = -1
        Exit Function
    End If
    slStr = RptSelCp!edcSelCTo1.Text                  'edit qtr
    ilRet = gVerifyInt(slStr, 1, 4)
    If ilRet = -1 Then
        mReset
        RptSelCp!edcSelCTo1.SetFocus                 'invalid qtr
        'gCmcGenCp = -1
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
    If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
        'gCmcGenCp = -1
        Exit Function
    End If
    slStr = RptSelCp!edcSelCFrom1.Text
    If Val(slStr) > 4 Then
        mReset
        RptSelCp!edcSelCTo.SetFocus                 'invalid # qtr
        Exit Function
    End If
    
    slStr = "G"
    If RptSelCp!rbcGrossNet(1).Value Then
        slStr = "N"
    End If
    If Not gSetFormula("GrossNet", "'" & Trim$(slStr) & "'") Then
            gCmcGenCp = -1
            Exit Function
    End If
    
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenCp = -1
        Exit Function
    End If
    If ilRet = -1 Then
        'gCmcGenCp = -1
        gCmcGenCp = 0         'invalid input
        Exit Function
    End If
    gCmcGenCp = 1
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
    RptSelCp!frcOutput.Enabled = igOutput
    RptSelCp!frcCopies.Enabled = igCopies
    'RptSelCp!frcWhen.Enabled = igWhen
    RptSelCp!frcFile.Enabled = igFile
    RptSelCp!frcOption.Enabled = igOption
    'RptSelCp!frcRptType.Enabled = igReportType
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
        ilYear = RptSelCp!edcSelCTo.Text                'starting year
        ilMonth = RptSelCp!edcSelCTo1.Text              'month
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
         'Crystal tests this formula for 1, 2, 3, or 4 and uses it to calculate
         'other report headers, etc.
         If Not gSetFormula("WeekQtrHeader", "'" & slStr & "'") Then
             mWeekQtrHdr = False
             Exit Function
         End If
End Function
