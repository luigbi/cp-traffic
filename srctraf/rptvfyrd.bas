Attribute VB_Name = "RptVfyRD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyrd.bas on Wed 6/17/09 @ 12:56 P
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
'Public lgNowTime As Long
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
'*      Procedure Name:gGenReportRD                    *
'*                                                     *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*
'*          Return : 0 =  either error in input, stay in
'*                   -1 = error in Crystal, return to
'*                        calling program
''*                       failure of gSetformula or another
'*                    1 = Crystal successfully completed
'*                    2 = successful Bridge
'*******************************************************
Function gCmcGenRD(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
    Dim slSelection As String
    Dim slDate As String
    Dim slDateThru As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    Dim llEDate As Long
    Dim llSDate As Long
    Dim ilOffSet As Integer

    gCmcGenRD = 0

    slDateThru = RptSelRD!edcStartDate.Text
    If gValidDate(slDateThru) Then
    Else
        mReset
        RptSelRD!edcStartDate.SetFocus
        Exit Function
    End If
    llSDate = gDateValue(slDateThru)   'get complete date with year default
    slStr = Format$(llSDate, "m/d/yy")   'get the year in case not entered


    llEDate = gDateValue(slDateThru)
    slStr = slStr & " - " & Trim$(Format$(llEDate + 6, "m/d/yy")) 'get the year in case not entered

    'send start/end dates for crystal report heading
     If Not gSetFormula("DatesHdr", "'" & slStr & "'") Then
         gCmcGenRD = -1
         Exit Function
     End If


    ilOffSet = Val(RptSelRD!edcEST.Text)   'gather offsets for the different time zones
    If (ilOffSet < -6) Or (ilOffSet > 6) Then        'offsets can only be a 6 hour (+/-) displacement
        mReset
        RptSelRD!edcEST.SetFocus
        Exit Function
    End If

    If Not gSetFormula("ESTOffSet", ilOffSet) Then
         gCmcGenRD = -1
         Exit Function
     End If

    ilOffSet = Val(RptSelRD!edcCST.Text)
    If (ilOffSet < -6) Or (ilOffSet > 6) Then        'offsets can only be a 6 hour (+/-) displacement
        mReset
        RptSelRD!edcCST.SetFocus
        Exit Function
    End If
    If Not gSetFormula("CSTOffSet", ilOffSet) Then
         gCmcGenRD = -1
         Exit Function
     End If

    ilOffSet = Val(RptSelRD!edcMST.Text)
    If (ilOffSet < -6) Or (ilOffSet > 6) Then        'offsets can only be a 6 hour (+/-) displacement
        mReset
        RptSelRD!edcMST.SetFocus
        Exit Function
    End If
    If Not gSetFormula("MSTOffSet", ilOffSet) Then
         gCmcGenRD = -1
         Exit Function
     End If

    ilOffSet = Val(RptSelRD!edcPST.Text)
    If (ilOffSet < -6) Or (ilOffSet > 6) Then        'offsets can only be a 6 hour (+/-) displacement
        mReset
        RptSelRD!edcPST.SetFocus
        Exit Function
    End If
    If Not gSetFormula("PSTOffSet", ilOffSet) Then
         gCmcGenRD = -1
         Exit Function
     End If

     If RptSelRD!ckcSuppressHdr.Value = vbChecked Then
        If Not gSetFormula("SuppressHdr", "'" & "Y" & "'") Then
            gCmcGenRD = -1
            Exit Function
        End If
     Else
        If Not gSetFormula("SuppressHdr", "'" & "N" & "'") Then
            gCmcGenRD = -1
            Exit Function
        End If
     End If

     gCurrDateTime slDate, slTime, slMonth, slDay, slYear
     slSelection = "({CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
     slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
     If Not gSetSelection(slSelection) Then
         gCmcGenRD = -1
         Exit Function
     End If

    gCmcGenRD = 1
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
    RptSelRD!frcOutput.Enabled = igOutput
    RptSelRD!frcCopies.Enabled = igCopies
    RptSelRD!frcFile.Enabled = igFile
    RptSelRD!frcOption.Enabled = igOption
    Beep
End Sub
