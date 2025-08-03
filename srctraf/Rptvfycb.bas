Attribute VB_Name = "RPTVFYCB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfycb.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Rptvfyct.Bas
'
' Release: 1.0
'
' Description:
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
'Public Const CNT_MGREVENUE = 31             'MG Revenue 6-16-00
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
'*      Procedure Name:gGenReportCb                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'
'       6-16-00 Remove all references to Contract "BR"
'               and Insertion Orders (reports are coded
'               in rptselct)
'*                                                     *
'*******************************************************
Function gCmcGenCb(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'   ilRet = gCmcGenCb(ilListIndex)
'
'   ilRet (O)-  -1= Terminate, error in crystal gsetselectio or gsetformula
'               0 = Crystal input error
'               1 = successful crystal report
'               2 = successful bridge report
'
    Dim ilRet As Integer
    gCmcGenCb = 0
    Select Case igRptCallType

    Case CONTRACTSJOB
        If (ilListIndex < 11) Or (ilListIndex = CNT_INSERTION) Or (ilListIndex = CNT_MGREVENUE) Or (ilListIndex = CNT_MISSED) Then
            ilRet = mCntJob1_10(ilListIndex, slLogUserCode)
        Else
            ilRet = mCntJob11Plus(ilListIndex, slLogUserCode)
        End If
        If ilRet = -1 Then
            gCmcGenCb = -1
            Exit Function
        ElseIf ilRet = 0 Then
            gCmcGenCb = 0
            Exit Function
        ElseIf ilRet = 2 Then
            gCmcGenCb = 2
            Exit Function
        End If
    End Select
    gCmcGenCb = 1
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportCb                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'       6-16-00 Remove all references to Contract "BR"
'               and Insertion Orders (reports are coded
'               in rptselct)
'
'*******************************************************
Function gGenReportCb() As Integer
    Dim ilListIndex As Integer
    ilListIndex = RptSelCb!lbcRptType.ListIndex
    Select Case igRptCallType
        Case CONTRACTSJOB
            If Not igUsingCrystal Then
                gGenReportCb = True
                Exit Function
            End If
            If (igRptType = 0) And (ilListIndex > 1) Then
                ilListIndex = ilListIndex + 1
            End If
            'If rbcRptType(0).Value Then
            If ilListIndex = CNT_SPTCOMBO Then
                '-------------------------------
                'TTP 10674 - Spot and Digital Line combo Export or report?
                If RptSelCb.rbcOutput(3) Then
                    'Don't Open Crystal
                Else
                    If Not gOpenPrtJob("SptCombo.Rpt") Then        'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
                        gGenReportCb = False
                        Exit Function
                    End If
                End If
            
            ElseIf ilListIndex = CNT_SPTSBYADVT Then
                'check if "Include ISCI/Creative Title" option
                If RptSelCb!ckcIncludeISCI.Value = vbChecked Then
                    If Not gOpenPrtJob("SptByCntISCI.Rpt") Then        'spots by advt, agency or slsp
                        gGenReportCb = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("SptByCnt.Rpt") Then        'spots by advt, agency or slsp
                        gGenReportCb = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = CNT_MGREVENUE Then
                If Not gOpenPrtJob("MgRev.Rpt") Then        '6-16-00 MG Revenue
                    gGenReportCb = False
                    Exit Function
                End If

            ElseIf ilListIndex = CNT_SPTSBYDATETIME Then 'Spot by Time
                If Not gOpenPrtJob("SptByDte.Rpt") Then
                    gGenReportCb = False
                    Exit Function
                End If

                'Report!crcReport.ReportFileName = sgRptPath & "SpotsDte.Rpt"
            'ElseIf rbcRptType(4).Value Then

            ElseIf ilListIndex = 6 Or ilListIndex = 7 Then  'Placement
                'DS
                'If Not gOpenPrtJob("ChfDscrp.Rpt") Then
                If Not gOpenPrtJob("SptDiscr.Rpt") Then
                    gGenReportCb = False
                    Exit Function
                End If
            ElseIf ilListIndex = 8 Then  'MG's
                If Not gOpenPrtJob("MG.Rpt") Then
                    gGenReportCb = False
                    Exit Function
                End If

            ElseIf ilListIndex = CNT_SPOTSALES Then  'Spot Sales by Veh and Adv
                If RptSelCb!rbcSelCSelect(0).Value Or RptSelCb!rbcSelCSelect(1).Value Then    'subtotals by none, date
                    If Not gOpenPrtJob("SpSlsVeh.Rpt") Then
                        gGenReportCb = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("SpSlsAdv.Rpt") Then                     'subtotals by advt, sales source
                        gGenReportCb = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = 14 Then 'Missed spots
                If RptSelCb!rbcSelC7(0).Value = True Then      'vehicle missed option
                    If Not gOpenPrtJob("SptByDte.Rpt") Then
                        gGenReportCb = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("SptMissbySlsp.Rpt") Then
                        gGenReportCb = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = CNT_ACCRUEDEFER Then           '12-21-06
                If RptSelCb!rbcSelCSelect(1).Value Then         'sales origin has different sorting
                    If Not gOpenPrtJob("AccruDeferRecap.Rpt") Then
                        gGenReportCb = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("AccruDefer.Rpt") Then
                        gGenReportCb = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = CNT_HILORATE Then          '6-1-10
                If Not gOpenPrtJob("HiLoRate.Rpt") Then
                    gGenReportCb = False
                    Exit Function
                End If
             ElseIf ilListIndex = CNT_DISCREP_SUM Then          '6-22-16
                If Not gOpenPrtJob("SptDiscrSum.Rpt") Then
                    gGenReportCb = False
                    Exit Function
                End If
            End If
    End Select
    gGenReportCb = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mCntJob1_10                   *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Contract reports    *
'*      3/29/99 show user entered Start/end Active &
'               entered dates in paperwork summary heading
'*                                                     *
'*******************************************************
Function mCntJob1_10(ilListIndex As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValidDate                                                                           *
'******************************************************************************************

    Dim slDate As String
    Dim slTime As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slSelection As String
    Dim ilPreview As Integer
    Dim illoop As Integer
    Dim slStr As String
    Dim slOr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slInclude As String
    Dim slExclude As String
    Dim slInclStatus As String
    Dim slMissedStart As String
    Dim slMissedEnd As String
    Dim slMGStart As String
    Dim slMGEnd As String

    mCntJob1_10 = 0
    If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_MGREVENUE Or ilListIndex = CNT_SPTCOMBO Then     'MG Revenue or Spots by Advertiser
'        If (RptSelCb!edcSelCFrom.Text <> "") And (RptSelCb!edcSelCFrom1.Text <> "") Then
        If (RptSelCb!CSI_CalFrom.Text <> "") And (RptSelCb!CSI_CalTo.Text <> "") Then        'start and end dates must be present
            If StrComp(RptSelCb!edcSelCFrom1.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCb!CSI_CalFrom.Text
                If gValidDate(slDate) Then
                    If Not mVerifyDateInput(RptSelCb!CSI_CalTo) Then
                        Exit Function
                    End If
                    'slDate = RptSelCb!edcSelCFrom1.Text
                    'If Not gValidDate(slDate) Then
                    '    mReset
                    '    RptSelCb!edcSelCFrom1.SetFocus
                    '    Exit Function
                    'End If
                Else
                    mReset
                    RptSelCb!CSI_CalFrom.SetFocus
                    Exit Function
                End If
            Else
                If Not mVerifyDateInput(RptSelCb!CSI_CalFrom) Then
                    Exit Function
                End If
                'slDate = RptSelCb!edcSelCFrom.Text
                'If Not gValidDate(slDate) Then
                '    mReset
                '    RptSelCb!edcSelCFrom.SetFocus
                '    Exit Function
                'End If
            End If
        ElseIf RptSelCb!CSI_CalFrom.Text <> "" Then
            If Not mVerifyDateInput(RptSelCb!CSI_CalFrom) Then
                Exit Function
            End If
            'slDate = RptSelCb!edcSelCFrom.Text
            'If Not gValidDate(slDate) Then
            '    mReset
            '    RptSelCb!edcSelCFrom.SetFocus
            '    Exit Function
            'End If
        ElseIf RptSelCb!CSI_CalTo.Text <> "" Then
            If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                If Not mVerifyDateInput(RptSelCb!CSI_CalTo) Then
                    Exit Function
                End If
                'slDate = RptSelCb!edcSelCFrom1.Text
                'If Not gValidDate(slDate) Then
                '    mReset
                '    RptSelCb!edcSelCFrom1.SetFocus
                '    Exit Function
                'End If
            End If
        End If
'        If (RptSelCb!edcSelCTo.Text <> "") And (RptSelCb!edcSelCTo1.Text <> "") Then
        If (RptSelCb!CSI_CalFrom2.Text <> "") And (RptSelCb!CSI_CalTo2.Text <> "") Then
            If StrComp(RptSelCb!CSI_CalTo2.Text, "TFN", 1) <> 0 Then
                slDate = RptSelCb!CSI_CalFrom2.Text
                If gValidDate(slDate) Then
                    If Not mVerifyDateInput(RptSelCb!CSI_CalTo2) Then
                        Exit Function
                    End If
                    'slDate = RptSelCb!edcSelCTo1.Text
                    'If Not gValidDate(slDate) Then
                    '    mReset
                    '    RptSelCb!edcSelCTo1.SetFocus
                    '    Exit Function
                    'End If
                Else
                    mReset
                    RptSelCb!CSI_CalFrom2.SetFocus
                    Exit Function
                End If
            Else
                If Not mVerifyDateInput(RptSelCb!CSI_CalFrom2) Then
                    Exit Function
                End If
                'slDate = RptSelCb!edcSelCTo.Text
                'If Not gValidDate(slDate) Then
                '    mReset
                '    RptSelCb!edcSelCTo.SetFocus
                '    Exit Function
                'End If
            End If
        ElseIf RptSelCb!CSI_CalFrom2.Text <> "" Then
            If Not mVerifyDateInput(RptSelCb!CSI_CalFrom2) Then
                Exit Function
            End If
            'slDate = RptSelCb!edcSelCTo.Text
            'If Not gValidDate(slDate) Then
            '    mReset
            '    RptSelCb!edcSelCTo.SetFocus
            '    Exit Function
            'End If
        ElseIf RptSelCb!CSI_CalTo2.Text <> "" Then
            If StrComp(RptSelCb!CSI_CalTo2.Text, "TFN", 1) <> 0 Then
                If Not mVerifyDateInput(RptSelCb!CSI_CalTo2) Then
                    Exit Function
                End If
                'slDate = RptSelCb!edcSelCTo1.Text
                'If Not gValidDate(slDate) Then
                 '   mReset
                '    RptSelCb!edcSelCTo1.SetFocus
                '    Exit Function
                'End If
            End If
        End If

'        If ((Trim$(RptSelCb!edcSelCFrom.Text) = "" Or Trim$(RptSelCb!edcSelCFrom1.Text) = "" Or Trim$(RptSelCb!edcSelCTo.Text) = "" Or Trim$(RptSelCb!edcSelCTo1.Text) = "") And (ilListIndex = CNT_MGREVENUE)) Then
        If ((Trim$(RptSelCb!CSI_CalFrom.Text) = "" Or Trim$(RptSelCb!CSI_CalTo.Text) = "" Or Trim$(RptSelCb!CSI_CalFrom2.Text) = "" Or Trim$(RptSelCb!CSI_CalTo2.Text) = "") And (ilListIndex = CNT_MGREVENUE)) Then
            MsgBox "Enter Missed and/or MG dates", vbOK
            Exit Function
        End If
        
        If RptSelCb!rbcSelCSelect(0).Value Then         'Advt sort
            If Not gSetFormula("SortBy", "'A'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf RptSelCb!rbcSelCSelect(1).Value Then     'agency
            If Not gSetFormula("SortBy", "'G'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else                                            'slsp
            If Not gSetFormula("SortBy", "'S'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If

        If RptSelCb!rbcSelCInclude(0).Value Then        'show spot prices
            If Not gSetFormula("ShowPrice", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowPrice", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        'MG revenue doesnt ask to show status, defaults to always show it
        If RptSelCb!rbcSelC7(0).Value Then        'Show status column, yes
            If Not gSetFormula("ShowStatus", "'Y'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else                                       'Show status column, NO
            If Not gSetFormula("ShowStatus", "'N'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        
        If ilListIndex = CNT_MGREVENUE Then             '3-27-11
            If RptSelCb!rbcSelC9(0).Value = True Then       'show billed only
                slStr = "Billed Only"
            ElseIf RptSelCb!rbcSelC9(1).Value = True Then       'show unbilled only
                slStr = "Unbilled Only"
            Else
                'show both billed and unbilled
                slStr = "Billed, Unbilled"
            End If
            If Not gSetFormula("BillHdrInput", "'" & slStr & "'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If
        
        If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then           'only send formulas for Spots by Advt
            If RptSelCb!rbcSelC4(0).Value Then        'Sort station, then date
                If Not gSetFormula("DateOrStation", "'S'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else                                       'sort date, then station
                If Not gSetFormula("DateOrStation", "'D'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If

            End If
            '2-1-01 Detail or Summary versions
            If RptSelCb!rbcSelC9(0).Value Then        'Detail
                If Not gSetFormula("SummaryOnly", "'D'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else                                       'Summary
                If Not gSetFormula("SummaryOnly", "'S'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If

            End If
        End If
        'gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        'slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        'slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        'slSelection = gGRFSelectionForCrystal() '2/10/21
        slSelection = gGRFSelectionForCrystalRandom() 'TTP 10077 -Spots by Advertiser Report Speed-up for muti-users running report
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
'*********************************************************************************************
'*********************************************************************************************
    ElseIf (ilListIndex = CNT_SPTSBYDATETIME Or ilListIndex = CNT_MISSED) Then    'Spots by Date & Times, Missed spots
'        If (RptSelCb!edcSelCFrom.Text <> "") And (RptSelCb!edcSelCTo.Text <> "") Then
        If (RptSelCb!CSI_CalFrom.Text <> "") And (RptSelCb!CSI_CalTo.Text <> "") Then
                If StrComp(RptSelCb!edcSelCTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                        slDate = RptSelCb!CSI_CalTo.Text
                        If Not gValidDate(slDate) Then
                            mReset
                            RptSelCb!CSI_CalTo.SetFocus
                            Exit Function
                        End If
                    Else
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                Else
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                End If
            ElseIf RptSelCb!CSI_CalFrom.Text <> "" Then
                slDate = RptSelCb!CSI_CalFrom.Text
                If Not gValidDate(slDate) Then
                    mReset
                    RptSelCb!CSI_CalFrom.SetFocus
                    Exit Function
                End If
            ElseIf RptSelCb!CSI_CalTo.Text <> "" Then
                If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalTo.Text
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCb!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                End If
            End If
            slStr = RptSelCb!edcSelCTo1.Text            'contract # entered?
            If Not mVerifyNumber(slStr) Then
                mReset
                RptSelCb!edcSelCTo1.SetFocus
                Exit Function
            End If

            If RptSelCb!rbcSelCSelect(0).Value Then        'show spot prices
                If Not gSetFormula("ShowPrice", "'Y'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowPrice", "'N'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If

            '4-20-06 send formula to report to skip to new page for game vehicles
            slStr = "N"         'default to dont skip to new page each new game
            If ilListIndex = CNT_SPTSBYDATETIME Then
                illoop = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
                If (illoop And USINGSPORTS) = USINGSPORTS Then 'Using Sports
                    If RptSelCb!ckcSelC13(0).Value = vbChecked Then
                        slStr = "Y"
                    End If
                End If
                If Not gSetFormula("SkipPage", "'" & slStr & "'") Then
                    mCntJob1_10 = -1
                    Exit Function
                End If
            End If

            'gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            'slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            'slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
            slSelection = gGRFSelectionForCrystal()
            If Not gSetSelection(slSelection) Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'ElseIf (rbcRptType(0).Value) Or (rbcRptType(1).Value) Then
'*********************************************************************************************
'*********************************************************************************************

    ElseIf ilListIndex = 6 Then 'Placement
        'If (Not igUsingCrystal) Then
'            If (RptSelCb!edcSelCFrom.Text <> "") And (RptSelCb!edcSelCFrom1.Text <> "") Then
            If (RptSelCb!CSI_CalFrom.Text <> "") And (RptSelCb!CSI_CalTo.Text <> "") Then           '9-11-19 use csi calendar vs edit box
                If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                        slDate = RptSelCb!CSI_CalTo.Text
                        If gValidDate(slDate) Then
                        Else
                            mReset
                            RptSelCb!CSI_CalTo.SetFocus
                            Exit Function
                        End If
                    Else
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                Else
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                        'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        'slSelection = "{SDF_Spot_Detail.sdfDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    Else
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                End If
'            ElseIf RptSelCb!edcSelCFrom.Text <> "" Then
            ElseIf RptSelCb!CSI_CalFrom.Text <> "" Then
                slDate = RptSelCb!CSI_CalFrom.Text
                If gValidDate(slDate) Then
                    'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    'slSelection = "{SDF_Spot_Detail.sdfDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    mReset
                    RptSelCb!CSI_CalFrom.SetFocus
                    Exit Function
                End If
'            ElseIf RptSelCb!edcSelCFrom1.Text <> "" Then
            ElseIf RptSelCb!CSI_CalTo.Text <> "" Then
                If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalTo.Text
                    If gValidDate(slDate) Then
                        'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        'slSelection = "{SDF_Spot_Detail.sdfDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    Else
                        mReset
                        RptSelCb!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                End If
            End If
            If RptSelCb!rbcOutput(0).Value Then
                ilPreview = True
            ElseIf RptSelCb!rbcOutput(1).Value Then
                ilPreview = False
            End If
            lgStartingCntrNo = 0
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'gCntrDispRpt False, ilPreview, "CntrDisp.Lst", Val(slLogUserCode)
            gCntrDispRpt False

            mCntJob1_10 = 1
            Exit Function
    ElseIf ilListIndex = 7 Then 'Discrepancies
        'If (Not igUsingCrystal) Then
'            If (RptSelCb!edcSelCFrom.Text <> "") And (RptSelCb!edcSelCFrom1.Text <> "") Then
            If (RptSelCb!CSI_CalFrom.Text <> "") And (RptSelCb!CSI_CalTo.Text <> "") Then       '9-11-19 use csi calendar control vs edit box
                If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                        slDate = RptSelCb!CSI_CalTo.Text
                        If gValidDate(slDate) Then
                        Else
                            mReset
                            RptSelCb!CSI_CalTo.SetFocus
                            Exit Function
                        End If
                    Else
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                Else
                    slDate = RptSelCb!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                    Else
                        mReset
                        RptSelCb!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                End If
            ElseIf RptSelCb!CSI_CalFrom.Text <> "" Then
                slDate = RptSelCb!CSI_CalFrom.Text
                If gValidDate(slDate) Then
                Else
                    mReset
                    RptSelCb!CSI_CalFrom.SetFocus
                    Exit Function
                End If
            ElseIf RptSelCb!CSI_CalTo.Text <> "" Then
                If StrComp(RptSelCb!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                    slDate = RptSelCb!CSI_CalTo.Text
                    If gValidDate(slDate) Then
                    Else
                        mReset
                        RptSelCb!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                End If
            End If
            'If (RptSelCb!ckcAll.Value) And (lgOrigCntrNo > 0) Then
            If (RptSelCb!ckcAll.Value = vbChecked) Then
                If Trim$(RptSelCb!edcSelCTo.Text) <> "" Then
                    lgStartingCntrNo = Val(RptSelCb!edcSelCTo.Text)
                Else
                    lgStartingCntrNo = 0
                End If
            Else
                lgStartingCntrNo = 0
            End If
            If RptSelCb!rbcOutput(0).Value Then
                ilPreview = True
            ElseIf RptSelCb!rbcOutput(1).Value Then
                ilPreview = True
            End If
            'DS
            'gCntrDispRpt True, ilPreview, "CntrDisp.Lst", Val(slLogUserCode)
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob1_10 = -1
                Exit Function
            End If
            gCntrDispRpt True
            mCntJob1_10 = 1
            Exit Function
        'End If
    ElseIf ilListIndex = CNT_MG Then 'MG's
        slMissedStart = RptSelCb!edcSelCFrom.Text
        If slMissedStart <> "" Then
            If gValidDate(slMissedStart) Then
                gObtainYearMonthDayStr slMissedStart, True, slYear, slMonth, slDay
                slSelection = "{SMF_Spot_MG_Specs.smfMissedDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            Else
                mReset
                RptSelCb!edcSelCFrom.SetFocus
                Exit Function
            End If
        End If
        slMissedEnd = RptSelCb!edcSelCFrom1.Text
        If slMissedEnd <> "" Then
            If gValidDate(slMissedEnd) Then
                gObtainYearMonthDayStr slMissedEnd, True, slYear, slMonth, slDay
                If slSelection = "" Then
                    slSelection = "{SMF_Spot_MG_Specs.smfMissedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfMissedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If
            Else
                mReset
                RptSelCb!edcSelCFrom1.SetFocus
                Exit Function
            End If
        End If
        'Test entered makegood dates
        slMGStart = RptSelCb!edcSelCTo.Text
        If slMGStart <> "" Then
            If gValidDate(slMGStart) Then
                gObtainYearMonthDayStr slMGStart, True, slYear, slMonth, slDay
                If slSelection = "" Then
                    slSelection = "{SMF_Spot_MG_Specs.smfActualDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfActualDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If
            Else
                mReset
                RptSelCb!edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        slMGEnd = RptSelCb!edcSelCTo1.Text
        If slMGEnd <> "" Then
            If gValidDate(slMGEnd) Then
                gObtainYearMonthDayStr slMGEnd, True, slYear, slMonth, slDay
                If slSelection = "" Then
                    slSelection = "{SMF_Spot_MG_Specs.smfActualDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    slSelection = slSelection & " And " & "{SMF_Spot_MG_Specs.smfActualDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If
            Else
                mReset
                RptSelCb!edcSelCTo1.SetFocus
                Exit Function
            End If
        End If
        If slSelection = "" Then
            slSelection = "{SDF_Spot_Detail.sdfSpotType} <> 'X' and  {SDF_Spot_Detail.sdfPriceType} <> 'N'"
        Else
            slSelection = slSelection & " And " & "{SDF_Spot_Detail.sdfSpotType} <> 'X' and {SDF_Spot_Detail.sdfPriceType} <> 'N'"
        End If
        If slSelection <> "" Then
            slInclStatus = " and ("
        Else
            slInclStatus = "("
        End If
        slOr = ""
        slStr = ""
        If RptSelCb!ckcSelC3(0).Value = vbChecked Then           'include makegoods
            slInclStatus = slInclStatus & "{SMF_Spot_MG_Specs.smfSchStatus} = 'G'"
            slOr = " or "
            slStr = "Makegoods"
        End If
        If RptSelCb!ckcSelC3(1).Value = vbChecked Then           'include outsides
            slInclStatus = slInclStatus & slOr & " {SMF_Spot_MG_Specs.smfSchStatus} = 'O'"
            If slStr = "" Then
                slStr = "Outsides"
            Else
                slStr = slStr & " and Outsides"
            End If
        End If
        If Not gSetFormula("MG&Out", "'" & slStr & "'") Then
            mCntJob1_10 = -1
            Exit Function
        End If

        If RptSelCb!rbcSelCSelect(0).Value Then           'select vehicles for missed
            If Not gSetFormula("MissMGVeh", "'Vehicle selection for Missed Vehicles'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf RptSelCb!rbcSelCSelect(1).Value Then           'select vehicles for mg/out
            If Not gSetFormula("MissMGVeh", "'Vehicle selection for MG/Outside Vehicles'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        ElseIf RptSelCb!rbcSelCSelect(2).Value Then           'select vehicles for missed and mg/out
            If Not gSetFormula("MissMGVeh", "'Vehicle selection for either Missed or MG/Outside Vehicles'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        Else                                                'select vehicles for either missed or mg/out
            If Not gSetFormula("MissMGVeh", "'Vehicle selection for both Missed and MG/Outside Vehicles'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
        End If

        slInclStatus = slInclStatus & ")"
        slSelection = slSelection & slInclStatus
        If Not RptSelCb!ckcAll.Value = vbChecked Then          'not all vehicles selected
            If slSelection <> "" Then
                slSelection = "(" & slSelection & ") " & " and ("
                slOr = ""
            Else
                slSelection = "("
                slOr = ""
            End If
            'setup selective vehicles
            For illoop = 0 To RptSelCb!lbcSelection(6).ListCount - 1 Step 1
                If RptSelCb!lbcSelection(6).Selected(illoop) Then
                    slNameCode = tgCSVNameCode(illoop).sKey    'RptSelCb!lbcCSVNameCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    If RptSelCb!rbcSelCSelect(0).Value Then             'check missed vehicle
                        slSelection = slSelection & slOr & "{SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
                    ElseIf RptSelCb!rbcSelCSelect(1).Value Then           'match on mg vehicle                                              'check mg/out vehicle
                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode)
                    ElseIf RptSelCb!rbcSelCSelect(2).Value Then           'match on either missed or mg vehicle                                              'check mg/out vehicle
                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode) & " or " & " {SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
                    Else                                                'match on both missed or mg vehicle
                        slSelection = slSelection & slOr & "{SDF_Spot_Detail.sdfVefCode} = " & Trim$(slCode) & " and " & " {SMF_Spot_MG_Specs.smfOrigSchVef} = " & Trim$(slCode)
                    End If
                    slOr = " Or "
                End If
            Next illoop
            slSelection = slSelection & ")"
        End If
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

    End If

    'For Spots by Advt & Spots by Date & Time : send formulas to respective reports indicating
    'which spot types are included/excluded
    If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTSBYDATETIME Or ilListIndex = CNT_MISSED Or ilListIndex = CNT_SPTCOMBO Then
        slExclude = ""
        slInclude = ""
        If ilListIndex = CNT_SPTSBYADVT Or ilListIndex = CNT_SPTCOMBO Then
            gIncludeExcludeCkc RptSelCb!ckcSelC3(0), slInclude, slExclude, "Holds"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(1), slInclude, slExclude, "Orders"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(2), slInclude, slExclude, "Std"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(3), slInclude, slExclude, "Resv"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(4), slInclude, slExclude, "Rem"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(5), slInclude, slExclude, "DR"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(6), slInclude, slExclude, "PI"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(7), slInclude, slExclude, "PSA"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(8), slInclude, slExclude, "Promo"
            If RptSelCb!ckcSelC8(0).Value = vbChecked Then
                gIncludeExcludeCkc RptSelCb!ckcSelC8(0), slInclude, slExclude, "Acq Rates"
            End If
            If RptSelCb!ckcIncludeISCI.Value = vbChecked Then
                gIncludeExcludeCkc RptSelCb!ckcIncludeISCI, slInclude, slExclude, "ISCI"
            End If
        ElseIf ilListIndex = CNT_SPTSBYDATETIME Or ilListIndex = CNT_MISSED Then           '4-7-20 implement selectivity of cnt types in Missed report
            gIncludeExcludeCkc RptSelCb!ckcSelC10(0), slInclude, slExclude, "Holds"
            gIncludeExcludeCkc RptSelCb!ckcSelC10(1), slInclude, slExclude, "Orders"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(0), slInclude, slExclude, "Std"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(1), slInclude, slExclude, "Resv"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(2), slInclude, slExclude, "Rem"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(3), slInclude, slExclude, "DR"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(4), slInclude, slExclude, "PI"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(5), slInclude, slExclude, "PSA"
            gIncludeExcludeCkc RptSelCb!ckcSelC6(6), slInclude, slExclude, "Promo"
        End If
        
        gIncludeExcludeCkc RptSelCb!ckcSelC5(0), slInclude, slExclude, "Charge"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(1), slInclude, slExclude, "0.00"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(2), slInclude, slExclude, "ADU"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(3), slInclude, slExclude, "Bonus"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(4), slInclude, slExclude, "+Fill"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(5), slInclude, slExclude, "-Fill"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(6), slInclude, slExclude, "N/C"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(8), slInclude, slExclude, "Recap"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(9), slInclude, slExclude, "Spinoff"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(7), slInclude, slExclude, "MG"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(10), slInclude, slExclude, "BB"
        
        'only show inclusion/exclusion of cntr/feed spots if not included
        If tgSpf.sSystemType = "R" Then         'radio system are only ones that can have feed spots
            If Not RptSelCb!ckcSelC12(0).Value = vbChecked Then
                gIncludeExcludeCkc RptSelCb!ckcSelC12(0), slInclude, slExclude, "Contract spots"
            End If
            If Not RptSelCb!ckcSelC12(1).Value = vbChecked Then
                gIncludeExcludeCkc RptSelCb!ckcSelC12(1), slInclude, slExclude, "Feed spots"
            End If
        End If
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                mCntJob1_10 = -1
            End If
        Else
            If Not gSetFormula("Included", "'" & " " & "'") Then
                mCntJob1_10 = -1
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                mCntJob1_10 = -1
            End If
        Else
            If Not gSetFormula("Excluded", "'" & " " & "'") Then
                mCntJob1_10 = -1
            End If
        End If
    End If

    If ilListIndex = 3 Then
        If RptSelCb!rbcSelCSelect(0).Value Then
            If Not gSetFormula("Type", "'All Spots'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'Report!crcReport.Formulas(0) = "Type= 'All Spots'"
        ElseIf RptSelCb!rbcSelCSelect(1).Value Then
            If Not gSetFormula("Type", "'Missed Only'") Then
                mCntJob1_10 = -1
                Exit Function
            End If
            'Report!crcReport.Formulas(0) = "Type= 'Missed Only'"
        End If
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    ElseIf ilListIndex = 5 Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    ElseIf ilListIndex = 6 Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    ElseIf ilListIndex = 8 Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If
    ElseIf ilListIndex = 9 Then
        If Not gSetSelection(slSelection) Then
            mCntJob1_10 = -1
            Exit Function
        End If

    End If
    mCntJob1_10 = 1
    Exit Function
End Function

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
    Dim slDate As String
    Dim slTime As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slSelection As String
    Dim ilPreview As Integer
    Dim slStr As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slRateInclude As String
    Dim slRateExclude As String
    Dim slTypeInclude As String
    Dim slTypeExclude As String
    Dim llDate1 As Long
    Dim llDate2 As Long
    Dim ilDays As Integer
    Dim ilRet As Integer
    Dim slBillCycleInclude As String
    Dim slBillCycleExclude As String
    
    mCntJob11Plus = 0
    If ilListIndex = CNT_SPOTSALES Then         'spot sales by vehicle or adv
        slDate = RptSelCb!CSI_CalFrom.Text      '9-11-19 use csi calendar controls vs edit box
        If slDate <> "" Then
            If gValidDate(slDate) Then
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'slSelection = "{STF_Spot_Tracking.stfCreateDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            Else
                mReset
                RptSelCb!CSI_CalFrom.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCb!CSI_CalTo.Text
        If slDate <> "" Then
            If gValidDate(slDate) Then
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'If slSelection = "" Then
                '    slSelection = "{STF_Spot_Tracking.stfLogDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'Else
                '    slSelection = slSelection & " And " & "{STF_Spot_Tracking.stfLogDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'End If
            Else
                mReset
                RptSelCb!CSI_CalTo.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCb!edcSelCTo.Text             'start time
        If slDate <> "" Then
            If gValidTime(slDate) Then
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'slSelection = "{STF_Spot_Tracking.stfCreateDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            Else
                mReset
                RptSelCb!edcSelCTo.SetFocus
                Exit Function
            End If
        End If
        slDate = RptSelCb!edcSelCTo1.Text             'End time
        If slDate <> "" Then
            If gValidTime(slDate) Then
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'slSelection = "{STF_Spot_Tracking.stfCreateDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            Else
                mReset
                RptSelCb!edcSelCTo1.SetFocus
                Exit Function
            End If
        End If
        If RptSelCb!rbcSelCSelect(0).Value Or RptSelCb!rbcSelCSelect(1).Value Then
            slSelection = gGRFSelectionForCrystal()
        Else
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        End If
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
        If RptSelCb!rbcOutput(0).Value Then
            ilPreview = True
        ElseIf RptSelCb!rbcOutput(1).Value Then
            ilPreview = False
        End If
        If RptSelCb!rbcSelCSelect(0).Value Or RptSelCb!rbcSelCSelect(1).Value Then    'subtotals by none, date
            gSpotSalesVehRpt
        Else
            gSpotSalesAdvtRpt               'subtotals by advt, sales source
        End If
        mCntJob11Plus = 1
        Exit Function
     ElseIf ilListIndex = CNT_ACCRUEDEFER Then
'        slDate = RptSelCb!edcSelCFrom.Text
        slDate = RptSelCb!CSI_CalFrom.Text
        If slDate <> "" Then
            If Not mVerifyDateInput(RptSelCb!CSI_CalFrom) Then
                Exit Function
            End If
        End If
'        slDate = RptSelCb!edcSelCFrom1.Text
        slDate = RptSelCb!CSI_CalTo.Text
        If slDate <> "" Then
            If Not mVerifyDateInput(RptSelCb!CSI_CalTo) Then
                Exit Function
            End If
        End If

        If llDate1 > llDate2 Then
            mReset
            RptSelCb!CSI_CalTo.SetFocus
            Exit Function
        End If
        'Send crystal sort option
        If RptSelCb!rbcSelCSelect(0).Value Then     'sort by sales source
            If Not gSetFormula("SortBy", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        ElseIf RptSelCb!rbcSelCSelect(1).Value Then         'sort by sales origin
            If Not gSetFormula("SortBy", "'O'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SortBy", "'V'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

         If RptSelCb!ckcSelC12(0).Value = vbChecked Then     'include spot counts?
            If Not gSetFormula("ShowSpotCounts", "'Y'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else                                                'do not show spot counts
            If Not gSetFormula("ShowSpotCounts", "'N'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

'        slDate = RptSelCb!edcSelCFrom.Text
        slDate = RptSelCb!CSI_CalFrom.Text      '9-11-19 use csi calendar controls vs edit box
        slDate = Format$(gDateValue(slDate), "m/d/yy")
'        slStr = RptSelCb!edcSelCFrom1.Text
        slStr = RptSelCb!CSI_CalTo.Text
        slStr = Format$(gDateValue(slStr), "m/d/yy")
        If Not gSetFormula("DatesRequested", "'" & slDate & " - " & slStr & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If

        gIncludeExcludeCkc RptSelCb!ckcSelC6(0), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(1), slInclude, slExclude, "Resv"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(2), slInclude, slExclude, "Rem"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(3), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(4), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(5), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelCb!ckcSelC6(6), slInclude, slExclude, "Promo"

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

        gIncludeExcludeCkc RptSelCb!ckcSelC3(0), slTypeInclude, slTypeExclude, "Air Time"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(1), slTypeInclude, slTypeExclude, "REP"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(2), slTypeInclude, slTypeExclude, "NTR"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(3), slTypeInclude, slTypeExclude, "Hardcost"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(4), slTypeInclude, slTypeExclude, "Polit"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(5), slTypeInclude, slTypeExclude, "Non-Polit"
        If Len(slTypeInclude) > 0 Then
            If Not gSetFormula("TypeIncluded", "'" & slTypeInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Len(slTypeExclude) > 0 Then
            If Not gSetFormula("TypeExcluded", "'" & slTypeExclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        gIncludeExcludeCkc RptSelCb!ckcSelC5(0), slRateInclude, slRateExclude, "Charge"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(1), slRateInclude, slRateExclude, "0.00"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(2), slRateInclude, slRateExclude, "ADU"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(3), slRateInclude, slRateExclude, "Bonus"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(4), slRateInclude, slRateExclude, "NC"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(5), slRateInclude, slRateExclude, "Recap"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(6), slRateInclude, slRateExclude, "Spinoff"
        gIncludeExcludeCkc RptSelCb!ckcSelC5(7), slRateInclude, slRateExclude, "MG"
        If Len(slRateInclude) > 0 Then
            If Not gSetFormula("RateIncluded", "'" & slRateInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        If Len(slRateExclude) > 0 Then
            If Not gSetFormula("RateExcluded", "'" & slRateExclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        If RptSelCb!ckcSelC10(0).Value = vbChecked Then     'summary only
            If Not gSetFormula("SummaryOnly", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SummaryOnly", "'D'") Then   'detail
                mCntJob11Plus = -1
                Exit Function
            End If
        End If

        slSelection = gGRFSelectionForCrystal()
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
        
        gIncludeExcludeCkc RptSelCb!ckcSelC13(0), slBillCycleInclude, slBillCycleExclude, "Cal"
        gIncludeExcludeCkc RptSelCb!ckcSelC13(1), slBillCycleInclude, slBillCycleExclude, "Std"
        gIncludeExcludeCkc RptSelCb!ckcSelC13(2), slBillCycleInclude, slBillCycleExclude, "Weekly"
        If Len(slBillCycleInclude) > 0 Then
            slBillCycleInclude = "Billing Cycles: " & Trim$(slBillCycleInclude)
            If Not gSetFormula("BillCycle", "'" & slBillCycleInclude & "'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        
    ElseIf ilListIndex = CNT_HILORATE Then      '6-1-10
'        If Not mVerifyDateInput(RptSelCb!edcSelCFrom) Then      'verify date input
        If Not mVerifyDateInput(RptSelCb!CSI_CalFrom) Then      'verify date input  9-11-19 use csi calendar control vs edit box
            Exit Function
        End If
        slStr = RptSelCb!edcSelCTo.Text         'verify # days input, valid 1 -365
        ilRet = gVerifyInt(slStr, 1, 366)                   '1-366
        If ilRet = -1 Then
            mReset
            RptSelCb!edcSelCTo.SetFocus                 'invalid
            Exit Function
        End If
'        slDate = RptSelCb!edcSelCFrom.Text      'last date to include
        slDate = RptSelCb!CSI_CalFrom.Text      'last date to include
        slDate = Format$(gDateValue(slDate), "m/d/yy")
        llDate1 = gDateValue(slDate)
        ilDays = Val(RptSelCb!edcSelCTo.Text)         '# days back to start
        slStr = Format$((llDate1 - ilDays + 1), "m/d/yy")
        llDate2 = gDateValue(slStr)
        If ilDays = 1 Then
            slDay = "(" & str$(ilDays) & " day)"
        Else
            slDay = "(" & str$(ilDays) & " days)"
        End If
        If Not gSetFormula("DatesRequested", "'" & slStr & " - " & slDate & " " & slDay & "'") Then
            mCntJob11Plus = -1
            Exit Function
        End If
        gIncludeExcludeCkc RptSelCb!ckcSelC3(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(1), slInclude, slExclude, "Orders"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(2), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(3), slInclude, slExclude, "Resv"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(4), slInclude, slExclude, "Rem"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(5), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(6), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(7), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelCb!ckcSelC3(8), slInclude, slExclude, "Promo"

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

        If RptSelCb!rbcSelC4(1).Value = True Then     'summary only
            If Not gSetFormula("SummaryOnly", "'S'") Then
                mCntJob11Plus = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SummaryOnly", "'D'") Then   'detail
                mCntJob11Plus = -1
                Exit Function
            End If
        End If
        
        slSelection = gGRFSelectionForCrystal()
        If Not gSetSelection(slSelection) Then
            mCntJob11Plus = -1
            Exit Function
        End If
    ElseIf ilListIndex = CNT_DISCREP_SUM Then
        'verify user input dates
        slYear = Trim$(RptSelCb!edcSelCTo.Text)
        ilRet = gVerifyYear(slYear)
        If ilRet = 0 Then
            mReset
            RptSelCb!edcSelCTo.SetFocus
            Exit Function
        End If
        
        slStr = Trim$(RptSelCb!edcSelCFrom1.Text)
        lgStartingCntrNo = gVerifyLong(slStr, 0, 999999999)
        If lgStartingCntrNo = -1 Then                     'error
            mReset
            RptSelCb!edcSelCFrom1.SetFocus
            Exit Function
        End If
        
        slStr = RptSelCb!cbcSet1.Text
        slStr = slStr & " " & Trim$(RptSelCb!edcSelCTo.Text)
        If Not gSetFormula("MonthYearSelected", "'" & slStr & "'") Then
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
    ElseIf ilListIndex = CNT_SPTCOMBO Then
        '-------------------------------
        'TTP 10674 - Spot and Digital Line combo Export or report?
        If RptSelCb.rbcOutput(3) Then
            'Don't Open Crystal, we are Exporing to CSV
        Else
            gIncludeExcludeCkc RptSelCb!ckcSelC3(0), slInclude, slExclude, "Holds"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(1), slInclude, slExclude, "Orders"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(2), slInclude, slExclude, "Std"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(3), slInclude, slExclude, "Resv"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(4), slInclude, slExclude, "Rem"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(5), slInclude, slExclude, "DR"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(6), slInclude, slExclude, "PI"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(7), slInclude, slExclude, "PSA"
            gIncludeExcludeCkc RptSelCb!ckcSelC3(8), slInclude, slExclude, "Promo"
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
            'There are two sort options:
            ' 1. "by advertiser" sorts the data alphabetically by
            '   a. advertiser,
            '   b. contract number,
            '   c. vehicle,
            '   d. spot air date and spot air time / digital line start date.
            ' 2. "by vehicle" sorts the data by
            '   a. vehicle,
            '   b. spot air date and spot air time / digital line start date.
            If RptSelCb!rbcSelC4(0).Value Then         'Advt sort
                If Not gSetFormula("SortBy", "'A'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            Else                                       'Vehicle Sort
                If Not gSetFormula("SortBy", "'V'") Then
                    mCntJob11Plus = -1
                    Exit Function
                End If
            End If
            '------------------------------------------
            'Get Current Date/Time, Set GRF GenDate/Time selection
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mCntJob11Plus = -1
                Exit Function
            End If
            gPackDate slDate, igNowDate(0), igNowDate(1)
            gPackTime slTime, igNowTime(0), igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
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
    RptSelCb!frcOutput.Enabled = igOutput
    RptSelCb!frcCopies.Enabled = igCopies
    'RptSelCb!frcWhen.Enabled = igWhen
    RptSelCb!frcFile.Enabled = igFile
    RptSelCb!frcOption.Enabled = igOption
    'RptSelCb!frcRptType.Enabled = igReportType
    Beep
End Sub

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
    ilYear = RptSelCb!edcSelCTo.Text                'starting year
    If ilYear < 100 Then           'only 2 digit year input ie.  96, 95,
        If ilYear < 50 Then        'adjust for year 1900 or 2000
            ilYear = 2000 + ilYear
        Else
            ilYear = 1900 + ilYear
        End If
    End If
    
    ilMonth = RptSelCb!edcSelCTo1.Text              'month
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

Public Function mVerifyDateInput(DateInput As Control) As Integer
    Dim slDate As String
    mVerifyDateInput = True
    slDate = DateInput.Text
    If Not gValidDate(slDate) Then
        mReset
        DateInput.SetFocus
        mVerifyDateInput = False
        Exit Function
    End If
End Function
