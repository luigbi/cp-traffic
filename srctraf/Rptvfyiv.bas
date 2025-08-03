Attribute VB_Name = "RPTVFYIV"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyiv.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  mAvgUnitsOptions                                                                      *
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
'Dim hmLcf As Integer            'Library calendar file handle
'Dim tmLcf As LCF                'LCF record image
'Dim tmLcfSrchKey As LCFKEY0            'LCF record image
'Dim imLcfRecLen As Integer        'LCF record length
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportIV                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenIv(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'   ilRet = gCmcGenIv(ilListIndex)
'
'   ilRet (O)-  -1= Terminate, error in crystal gsetselectio or gsetformula
'               0 = Crystal input error
'               1 = successful crystal report
'               2 = successful bridge report
'
    Dim ilLoop As Integer
    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    gCmcGenIv = 0

    'Freezes igNowTime and igNowDate. See comments in RptCrIv/gCRQAvailsClearIV
    If igJobRptNo = 1 Then
        gCurrDateTime slStr, sgTime, slMonth, slDay, slYear
    End If
    'If Not gSetFormula("AsOfT", Trim$(Str$(CLng(gTimeToCurrency(sgTime, False))))) Then
    '    gCmcGenIv = -1
    '    Exit Function
    'End If

'    slStr = RptSelIv!edcSelCFrom.Text
    slStr = RptSelIv!CSI_CalFrom.Text       '9-6-17 use csi calendar control vs edit box
    slStr = Format$(gDateValue(slStr), "m/d/yy")
    If Not gSetFormula("EffDate", "'" & slStr & "'") Then
        gCmcGenIv = -1
        Exit Function
    End If

    slSelection = ""
    gUnpackCurrDateTime slDate, sgTime, slMonth, slDay, slYear
    If slSelection = "" Then
        slSelection = "{AVR_Quarterly_Avail.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({AVR_Quarterly_Avail.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(sgTime, False))))
    Else
        slSelection = "(" & slSelection & ")" & " And " & "(" & "{AVR_Quarterly_Avail.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({AVR_Quarterly_Avail.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(sgTime, False)))) & ")"
    End If
    If Not gSetSelection(slSelection) Then
        gCmcGenIv = -1
        Exit Function
    End If
    'Send selected rate card to Crystal
    slStr = ""
'    For ilLoop = 0 To RptSelIv!lbcSelection(12).ListCount - 1 Step 1
    For ilLoop = 0 To RptSelIv!lbcSelection(1).ListCount - 1 Step 1         '9-6-19 extra list boxes removed, changed index 12 to 1
'        If RptSelIv!lbcSelection(12).Selected(ilLoop) = True Then
        If RptSelIv!lbcSelection(1).Selected(ilLoop) = True Then
'            slSelection = "{RCF_Rate_Card.rcfName} = " & "'" & Trim$(RptSelIv!lbcSelection(12).List(ilLoop)) & "'"
            slSelection = "{RCF_Rate_Card.rcfName} = " & "'" & Trim$(RptSelIv!lbcSelection(1).List(ilLoop)) & "'"
'            slStr = Trim$(RptSelIv!lbcSelection(12).List(ilLoop))
            slStr = Trim$(RptSelIv!lbcSelection(1).List(ilLoop))
        End If
    Next ilLoop
    If Not gSetFormula("RateCard", "'" & slStr & "'") Then
        gCmcGenIv = -1
        Exit Function
    End If

    If RptSelIv!rbcSelC7(0).Value = True Then 'Pass Avail or Inv selection to Crystal
        slStr = "A"
    Else
        slStr = "I"
    End If
    If Not gSetFormula("AvailInv", "'" & slStr & "'") Then
        gCmcGenIv = -1
        Exit Function
    End If

    ilRet = mContractJob(0, slLogUserCode)
    gCmcGenIv = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportIV                      *
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
Function gGenReportIV() As Integer
    If igJobRptNo = 1 Then
        If Not gOpenPrtJob("invvalwk.Rpt") Then
            gGenReportIV = False
            Exit Function
        End If
        gGenReportIV = True
    ElseIf igJobRptNo = 2 Then
        If Not gOpenPrtJob("invvalmo.Rpt") Then
            gGenReportIV = False
            Exit Function
        End If
        gGenReportIV = True
    End If
End Function
'******************************************************************************
'
'                Procedure name:  mAvgUnitsOption
'                Created : 6/4/96   DH
'
'                <input & output> slSelection - formula to Crystal for
'                                               Header Type and Status inclusions
'                                               and exclusions
'                                 slExclude   - Excluded description for header
'                                 slInclude   - Included description for header
'                <output>  mAvgUnitsOption    - -1 for error, else 0
Function mAvgUnitsOptions(slSelection As String, slExclude As String, slInclude As String) As Integer 'VBC NR

    mAvgUnitsOptions = 0 'VBC NR
    If Not gSetCheck(RptSelIv!ckcSelC3(0).Value) Then 'VBC NR
        'exclude trades
        slExclude = "Trades" 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfPctTrade} <> 10000" 'VBC NR
    Else                                'show trades as inclusion 'VBC NR
        slInclude = "Trades" 'VBC NR
    End If 'VBC NR
    If Not gSetCheck(RptSelIv!ckcSelC3(1).Value) Then 'VBC NR
        'exclude no charge
        If Len(slExclude) = 0 Then 'VBC NR
            slExclude = "No Charge" 'VBC NR
        Else 'VBC NR
            slExclude = slExclude & ", No Charge" 'VBC NR
        End If 'VBC NR
        slSelection = slSelection & " And {CFF_Contract_Flight.cffActPrice} <> 0" 'VBC NR
    Else                                'show no charge as inclusion 'VBC NR
        If Len(slInclude) = 0 Then 'VBC NR
            slInclude = "No Charge" 'VBC NR
        Else 'VBC NR
            slInclude = slInclude & ", No Charge" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    If Not gSetCheck(RptSelIv!ckcSelC3(2).Value) Then 'VBC NR
        'exclude Direct response
        If Len(slExclude) = 0 Then 'VBC NR
            slExclude = "Direct Response" 'VBC NR
        Else 'VBC NR
            slExclude = slExclude & ", Direct Response" 'VBC NR
        End If 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfType} <> 'R'" 'VBC NR
    Else 'VBC NR
        If Len(slInclude) = 0 Then 'VBC NR
            slInclude = "Direct Response" 'VBC NR
        Else 'VBC NR
            slInclude = slInclude & ", Direct Response" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    If Not gSetCheck(RptSelIv!ckcSelC3(3).Value) Then 'VBC NR
        'exclude per inquiry
        If Len(slExclude) = 0 Then 'VBC NR
            slExclude = "Per Inquiry" 'VBC NR
        Else 'VBC NR
            slExclude = slExclude & ", Per Inquiry" 'VBC NR
        End If 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfType} <> 'Q'" 'VBC NR
    Else 'VBC NR
        If Len(slInclude) = 0 Then 'VBC NR
            slInclude = "Per Inquiry" 'VBC NR
        Else 'VBC NR
            slInclude = slInclude & ", Per Inquiry" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR

    If Not gSetCheck(RptSelIv!ckcSelC3(4).Value) Then 'VBC NR
        'exclude Holds
        If Len(slExclude) = 0 Then 'VBC NR
            slExclude = "Holds" 'VBC NR
        Else 'VBC NR
            slExclude = slExclude & ", Holds" 'VBC NR
        End If 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfStatus} <> 'H'" 'VBC NR
    Else 'VBC NR
        If Len(slInclude) = 0 Then 'VBC NR
            slInclude = "Holds" 'VBC NR
        Else 'VBC NR
            slInclude = slInclude & ", Holds" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR

    If Not gSetCheck(RptSelIv!ckcSelC3(5).Value) Then 'VBC NR
        'exclude Orders
        If Len(slExclude) = 0 Then 'VBC NR
            slExclude = "Orders" 'VBC NR
        Else 'VBC NR
            slExclude = slExclude & ", Orders" 'VBC NR
        End If 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfStatus} <> 'O'" 'VBC NR
    Else 'VBC NR
        If Len(slInclude) = 0 Then 'VBC NR
            slInclude = "Orders" 'VBC NR
        Else 'VBC NR
            slInclude = slInclude & ", Orders" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR

    If Len(slInclude) > 0 Then 'VBC NR
        If Not gSetFormula("Included", "'" & slInclude & "'") Then 'VBC NR
            mAvgUnitsOptions = -1 'VBC NR
        End If 'VBC NR
    Else 'VBC NR
        If Not gSetFormula("Included", "'" & " " & "'") Then 'VBC NR
            mAvgUnitsOptions = -1 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    If Len(slExclude) > 0 Then 'VBC NR
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then 'VBC NR
            mAvgUnitsOptions = -1 'VBC NR
        End If 'VBC NR
    Else 'VBC NR
        If Not gSetFormula("Excluded", "'" & " " & "'") Then 'VBC NR
            mAvgUnitsOptions = -1 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
End Function 'VBC NR
'*******************************************************
'*                                                     *
'*      Procedure Name:mContractJob                   *
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
Function mContractJob(ilListIndex As Integer, slLogUserCode As String) As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slInclude As String
    Dim slExclude As String
    mContractJob = 0
'    slDate = RptSelIv!edcSelCFrom.Text
    slDate = RptSelIv!CSI_CalFrom.Text
    If slDate <> "" Then
        If Not gValidDate(slDate) Then
            mReset
            RptSelIv!CSI_CalFrom.SetFocus
            Exit Function
        End If
    End If
    'Freezes igNowTime and igNowDate. See comments in RptCrIv/gCRQAvailsClearIV
    'If igJobRptNo = 1 Then
    '    gCurrDateTime slDate, sgTime, slMonth, slDay, slYear    'get current date and time for headers or keys to prepass files
    'End If
    'slSelection = "{AVR_Quarterly_Avails.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    'slSelection = slSelection & " And {AVR_Quarterly_Avails.avrGenTime} = " & Trim$(Str$(CLng(gTimeToCurrency(sgTime, False))))
    'If Not gSetSelection(slSelection) Then
    '    mContractJob = -1
    '    Exit Function
    'End If
    gUnpackDate igNowDate(0), igNowDate(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    slExclude = ""
    slInclude = ""

    gIncludeExcludeCkc RptSelIv!ckcSelC3(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelIv!ckcSelC3(1), slInclude, slExclude, "Orders"

    gIncludeExcludeCkc RptSelIv!ckcSelC5(0), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelIv!ckcSelC5(1), slInclude, slExclude, "Reserved"
    gIncludeExcludeCkc RptSelIv!ckcSelC5(2), slInclude, slExclude, "Remnant"

    gIncludeExcludeCkc RptSelIv!ckcSelC6(0), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelIv!ckcSelC6(1), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelIv!ckcSelC6(2), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelIv!ckcSelC6(3), slInclude, slExclude, "Promo"

    gIncludeExcludeCkc RptSelIv!ckcSelC8(0), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSelIv!ckcSelC8(1), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelIv!ckcSelC8(2), slInclude, slExclude, "Extra"


    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            mContractJob = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            mContractJob = -1
            Exit Function
        End If
    End If
    mContractJob = 1               'return, all OK
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
    RptSelIv!frcOutput.Enabled = igOutput
    RptSelIv!frcCopies.Enabled = igCopies
    'RptSelIv!frcWhen.Enabled = igWhen
    RptSelIv!frcFile.Enabled = igFile
    RptSelIv!frcOption.Enabled = igOption
    'RptSelIv!frcRptType.Enabled = igReportType
    Beep
End Sub
