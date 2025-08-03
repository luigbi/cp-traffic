Attribute VB_Name = "RPTVFYIA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyia.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Rptvfyia.Bas
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
''Global Const CNT_INVAFF = 32                'Affidavit of Performance
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
'*      Procedure Name:gGenReportIA                      *
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
Function gCmcGenIA(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
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
    Dim slTime As String
    Dim slTemp As String
    Dim ilBlanksBeforeLogo As Integer
    Dim ilBlanksAfterLogo As Integer
    gCmcGenIA = 0
    
    '4-27-12  test for vehicle name word wrap for all Affidavit of Performance rpts (for form reports only)
    If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
        If Not gSetFormula("WordWrapVehicle", "'Y'") Then
            gCmcGenIA = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("WordWrapVehicle", "'N'") Then
            gCmcGenIA = -1
            Exit Function
        End If
    End If


    '12-22-06 option to show Inv # on form
    If RptSelIA!ckcShowInvNo.Value = vbChecked Then
        If Not gSetFormula("ShowInvNo", "'Y'") Then
            gCmcGenIA = -1
             Exit Function
         End If
    Else
        If Not gSetFormula("ShowInvNo", "'N'") Then
            gCmcGenIA = -1
             Exit Function
         End If
    End If

    If (igJobRptNo = 1) Then
        'test month
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{IVR_Invoice_Rpt.ivrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({IVR_Invoice_Rpt.ivrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenIA = -1
            Exit Function
        End If
        
'       slDate = RptSelIA!edcSelCFrom.Text
        slDate = RptSelIA!CSI_CalFrom.Text      '8-27-19 use csi cal control vs edit box
        If Not gValidDate(slDate) Then
            mReset
            RptSelIA!CSI_CalFrom.SetFocus
            Exit Function
        End If
'        slDate = RptSelIA!edcSelCFrom1.Text
        slDate = RptSelIA!CSI_CalTo.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                mReset
                RptSelIA!CSI_CalTo.SetFocus
                Exit Function
            End If
        End If
        
'        slTemp = RptSelIA!edcSelCFrom.Text   'Get the user entered month
'        slMonth = mVerifyMonth(slTemp)
'        If slMonth = "0" Then
'            mReset
'            RptSelIA!edcSelCFrom.SetFocus
'            Exit Function
'        End If
'        'test year
'        slTemp = RptSelIA!edcSelCFrom1.Text   'Get the user entered year
'        slYear = Trim(Str$(gVerifyYear(slTemp)))
'        If slYear = "0" Then
'            mReset
'            RptSelIA!edcSelCFrom1.SetFocus
'            Exit Function
'        End If

        'The formulas sent to Crystal for the summary are sent (for common code) but not used
        '1-12-09 get the default terms incase terms not defined with agency or dir advt
        If Not gSetFormula("DefaultTerms", "'" & sgDefaultTerms & "'") Then
            gCmcGenIA = -1
            Exit Function
        End If

        '3-21-03 Send blanks to show in header to align to fit in windowed envelope
        If tgSpf.sExport = "0" Or tgSpf.sExport = "N" Or tgSpf.sExport = "Y" Or tgSpf.sExport = "" Then
            ilBlanksBeforeLogo = 0
        Else
            ilBlanksBeforeLogo = Val(tgSpf.sExport)
        End If
        If Trim$(tgSpf.sImport) = "0" Or Trim$(tgSpf.sImport) = "N" Or Trim$(tgSpf.sImport) = "Y" Or Trim$(tgSpf.sImport) = "" Then
            ilBlanksAfterLogo = 0
        Else
            ilBlanksAfterLogo = Val(tgSpf.sImport)
        End If
        If Not gSetFormula("BlanksBeforeLogo", ilBlanksBeforeLogo) Then
            gCmcGenIA = -1
            Exit Function
        End If
        If Not gSetFormula("BlanksAfterLogo", ilBlanksAfterLogo) Then
            gCmcGenIA = -1
            Exit Function
        End If

        If RptSelIA!rbcSortBy(1).Value Then     'sort by ISCI
            If RptSelIA!ckcShowRate.Value = vbChecked Then      'show rates ?
                If Not gSetFormula("HideRate", "'N'") Then
                    gCmcGenIA = -1
                     Exit Function
                End If
            Else
                If Not gSetFormula("HideRate", "'Y'") Then
                    gCmcGenIA = -1
                     Exit Function
                End If
            End If

            If RptSelIA!ckcSelC10(0).Value = vbChecked Then
                If Not gSetFormula("SkipPageNewISCI", "'Y'") Then
                    gCmcGenIA = -1
                     Exit Function
                End If
            Else
                If Not gSetFormula("SkipPageNewISCI", "'N'") Then
                    gCmcGenIA = -1
                     Exit Function
                End If
            End If
        End If
        
        If RptSelIA!ckcShowScript.Value = vbChecked Then            '4-25-17 show script
            If Not gSetFormula("ShowScript", "'Y'") Then
                gCmcGenIA = -1
                 Exit Function
            End If
        Else
            If Not gSetFormula("ShowScript", "'N'") Then
                gCmcGenIA = -1
                 Exit Function
            End If
        End If
        
         
    End If

    If (igJobRptNo = 2) Then
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        slSelection = "{IVR_Invoice_Rpt.ivrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({IVR_Invoice_Rpt.ivrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

        If Not gSetSelection(slSelection) Then
            gCmcGenIA = -1
            Exit Function
        End If

    End If

    ilRet = 1
    If ilRet = -1 Then
        gCmcGenIA = -1
        Exit Function
    ElseIf ilRet = 0 Then
        gCmcGenIA = 0
        Exit Function
    ElseIf ilRet = 2 Then
        gCmcGenIA = 2
        Exit Function
    End If

    gCmcGenIA = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportIA                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*          3-27-03 add option to skip to new page each vehicle
'*******************************************************
Function gGenReportIA() As Integer

    If RptSelIA!ckcUseCountAff.Value = vbChecked Then       'use affidavit that shows combination of air times + station counts by wek
        If Not gOpenPrtJob("InvAffBarter.Rpt") Then
            gGenReportIA = False
            Exit Function
        End If
    Else
        If RptSelIA!rbcSelC7(0).Value Then            'detail report
            If RptSelIA!rbcSortBy(0).Value Then       'sort by vehicle
                If RptSelIA!ckcSelC10(0) = vbChecked Then     'skip to new page each vehicle
                    If RptSelIA!ckcShowRate.Value = vbChecked Then  'show rate?
                        If Not gOpenPrtJob("InvAffVHRate.Rpt") Then
                            gGenReportIA = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("InvAffVH.Rpt") Then
                            gGenReportIA = False
                            Exit Function
                        End If
                    End If
                Else
                    If RptSelIA!ckcShowRate.Value = vbChecked Then      'show rate?
                        If Not gOpenPrtJob("InvAffRate.Rpt") Then
                            gGenReportIA = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("InvAff.Rpt") Then
                            gGenReportIA = False
                            Exit Function
                        End If
                    End If
                End If
            Else                            'sort by isci
                If Not gOpenPrtJob("InvISCI.Rpt") Then
                    gGenReportIA = False
                    Exit Function
                End If
            End If
        ElseIf RptSelIA!rbcSelC7(1).Value Then        'summary report
            If ((tgSpf.sInvAirOrder = "O") And (tgSpf.sBLaserForm = "2")) Then
                If Not gOpenPrtJob("InvAfSm2.Rpt") Then
                    gGenReportIA = False
                    Exit Function
                End If
            Else
                If Not gOpenPrtJob("InvAffSm.Rpt") Then
                    gGenReportIA = False
                    Exit Function
                End If
            End If
        ElseIf RptSelIA!rbcSelC7(2).Value Then        'both reports
            If igJobRptNo = 1 Then
    '            If RptSelIA!rbcSortBy(0).Value Then     'payee
    '                If Not gOpenPrtJob("InvAff.Rpt") Then
    '                    gGenReportIA = False
    '                    Exit Function
    '                End If
    '            Else
    '                If Not gOpenPrtJob("InvISCI.Rpt") Then
    '                    gGenReportIA = False
    '                    Exit Function
    '                End If
    '            End If
                If RptSelIA!rbcSortBy(0).Value Then       'sort by vehicle
                    If RptSelIA!ckcSelC10(0) = vbChecked Then     'skip to new page each vehicle
                        If RptSelIA!ckcShowRate.Value = vbChecked Then  'show rate?
                            If Not gOpenPrtJob("InvAffVHRate.Rpt") Then
                                gGenReportIA = False
                                Exit Function
                            End If
                        Else
                            If Not gOpenPrtJob("InvAffVH.Rpt") Then
                                gGenReportIA = False
                                Exit Function
                            End If
                        End If
                    Else
                        If RptSelIA!ckcShowRate.Value = vbChecked Then      'show rate?
                            If Not gOpenPrtJob("InvAffRate.Rpt") Then
                                gGenReportIA = False
                                Exit Function
                            End If
                        Else
                            If Not gOpenPrtJob("InvAff.Rpt") Then
                                gGenReportIA = False
                                Exit Function
                            End If
                        End If
                    End If
                Else                            'sort by isci
                    If Not gOpenPrtJob("InvISCI.Rpt") Then
                        gGenReportIA = False
                        Exit Function
                    End If
                End If
            End If
            If igJobRptNo = 2 Then
                If ((tgSpf.sInvAirOrder = "O") And (tgSpf.sBLaserForm = "2")) Then
                    If Not gOpenPrtJob("InvAfSm2.Rpt") Then
                        gGenReportIA = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("InvAffSm.Rpt") Then
                        gGenReportIA = False
                        Exit Function
                    End If
                End If
            End If
    
        End If
    End If
    gGenReportIA = True
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
    RptSelIA!frcoutput.Enabled = igOutput
    RptSelIA!frcCopies.Enabled = igCopies
    'RptSelIA!frcWhen.Enabled = igWhen
    RptSelIA!frcFile.Enabled = igFile
    RptSelIA!frcOption.Enabled = igOption
    'RptSelIA!frcRptType.Enabled = igReportType
    Beep
End Sub
''********************************************************
''*                                                      *
''*   Procedure Name:mVerifyMonth                        *
''*                                                      *
''*   Created:2/5/01       By:D. Smith                   *
''*   Modified:            By:                           *
''*                                                      *
''*   Comments: Test if entered month is between 1-12    *
''*             or Jan-Dec.  If so return the numeric    *
''*             string representation of the month       *
''*             otherwise return "0"                     *
''*                                                      *
''********************************************************
'Function mVerifyMonth(slInput As String) As String
'Dim ilIntVal As Integer
'    'Did they enter a number
'    If IsNumeric(slInput) Then
'        ilIntVal = Val(slInput)
'        If ((ilIntVal > 0) And (ilIntVal < 13)) Then
'            mVerifyMonth = slInput
'        Else
'            mVerifyMonth = "0"
'        End If
'    Else
'        Select Case UCase(slInput)
'            Case "Jan", "January"
'                mVerifyMonth = "1"
'            Case "FEB", "FEBRUARY"
'                mVerifyMonth = "2"
'            Case "MAR", "MARCH"
'                mVerifyMonth = "3"
'            Case "APR", "APRIL"
'                mVerifyMonth = "4"
'            Case "MAY"
'                mVerifyMonth = "5"
'            Case "JUN", "JUNE"
'                mVerifyMonth = "6"
'            Case "JUL", "JULY"
'                mVerifyMonth = "7"
'            Case "AUG", "AUGUST"
'                mVerifyMonth = "8"
'            Case "SEP", "SEPT", "SEPTEMBER"
'                mVerifyMonth = "9"
'            Case "OCT", "OCTOBER"
'                mVerifyMonth = "10"
'            Case "NOV", "NOVEMBER"
'                mVerifyMonth = "11"
'            Case "DEC", "DECEMBER"
'                mVerifyMonth = "12"
'            Case Else
'                mVerifyMonth = "0"
'        End Select
'    End If
'End Function
