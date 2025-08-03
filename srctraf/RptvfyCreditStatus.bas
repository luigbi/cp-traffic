Attribute VB_Name = "RPTVFYCREDITSTATUS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfy.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  sgRptSelCreditStatusAgencyCodeTag         sgRptSelCreditStatusSalespersonCodeTag                              *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  mAvgUnitsOptions                                                                      *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCreditStatus.Bas
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
Public tgRptSelCreditStatusAgencyCode() As SORTCODE
Public tgRptSelCreditStatusAgencyCodeCt() As SORTCODE
Public tgRptSelCreditStatusSalespersonCode() As SORTCODE
Public tgRptSelCreditStatusSalespersonCodeCt() As SORTCODE
'11/2/11: Moved to RptRec.Bas
'Public tgRptSelCreditStatusAdvertiserCode() As SORTCODE
Public sgRptSelCreditStatusAdvertiserCodeTag As String
Public tgRptSelCreditStatusAdvertiserCodeCb() As SORTCODE
Public tgRptSelCreditStatusAdvertiserCodeCt() As SORTCODE
Public tgRptSelCreditStatusNameCode() As SORTCODE
Public sgRptSelCreditStatusNameCodeTag As String
Public tgRptSelCreditStatusNameCodePP() As SORTCODE
Public tgRptSelCreditStatusBudgetCode() As SORTCODE
Public sgRptSelCreditStatusBudgetCodeTag As String
Public tgRptSelCreditStatusBudgetCodeAP() As SORTCODE
Public tgRptSelCreditStatusBudgetCodeCB() As SORTCODE
Public tgRptSelCreditStatusBudgetCodeCT() As SORTCODE
Public sgRptSelCreditStatusBudgetCodeTagCT As String
Public tgRptSelCreditStatusBudgetCodePS() As SORTCODE
Public sgRptSelCreditStatusBudgetCodeTagPS As String
Public tgRptSelCreditStatusBudgetCodeSP() As SORTCODE
Public sgRptSelCreditStatusBudgetCodeTagSP As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
Public tgRptSelCreditStatusDemoCodeCB() As SORTCODE
Public tgRptSelCreditStatusDemoCodeCP() As SORTCODE
Public sgRptSelCreditStatusDemoCodeTagCP As String
Public tgRptSelCreditStatusDemoCodeCT() As SORTCODE
Public sgRptSelCreditStatusDemoCodeTagCT As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Rate Card
Public smRCTag As String
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
'Public Const COLL_AGEOWNER = 13             'Ageing by Owner
'Public Const COLL_AGESS = 14                'Ageing by Sales Source
'Public Const COLL_AGEPRODUCER = 15            '2-10-00 Ageing by Producer
'Public Const COPY_REGIONS = 15              '7-18-00 copy by regions
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
'Global spot types for Spots by Advt & spots by Date & Time
'bit selectivity for charged and different types of no charge spots
'bits defined right to left (0 to 9)
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
Dim lmStartDates() As Long          'array of 13 bdcst or corp start dates
Dim lmEndDates() As Long            'array of 13 bdcst or corp end dates
Public Const HardCost = 0
Public Const Airtime = 1
Public Const NTR = 2
Private Const Correct = 0
Private Const Incorrect = -1
'

Function gCmcGen(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String, Optional slYear As String = "", Optional slMonth As String = "", Optional slDay As String = "", Optional slTime As String = "") As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBlanksBeforeLogo            ilBlanksAfterLogo                                       *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim slSelection As String
    Dim slLastName As String
    Dim slFirstName As String
    Dim slName As String
    Dim slNameCode As String
    Dim slCity As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slOr As String
    Dim slDate As String
    Dim slDateFrom As String
    Dim slDateTo As String
    Dim slYear2 As String
    Dim slYearCurr As String
    Dim slMonth2 As String
    Dim slMonthCurr As String
    Dim slEarliestMM As String
    Dim slEarliestYY As String
    Dim slLatestMM As String
    Dim slLatestYY As String
    Dim slDay2 As String
    Dim slStr As String
    Dim ilFormulaNo As Integer
    Dim slTime2 As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim ilPreview As Integer
    Dim slBaseDate As String
    Dim llDate1 As Long
    Dim llDate2 As Long
    Dim slSortStr As String
    Dim slHeader As String
    Dim slGenYear As String
    Dim slGenMonth As String
    Dim slGenDay As String
    Dim slInclude As String
    Dim slExclude As String
    Dim ilInclHardCost As Integer
    Dim ilSaveMonth As Integer
    gCmcGen = 0
    Select Case igRptCallType
        
        Case COLLECTIONSJOB
           
            If ilListIndex = 7 Then 'agency and advertiser Credit Status
                slSelection = ""
                If igJobRptNo = 1 Then  'Agency
                    'If Not RptSelCreditStatus!ckcSel1(0).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{AGF_Agencies.agfCreditRestr} <>" & "'N'"
                    '    Else
                    '        slSelection = "{AGF_Agencies.agfCreditRestr} <>" & "'N'"
                    '    End If
                    'End If
                    'If Not RptSelCreditStatus!ckcSel1(1).Value = vbChecked Then
                   '     If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    '    Else
                    '        slSelection = "{@Credit Used} <>" & "0"
                    '    End If
                    'End If
                    'If RptSelCreditStatus!ckcADate.Value = vbChecked Then          'see if only overdue accounts required
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "(({AGF_Agencies.agfCreditRestr} = 'W' or {AGF_Agencies.agfCreditRestr} = 'M' or {AGF_Agencies.agfCreditRestr} = 'T' ) and  ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > 0)) "
                    '        slSelection = slSelection & " or (( {AGF_Agencies.agfCreditRestr} = 'L' ) and  ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > ({AGF_Agencies.agfCreditLimit}/100))) or (({AGF_Agencies.agfCreditRestr} = 'P') and ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > 0))"
                    '    Else
                    '        slSelection = "(({AGF_Agencies.agfCreditRestr} = 'W' or {AGF_Agencies.agfCreditRestr} = 'M' or {AGF_Agencies.agfCreditRestr} = 'T' ) and  ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > 0)) "
                    '        slSelection = slSelection & "  or (( {AGF_Agencies.agfCreditRestr} = 'L' ) and  ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > ({AGF_Agencies.agfCreditLimit}/100))) or (({AGF_Agencies.agfCreditRestr} = 'P') and ({AGF_Agencies.agfCurrAR} + {AGF_Agencies.agfUnbilled} > 0))"
                    '    End If
                    'End If

                    'always select credit limit, plus all cash inadvance clients
                    slSelection = "({AGF_Agencies.agfCreditRestr} = 'L' or {AGF_Agencies.agfCreditRestr} = 'W' or {AGF_Agencies.agfCreditRestr} = 'M' or {AGF_Agencies.agfCreditRestr} = 'T' or {AGF_Agencies.agfCrdApp} = 'R') "

                    'include zero balance?
                    If RptSelCreditStatus!ckcSel1(1).Value = vbChecked Then         'include zero balance clients

                    Else
                        slSelection = "(" & slSelection & " and ({@Credit Used} <> 0) )"
                    End If

                    'include no new orders
                    If RptSelCreditStatus!ckcADate.Value = vbChecked Then
                        slSelection = slSelection & " or ( {AGF_Agencies.agfCreditRestr} = 'P') "
                    End If

                    If RptSelCreditStatus!ckcSel1(0).Value = vbChecked Then         'include unrestricted
                        slSelection = slSelection & " or ( {AGF_Agencies.agfCreditRestr} = 'N') "
                    End If
                    
                    'TTP 9893
                    slSelection = slSelection & " And {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                Else    'Advertiser
                    'If Not RptSelCreditStatus!ckcSel1(0).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{ADF_Advertisers.adfCreditRestr} <>" & "'N'"
                    '    Else
                    '        slSelection = "{ADF_Advertisers.adfCreditRestr} <>" & "'N'"
                    '    End If
                    'End If
                    'If Not RptSelCreditStatus!ckcSel1(1).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    '    Else
                    '        slSelection = "{@Credit Used} <>" & "0"
                    '    End If
                    'End If
                    'If RptSelCreditStatus!ckcADate.Value = vbChecked Then          'see if only overdue accounts required
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "(({ADF_Advertisers.adfCreditRestr} = 'W' or {ADF_Advertisers.adfCreditRestr} = 'M' or {ADF_Advertisers.adfCreditRestr} = 'T' ) and ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled} > 0)) "
                    '        slSelection = slSelection & " or (( {ADF_Advertisers.adfCreditRestr} = 'L' ) and  ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled}> ({ADF_Advertisers.adfCreditLimit}/100))) or (({ADF_Advertisers.adfCreditRestr} = 'P') and ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled} > 0)) "
                    '    Else
                    '        slSelection = "(({ADF_Advertisers.adfCreditRestr} = 'W' or {ADF_Advertisers.adfCreditRestr} = 'M' or {ADF_Advertisers.adfCreditRestr} = 'T' ) and  ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled} > 0)) "
                    '        slSelection = slSelection & " or (( {ADF_Advertisers.adfCreditRestr} = 'L' ) and  ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled}> ({ADF_Advertisers.adfCreditLimit}/100))) or (({ADF_Advertisers.adfCreditRestr} = 'P') and ({ADF_Advertisers.adfCurrAR} + {ADF_Advertisers.adfUnbilled} > 0)) "
                    '    End If
                    'End If

                    'always select credit limit, plus all cash inadvance clients
                    slSelection = "({ADF_Advertisers.adfCreditRestr} = 'L' or {ADF_Advertisers.adfCreditRestr} = 'W' or {ADF_Advertisers.adfCreditRestr} = 'M' or {ADF_Advertisers.adfCreditRestr} = 'T' or {ADF_Advertisers.adfCrdApp} = 'R') "

                    'include zero balance?
                    If RptSelCreditStatus!ckcSel1(1).Value = vbChecked Then         'include zero balance clients

                    Else
                        slSelection = "(" & slSelection & " and ({@Credit Used} <> 0) )"
                    End If

                    'include no new orders
                    If RptSelCreditStatus!ckcADate.Value = vbChecked Then
                        slSelection = slSelection & " or ( {ADF_Advertisers.adfCreditRestr} = 'P') "
                    End If

                    If RptSelCreditStatus!ckcSel1(0).Value = vbChecked Then         'include unrestricted
                        slSelection = slSelection & " or ( {ADF_Advertisers.adfCreditRestr} = 'N') "
                    End If

                    'TTP 9893
                    slSelection = slSelection & " And {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                End If

                '2-24-05 determine to show action comments
                If RptSelCreditStatus!ckcInclCommentsA.Value = vbChecked Then
                    'yes, include comments.  If no date entered, show all comments
                    'force date entered as the earliest date possible
                    If RptSelCreditStatus!edcSelA.Text = "" Then
                        slYear2 = "1970"
                        slMonth2 = "1"
                        slDay2 = "1"
                    Else
                        slDateFrom = RptSelCreditStatus!edcSelA.Text
                        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
                        gObtainYearMonthDayStr slDateFrom, True, slYear2, slMonth2, slDay2

                    End If
                Else                    'dont show any comments
                    slYear2 = "2020"
                    slMonth2 = "12"
                    slDay2 = "31"
                End If
                If Not gSetFormula("ShowCommentsAsOf", "Date(" & slYear2 & "," & slMonth2 & "," & slDay2 & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

                If RptSelCreditStatus!ckcDelinquentOnly.Value = vbChecked Then      'delinquents (overdue) only
                    If Not gSetFormula("DelinquentOnly", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("DelinquentOnly", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

            End If

            If Not gSetSelection(slSelection) Then
                gCmcGen = -1
                Exit Function
            End If
    End Select
    gCmcGen = 1
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
Function gGenReport(Optional slYear As String = "", Optional slMonth As String = "", Optional slDay As String = "", Optional slTime As String = "") As Integer
    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slDateFrom As String
    Dim slDateTo As String
    Dim ilListIndex As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llDate2 As Long
    Dim llTemp As Long
    Dim blCustomizeLogo As Boolean
    
    blCustomizeLogo = False             'default to generic client logo for all reports, except statements which use the logo associated with SAles Source
                                        'Flag is sent when opening .rpt; its an optional field defaulted to False if not sent

    ilListIndex = RptSelCreditStatus!lbcRptType.ListIndex
    Select Case igRptCallType
      
        Case COLLECTIONSJOB
            If ilListIndex = 7 Then 'Credit Status
                slDate = RptSelCreditStatus!edcSelA.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSelCreditStatus!edcSelA.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
           
                If igJobRptNo = 1 Then
                    If Not gOpenPrtJob("CreditAg.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("CreditAd.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            End If
        Case GENERICBUTTON
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            slStr = RptSelCreditStatus!edcSelA.Text
            If InStr(slStr, ".") = 0 Then
                slStr = slStr & ".Rpt"
            End If
            If Not gOpenPrtJob("Generic\" & slStr) Then
                gGenReport = False
                Exit Function
            End If
    End Select
    gGenReport = True
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
    RptSelCreditStatus!frcOutput.Enabled = igOutput
    RptSelCreditStatus!frcCopies.Enabled = igCopies
    'RptSelCreditStatus!frcWhen.Enabled = igWhen
    RptSelCreditStatus!frcFile.Enabled = igFile
    RptSelCreditStatus!frcOption.Enabled = igOption
    'RptSelCreditStatus!frcRptType.Enabled = igReportType
    Beep
End Sub

'
'
'           mVerifyDate - verify date entered as valid
'           <input> edcDate - control field (edit box) containing date string
'                   ilDateReqd - true if date required, else false (no date ok)
'           <output> llDate - date converted as Long
'           <return> 0 = OK, 1 = invalid date entered
'
'           3-19-03 added new paramter to test date reqd
Function mVerifyDate(edcDate As control, llValidDate As Long, ilDateReqd As Integer) As Integer
Dim slDate As String
    llValidDate = 0
    mVerifyDate = 0
    slDate = edcDate.Text
    If (slDate <> "") Then              'date isnt reqd
        If Not gValidDate(slDate) Then
            mReset
            edcDate.SetFocus
            mVerifyDate = -1
            Exit Function
        Else
            llValidDate = gDateValue(slDate)
        End If
    Else                    'no date entered, is it reqd?
        If ilDateReqd Then
            'edcDate.SetFocus        'show focus on field that needs input
            mVerifyDate = -1
        End If
    End If
End Function
