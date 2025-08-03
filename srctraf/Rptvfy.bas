Attribute VB_Name = "RPTVFY"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfy.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  sgRptSelAgencyCodeTag         sgRptSelSalespersonCodeTag                              *
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
'Public lgNowTime As Long
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
Public tgRptSelAgencyCode() As SORTCODE
Public tgRptSelAgencyCodeCt() As SORTCODE
Public tgRptSelSalespersonCode() As SORTCODE
Public tgRptSelSalespersonCodeCt() As SORTCODE
'11/2/11: Moved to RptRec.Bas
'Public tgRptSelAdvertiserCode() As SORTCODE
Public sgRptSelAdvertiserCodeTag As String
Public tgRptSelAdvertiserCodeCb() As SORTCODE
Public tgRptSelAdvertiserCodeCt() As SORTCODE
Public tgRptSelNameCode() As SORTCODE
Public sgRptSelNameCodeTag As String
Public tgRptSelNameCodePP() As SORTCODE
Public tgRptSelBudgetCode() As SORTCODE
Public sgRptSelBudgetCodeTag As String
Public tgRptSelBudgetCodeAP() As SORTCODE
Public tgRptSelBudgetCodeCB() As SORTCODE
Public tgRptSelBudgetCodeCT() As SORTCODE
Public sgRptSelBudgetCodeTagCT As String
Public tgRptSelBudgetCodePS() As SORTCODE
Public sgRptSelBudgetCodeTagPS As String
Public tgRptSelBudgetCodeSP() As SORTCODE
Public sgRptSelBudgetCodeTagSP As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
Public tgRptSelDemoCodeCB() As SORTCODE
Public tgRptSelDemoCodeCP() As SORTCODE
Public sgRptSelDemoCodeTagCP As String
Public tgRptSelDemoCodeCT() As SORTCODE
Public sgRptSelDemoCodeTagCT As String
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
'Not used
'Dim lmStartDates() As Long          'array of 13 bdcst or corp start dates
'Dim lmEndDates() As Long            'array of 13 bdcst or corp end dates
Public Const HardCost = 0
Public Const Airtime = 1
Public Const NTR = 2
Private Const Correct = 0
Private Const Incorrect = -1
'
'**************************************************************
'*                                                            *
'*      Procedure Name:gGenReport                             *
'*                                                            *
'*             Created:6/16/93       By:D. LeVine             *
'*            Modified:              By:                      *
'*                                                            *
'*         Comments: Formula setups for Crystal               *
'*
'*          Return : 0 =  either error in input, stay in      *
'*                   -1 = error in Crystal, return to         *
'*                        calling program                     *
''*                       failure of gSetformula or another   *
'*                    1 = Crystal successfully completed      *
'*                    2 = successful Bridge                   *
'*          3/24/99 Exclude all types except Cash transactions*
'                   for Statements                            *
'       4/2/99 Chg merch/promo % from 3 to 2 dec. places      *
'*      12-10-99 dh option to show overdue accounts only in
'*                  Credit Status report
'*      9-5-03 Code for windowed envelope for Statements
'*      11-11-03 Test for non-payee transactions (direct advt chaged
'*               to reg agency
'*      3-21-05 Test if hard cost included in report; show in report
'               headers
'**************************************************************
Function gCmcGen(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String, Optional slYear As String = "", Optional slMonth As String = "", Optional slDay As String = "", Optional slTime As String = "") As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilBlanksBeforeLogo            ilBlanksAfterLogo                                       *
'******************************************************************************************

    Dim illoop As Integer
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
    Dim slCashFrom As String
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
    'Dim slTime As String
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
    Dim hlCxf As Integer
    Dim tlCxf As CXF
    Dim tlCxfSrchKey0 As LONGKEY0
    
    
    gCmcGen = 0
    Select Case igRptCallType
        Case SALESPEOPLELIST
            If ilListIndex = 1 Then
                slSelection = ""
                If Not (RptSel!ckcAll.Value = vbChecked) Then
                    For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
                        If RptSel!lbcSelection(1).Selected(illoop) Then
                            If slSelection <> "" Then
                                slSelection = slSelection & " Or " & "{MNF_Multi_Names.mnfName} =" & "'" & Trim$(RptSel!lbcSelection(1).List(illoop)) & "'"
                            Else
                                slSelection = "{MNF_Multi_Names.mnfName} =" & "'" & Trim$(RptSel!lbcSelection(1).List(illoop)) & "'"
                            End If
                        End If
                    Next illoop
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = 0 Then         'salespeople summary option
            
                If RptSel!ckcSelC7.Value = vbChecked Then         'include dormant
                    If Not gSetFormula("IncludeDormant", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("IncludeDormant", "'N'") Then       'Exclude dormant
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                If RptSel!ckcTrans.Value = vbChecked Then                   'include Commission Info
                    If Not gSetFormula("IncludeCommInfo", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("IncludeCommInfo", "'N'") Then       'Exclude Commission Info
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                slSelection = ""
                If Not (RptSel!ckcAll.Value = vbChecked) Then
                    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                        If RptSel!lbcSelection(0).Selected(illoop) Then
                            slNameCode = tgSalesperson(illoop).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                            slNameCode = RptSel!lbcSelection(0).List(ilLoop)
'                            ilRet = gParseItem(slNameCode, 1, ",", slLastName)
'                            ilRet = gParseItem(slNameCode, 2, ",", slFirstName)
                            If slSelection <> "" Then
                           '     slSelection = slSelection & " Or " & "{@Name} =" & "'" & Trim$(slName) & "'"
                                 slSelection = slSelection & " Or {SLF_Salespeople.slfCode} = " & Val(slCode)
                            Else
                           '     slSelection = "{@Name} =" & "'" & Trim$(slName) & "'"
                                 slSelection = " {SLF_Salespeople.slfCode} = " & Val(slCode)
                            End If
                        End If
                    Next illoop
                End If
                If RptSel!ckcSelC7.Value = vbUnchecked Then               'exclude dormant
                    If Trim$(slSelection) = "" Then
                        slSelection = "{SLF_Salespeople.slfState} <> " & "'D'"
                    Else
                        slSelection = "(" & slSelection & ")" & " and {SLF_Salespeople.slfState} <> " & "'D'"
                    End If
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            End If
        Case EVENTNAMESLIST
                slSelection = ""
                If Not (RptSel!ckcAll.Value = vbChecked) Then
                    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                        If RptSel!lbcSelection(0).Selected(illoop) Then
                            slName = RptSel!lbcSelection(0).List(illoop)
                            If slSelection <> "" Then
                                slSelection = slSelection & " Or " & "{ETF_Event_Types.etfName} =" & "'" & Trim$(RptSel!lbcSelection(0).List(illoop)) & "'"
                            Else
                                slSelection = "{ETF_Event_Types.etfName} =" & "'" & Trim$(RptSel!lbcSelection(0).List(illoop)) & "'"
                            End If
                        End If
                    Next illoop
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
       Case PROGRAMMINGJOB
            If igRptType = 3 Then           'program reports (vs links)
                'verify date input
                'Date: 10/4/2018 if SHOW PROGRAM START TIMES then restrict start and end date to 1 week
                'Warning: Dates must be within same Mo-Sun week
                If RptSel!ckcShowTimes.Value = vbChecked Then
                    If RptSel!CSI_CalFrom.Text = "" Then        'mandatory
'                    If RptSel!edcSelCFrom.Text = "" Then        'mandatory
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
'                    If RptSel!edcSelCFrom1.Text = "" Then       'mandatory
                    If RptSel!CSI_CalTo.Text = "" Then       'mandatory
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                    'validate Start date
'                    slDate = Trim$(RptSel!edcSelCFrom.Text)
                    slDate = Trim$(RptSel!CSI_CalFrom.Text)
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                    'validate End date
'                    slDate = Trim$(RptSel!edcSelCFrom1.Text)
                    slDate = Trim$(RptSel!CSI_CalTo.Text)
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                    'make sure start / end dates are within a week Monday - Sunday (make sure it's only for 1 week)
'                    If (((CDate(RptSel!edcSelCFrom1.Text) - CDate(RptSel!edcSelCFrom.Text)) > 6) Or _
'                        (DatePart("w", CDate(RptSel!edcSelCFrom.Text)) = vbSunday) Or _
'                        ((DatePart("w", CDate(RptSel!edcSelCFrom.Text)) > vbSunday) And (DatePart("w", CDate(RptSel!edcSelCFrom1.Text)) <> vbSunday) And _
'                        (DatePart("w", CDate(RptSel!edcSelCFrom.Text)) > DatePart("w", CDate(RptSel!edcSelCFrom1.Text))))) Then
                    If (((CDate(RptSel!CSI_CalTo.Text) - CDate(RptSel!CSI_CalFrom.Text)) > 6) Or _
                        (DatePart("w", CDate(RptSel!CSI_CalFrom.Text)) = vbSunday) Or _
                        ((DatePart("w", CDate(RptSel!CSI_CalFrom.Text)) > vbSunday) And (DatePart("w", CDate(RptSel!CSI_CalTo.Text)) <> vbSunday) And _
                        (DatePart("w", CDate(RptSel!CSI_CalFrom.Text)) > DatePart("w", CDate(RptSel!CSI_CalTo.Text))))) Then
                            RptSel!lacWarning.Visible = True
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                    End If
                End If
                
                'if Active, start date mandatory, end date can be open
                If RptSel!rbcSelC4(0).Value = True Then
'                    If RptSel!edcSelCFrom.Text = "" Then        'must be mandatory
                    If RptSel!CSI_CalFrom.Text = "" Then        'must be mandatory
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    Else
                        slDate = Trim$(RptSel!CSI_CalFrom.Text)
                        If Not gValidDate(slDate) Then
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                        End If
                    End If
                    
'                    If RptSel!edcSelCFrom1.Text <> "" Then        'Optional, can leave open for tfn
                    If RptSel!CSI_CalTo.Text <> "" Then        'Optional, can leave open for tfn
                        slDate = Trim$(RptSel!CSI_CalTo.Text)
                        If Not gValidDate(slDate) Then
                            mReset
                            RptSel!CSI_CalTo.SetFocus
                            Exit Function
                        End If
                    End If
                Else            'expired, start date optional, end date mandatory
                    If RptSel!CSI_CalFrom.Text <> "" Then        'optional, can get expired since the beginning of time
                        slDate = Trim$(RptSel!CSI_CalFrom.Text)
                        If Not gValidDate(slDate) Then
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                        End If
                    End If
                    
'                    If RptSel!edcSelCFrom1.Text = "" Then        'mandatory
                    If RptSel!CSI_CalTo.Text = "" Then        'mandatory
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        Exit Function
                    Else
                        slDate = Trim$(RptSel!CSI_CalTo.Text)
                        If Not gValidDate(slDate) Then
                            mReset
                            RptSel!CSI_CalTo.SetFocus
                            Exit Function
                        End If
                    End If
                End If
                'formulas for report
                If RptSel!rbcSelC4(0).Value = True Then         'active
                    If Not gSetFormula("LibraryType", "'A'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("LibraryType", "'E'") Then       'expired
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                'set formula to display Library Days / Times in the report
                If RptSel!ckcShowTimes.Value = vbChecked Then
                    If Not gSetFormula("ShowDaysTimes", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "DatesRequested")
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If

                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            Else                            'links
'                If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Then   'Selling to airing or Conflict
                If (ilListIndex = PRG_SELLTOAIR) Or (ilListIndex = PRG_AIRTOSELL) Or (ilListIndex = PRG_VEHAVAILCONFLICT) Then   'Selling to airing or Conflict
                    slSelection = ""
                    'If rbcRptType(2).Value Then 'Conflict
'                    If ilListIndex = 2 Then
                    If ilListIndex = PRG_VEHAVAILCONFLICT Then
'                        slDate = RptSel!edcSelA.Text
                        slDate = RptSel!CSI_CalDateA.Text           '12-11-19 change to use csi calendar control
                        If gValidDate(slDate) Then
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            If Not gSetFormula("As Of Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                            'Report!crcReport.Formulas(0) = "As Of Date= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                        Else
                            mReset
                            RptSel!CSI_CalDateA.SetFocus
                            Exit Function
                        End If
                        slSelection = "({VCF_Vehicle_Conflict.vcfTermDate} = Date(0,0,0) Or ({VCF_Vehicle_Conflict.vcfTermDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") And {VCF_Vehicle_Conflict.vcfEffDate} <= {VCF_Vehicle_Conflict.vcfEffDate}))"
                        If (Not RptSel!ckcSel2(0).Value = vbChecked) Or (Not RptSel!ckcSel2(1).Value = vbChecked) Or (Not RptSel!ckcSel2(2).Value = vbChecked) Then
                            If slSelection <> "" Then
                                slSelection = "(" & slSelection & ") " & " And ("
                                slOr = ""
                            Else
                                slSelection = "("
                                slOr = ""
                            End If
                            If (RptSel!ckcSel2(0).Value = vbChecked) Then
                                slSelection = slSelection & slOr & "{VCF_Vehicle_Conflict.vcfSellDay} = 0"
                                slOr = " Or "
                            End If
                            If (RptSel!ckcSel2(1).Value = vbChecked) Then
                                slSelection = slSelection & slOr & "{VCF_Vehicle_Conflict.vcfSellDay} = 6"
                                slOr = " Or "
                            End If
                            If (RptSel!ckcSel2(2).Value = vbChecked) Then
                                slSelection = slSelection & slOr & "{VCF_Vehicle_Conflict.vcfSellDay} = 7"
                            End If
                            slSelection = slSelection & ")"
                        End If
                        If Not (RptSel!ckcAll.Value = vbChecked) Then
                            If slSelection <> "" Then
                                slSelection = "(" & slSelection & ") " & " And ("
                                slOr = ""
                            Else
                                slSelection = "("
                                slOr = ""
                            End If
                            For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                                If RptSel!lbcSelection(0).Selected(illoop) Then
                                    slNameCode = tgSellNameCode(illoop).sKey 'RptSel!lbcSellNameCode.List(ilLoop)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    slSelection = slSelection & slOr & "{VCF_Vehicle_Conflict.vcfSellCode} = " & Trim$(slCode)
                                    slOr = " Or "
                                End If
                            Next illoop
                            slSelection = slSelection & ")"
                        End If
                        If Not gSetSelection(slSelection) Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else        'selling to airing or airing to selling
'                        slDate = RptSel!edcSelA.Text
                        slDate = RptSel!CSI_CalDateA.Text       '12-11-19 change to use csi calendar control
                        If gValidDate(slDate) Then
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            If Not gSetFormula("As Of Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                            'Report!crcReport.Formulas(0) = "As Of Date= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                        Else
                            mReset
                            RptSel!CSI_CalDateA.SetFocus
                            Exit Function
                        End If
                        
                        '7-25-14 option to include avail lengths associated with selling avails
                        If RptSel!ckcInclCommentsA.Value = vbChecked Then           'links report to show avail lengths
                            If ilListIndex = 0 Then
                                If Not gSetFormula("WhichLink", "'S'") Then
                                    gCmcGen = -1
                                    Exit Function
                                End If
                            Else
                                If Not gSetFormula("WhichLink", "'A'") Then
                                    gCmcGen = -1
                                    Exit Function
                                End If
                            End If
                            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                            slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                            slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                            If Not gSetSelection(slSelection) Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            slSelection = "({VLF_Vehicle_Linkages.vlfTermDate} = Date(0,0,0) Or ({VLF_Vehicle_Linkages.vlfTermDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") And {VLF_Vehicle_Linkages.vlfEffDate} <= {VLF_Vehicle_Linkages.vlfEffDate}))"
                            If Not (RptSel!ckcSel1(0).Value = vbChecked) Or Not (RptSel!ckcSel1(1).Value = vbChecked) Then
                                If (RptSel!ckcSel1(0).Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = slSelection & " And " & "{VLF_Vehicle_Linkages.vlfStatus} = 'C'"
                                    Else
                                        slSelection = "{VLF_Vehicle_Linkages.vlfStatus} = 'C'"
                                    End If
                                End If
                                If (RptSel!ckcSel1(1).Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = slSelection & " And " & "{VLF_Vehicle_Linkages.vlfStatus} = 'P'"
                                    Else
                                        slSelection = "{VLF_Vehicle_Linkages.vlfStatus} = 'P'"
                                    End If
                                End If
                            End If
                            'If rbcRptType(0).Value Then 'Selling to Airing
                            If ilListIndex = 0 Then
                                If (Not RptSel!ckcSel2(0).Value = vbChecked) Or (Not RptSel!ckcSel2(1).Value = vbChecked) Or (Not RptSel!ckcSel2(2).Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = "(" & slSelection & ") " & " And ("
                                        slOr = ""
                                    Else
                                        slSelection = "("
                                        slOr = ""
                                    End If
                                    If (RptSel!ckcSel2(0).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfSellDay} = 0"
                                        slOr = " Or "
                                    End If
                                    If (RptSel!ckcSel2(1).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfSellDay} = 6"
                                        slOr = " Or "
                                    End If
                                    If (RptSel!ckcSel2(2).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfSellDay} = 7"
                                    End If
                                    slSelection = slSelection & ")"
                                End If
                                If Not (RptSel!ckcAll.Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = "(" & slSelection & ") " & " And ("
                                        slOr = ""
                                    Else
                                        slSelection = "("
                                        slOr = ""
                                    End If
                                    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                                        If RptSel!lbcSelection(0).Selected(illoop) Then
                                            slNameCode = tgSellNameCode(illoop).sKey 'RptSel!lbcSellNameCode.List(ilLoop)
                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                            slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfSellCode} = " & Trim$(slCode)
                                            slOr = " Or "
                                        End If
                                    Next illoop
                                    slSelection = slSelection & ")"
                                End If
                                If Not gSetSelection(slSelection) Then
                                    gCmcGen = -1
                                    Exit Function
                                End If
                            Else    'Airing to Selling
                                If (Not RptSel!ckcSel2(0).Value = vbChecked) Or (Not RptSel!ckcSel2(1).Value = vbChecked) Or (Not RptSel!ckcSel2(2).Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = "(" & slSelection & ") " & " And ("
                                        slOr = ""
                                    Else
                                        slSelection = "("
                                        slOr = ""
                                    End If
                                    If (RptSel!ckcSel2(0).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfAirDay} = 0"
                                        slOr = " Or "
                                    End If
                                    If (RptSel!ckcSel2(1).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfAirDay} = 6"
                                        slOr = " Or "
                                    End If
                                    If (RptSel!ckcSel2(2).Value = vbChecked) Then
                                        slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfAirDay} = 7"
                                    End If
                                    slSelection = slSelection & ")"
                                End If
                                If Not (RptSel!ckcAll.Value = vbChecked) Then
                                    If slSelection <> "" Then
                                        slSelection = "(" & slSelection & ") " & " And ("
                                        slOr = ""
                                    Else
                                        slSelection = "("
                                        slOr = ""
                                    End If
                                    For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
                                        If RptSel!lbcSelection(1).Selected(illoop) Then
                                            slNameCode = tgAirNameCode(illoop).sKey    'RptSel!lbcAirNameCode.List(ilLoop)
                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                            slSelection = slSelection & slOr & "{VLF_Vehicle_Linkages.vlfAirCode} = " & Trim$(slCode)
                                            slOr = " Or "
                                        End If
                                    Next illoop
                                    slSelection = slSelection & ")"
                                End If
                                If Not gSetSelection(slSelection) Then
                                    gCmcGen = -1
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
'                ElseIf (ilListIndex = 3) Or (ilListIndex = 4) Then   'delivery
                ElseIf (ilListIndex = PRG_DELIVERY_BYVEH) Or (ilListIndex = PRG_DELIVERY_BYFEED) Then   'delivery
                    slSelection = ""
'                    slDate = RptSel!edcSelA.Text
                    slDate = RptSel!CSI_CalDateA.Text
                    If gValidDate(slDate) Then
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If Not gSetFormula("As Of Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        'Report!crcReport.Formulas(0) = "As Of Date= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                    Else
                        mReset
                        RptSel!CSI_CalDateA.SetFocus
                        Exit Function
                    End If
                    If (RptSel!ckcSel1(0).Value = vbChecked) Then
                        If slSelection = "" Then
                            slSelection = "{DLF_Delivery_Links.dlfCmmlSched} = 'Y'"
                        Else
                            slSelection = slSelection & " Or {DLF_Delivery_Links.dlfCmmlSched} = 'Y'"
                        End If
                    End If
                    'If by subfeed
                    If RptSel!ckcSel1(1).Value = vbChecked Then
                        If slSelection = "" Then
                            slSelection = "{DLF_Delivery_Links.dlfZone} = 'CST'"
                        Else
                            slSelection = slSelection & " Or {DLF_Delivery_Links.dlfZone} = 'CST'"
                        End If
                    End If
                    If (Not RptSel!ckcSel2(0).Value = vbChecked) Or (Not RptSel!ckcSel2(1).Value = vbChecked) Or (Not RptSel!ckcSel2(2).Value = vbChecked) Then
                        If slSelection <> "" Then
                            slSelection = "(" & slSelection & ") " & " And ("
                            slOr = ""
                        Else
                            slSelection = "("
                            slOr = ""
                        End If
                        If RptSel!ckcSel2(0).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{DLF_Delivery_Links.dlfAirDay} = '0'"
                            slOr = " Or "
                        End If
                        If RptSel!ckcSel2(1).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{DLF_Delivery_Links.dlfAirDay} = '6'"
                            slOr = " Or "
                        End If
                        If RptSel!ckcSel2(2).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{DLF_Delivery_Links.dlfAirDay} = '7'"
                        End If
                        slSelection = slSelection & ")"
                    End If
                    If Not RptSel!ckcAll.Value = vbChecked Then
                        If slSelection <> "" Then
                            slSelection = "(" & slSelection & ") " & " And ("
                            slOr = ""
                        Else
                            slSelection = "("
                            slOr = ""
                        End If
                        'If rbcRptType(0).Value Then 'Airing/Conventional vehicles
                        If ilListIndex = 3 Then
                            For illoop = 0 To RptSel!lbcSelection(3).ListCount - 1 Step 1
                                If RptSel!lbcSelection(3).Selected(illoop) Then
                                    slNameCode = tgVehicle(illoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    slSelection = slSelection & slOr & "{DLF_Delivery_Links.dlfvefCode} = " & Trim$(slCode)
                                    slOr = " Or "
                                End If
                            Next illoop
                        Else
                            For illoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                                If RptSel!lbcSelection(2).Selected(illoop) Then
                                    slNameCode = tgRptSelNameCode(illoop).sKey 'RptSel!lbcNameCode.List(ilLoop)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    slSelection = slSelection & slOr & "{DLF_Delivery_Links.dlfmnfFeed} = " & Trim$(slCode)
                                    slOr = " Or "
                                End If
                            Next illoop
                        End If
                        slSelection = slSelection & ")"
                    End If
                    'Add Removal of terminated delivery records
                    If slSelection <> "" Then
                        slSelection = "(" & slSelection & ")" & " And "
                    End If
                    slSelection = slSelection & "(({DLF_Delivery_Links.dlfTermDate} = Date(0,0,0)) Or ({DLF_Delivery_Links.dlfTermDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") And {DLF_Delivery_Links.dlfStartDate} <= {DLF_Delivery_Links.dlfTermDate}))"
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If
'                ElseIf (ilListIndex = 5) Or (ilListIndex = 6) Then   'Engineering
                ElseIf (ilListIndex = PRG_ENG_BYVEH) Or (ilListIndex = PRG_ENG_BYFEED) Then   'Engineering
                    slSelection = ""
'                    slDate = RptSel!edcSelA.Text
                    slDate = RptSel!CSI_CalDateA.Text           '12-11-19 change to use csi calendar control
                    If gValidDate(slDate) Then
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If Not gSetFormula("As Of Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        'Report!crcReport.Formulas(0) = "As Of Date= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                    Else
                        mReset
                        RptSel!CSI_CalDateA.SetFocus
                        Exit Function
                    End If
                    If RptSel!ckcSel1(0).Value = vbChecked Then
                        If slSelection = "" Then
                            slSelection = "{EGF_Engr_Links.egfFed} = 'Y'"
                        Else
                            slSelection = slSelection & " Or {EGF_Engr_Links.egfFed} = 'Y'"
                        End If
                    End If
                    'If by subfeed
                    If RptSel!ckcSel1(1).Value = vbChecked Then
                        If slSelection = "" Then
                            slSelection = "{EGF_Engr_Links.egfZone} = 'CST'"
                        Else
                            slSelection = slSelection & " Or {EGF_Engr_Links.egfZone} = 'CST'"
                        End If
                    End If
                    If (Not RptSel!ckcSel2(0).Value = vbChecked) Or (Not RptSel!ckcSel2(1).Value = vbChecked) Or (Not RptSel!ckcSel2(2).Value = vbChecked) Then
                        If slSelection <> "" Then
                            slSelection = "(" & slSelection & ") " & " And ("
                            slOr = ""
                        Else
                            slSelection = "("
                            slOr = ""
                        End If
                        If RptSel!ckcSel2(0).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{EGF_Engr_Links.egfAirDay} = '0'"
                            slOr = " Or "
                        End If
                        If RptSel!ckcSel2(1).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{EGF_Engr_Links.egfAirDay} = '6'"
                            slOr = " Or "
                        End If
                        If RptSel!ckcSel2(2).Value = vbChecked Then
                            slSelection = slSelection & slOr & "{EGF_Engr_Links.egfAirDay} = '7'"
                        End If
                        slSelection = slSelection & ")"
                    End If
                    If Not RptSel!ckcAll.Value = vbChecked Then
                        If slSelection <> "" Then
                            slSelection = "(" & slSelection & ") " & " And ("
                            slOr = ""
                        Else
                            slSelection = "("
                            slOr = ""
                        End If
                        'If rbcRptType(0).Value Then 'Airing/Conventional vehicles
                        If ilListIndex = 5 Then
                            For illoop = 0 To RptSel!lbcSelection(3).ListCount - 1 Step 1
                                If RptSel!lbcSelection(3).Selected(illoop) Then
                                    slNameCode = tgVehicle(illoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    slSelection = slSelection & slOr & "{EGF_Engr_Links.egfvefCode} = " & Trim$(slCode)
                                    slOr = " Or "
                                End If
                            Next illoop
                        Else
                            For illoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                                If RptSel!lbcSelection(2).Selected(illoop) Then
                                    slNameCode = tgRptSelNameCode(illoop).sKey 'RptSel!lbcNameCode.List(ilLoop)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                    slSelection = slSelection & slOr & "{EGF_Engr_Links.egfmnfFeed} = " & Trim$(slCode)
                                    slOr = " Or "
                                End If
                            Next illoop
                        End If
                        slSelection = slSelection & ")"
                    End If
                    'Add Removal of terminated delivery records
                    If slSelection <> "" Then
                        slSelection = "(" & slSelection & ")" & " And "
                    End If
                    slSelection = slSelection & "(({EGF_Engr_Links.egfTermDate} = Date(0,0,0)) Or ({EGF_Engr_Links.egfTermDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") And {EGF_Engr_Links.egfStartDate} <= {EGF_Engr_Links.egfTermDate}))"
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If
                ElseIf ilListIndex = PRG_AIRING_INV Then            '3-31-15
'                    ilRet = mVerifyDate(RptSel!edcSelA, llDate1, True)
                    ilRet = mVerifyDate(RptSel!CSI_CalDateA, llDate1, True)         '12-11-19 change to use csi calendar control
                    If ilRet = 0 Then
                        'backup to Monday
                        
                        ilFormulaNo = gWeekDayLong(llDate1)
                        
                        Do While ilFormulaNo <> 0           'backup MF to monday
                            llDate1 = llDate1 - 1
                            ilFormulaNo = gWeekDayLong(llDate1)
                        Loop
                        slStr = Format$(llDate1, "m/d/yy")
                        ilRet = mSendAsOfDateFormula(slStr)
                        If ilRet <> 0 Then
                            gCmcGen = 1
                        End If
                    End If
                    
                    'send days selected
                    slStr = ""
                    If RptSel!ckcSel2(0).Value = vbChecked Then     'mo-fr
                        slStr = "Mon-Fri"
                    End If
                    If RptSel!ckcSel2(1).Value = vbChecked Then
                        If Trim$(slStr) = "" Then
                            slStr = "Sat"
                        Else
                            slStr = slStr & ", Sat"
                        End If
                    End If
                    If RptSel!ckcSel2(2).Value = vbChecked Then
                        If Trim$(slStr) = "" Then
                            slStr = "Sun"
                        Else
                            slStr = slStr & ", Sun"
                        End If
                    End If
                    If Not gSetFormula("HeaderWhichDays", "'" & slStr & "'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    
                    If RptSel!ckcInclCommentsA.Value = vbChecked Then
                        slStr = ""
                        If RptSel!ckcADate.Value = vbChecked Then
                            slStr = "Discrepancy Only"
                        End If
                        If Not gSetFormula("DiscrepancyOnly", "'" & slStr & "'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                                        
                    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If

                End If
            End If                  'reports ( vs links)
        Case COLLECTIONSJOB
            If ilListIndex = COLL_PAYHISTORY Or ilListIndex = COLL_DISTRIBUTE Or ilListIndex = COLL_ACCTHIST Then    'payment history or cash distribution
'                If (RptSel!edcSelCFrom.Text <> "") And (RptSel!edcSelCTo.Text <> "") Then
                If (RptSel!CSI_CalFrom.Text <> "") And (RptSel!CSI_CalTo.Text <> "") Then       '8-28-19 use csi calendar control vs edit box
                    If StrComp(RptSel!CSI_CalTo.Text, "TFN", 1) <> 0 Then
                        slDate = RptSel!CSI_CalFrom.Text
                        If gValidDate(slDate) Then
                            slDate = RptSel!CSI_CalTo.Text
                            If gValidDate(slDate) Then
                            Else
                                mReset
                                RptSel!CSI_CalFrom.SetFocus
                                Exit Function
                            End If
                        Else
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                        End If
                    Else
                        slDate = RptSel!CSI_CalFrom.Text
                        If gValidDate(slDate) Then
                        Else
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                        End If
                    End If
                ElseIf RptSel!CSI_CalFrom.Text <> "" Then
                    slDate = RptSel!CSI_CalFrom.Text
                    If gValidDate(slDate) Then
                    Else
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                ElseIf RptSel!CSI_CalFrom.Text <> "" Then
                    If StrComp(RptSel!CSI_CalFrom.Text, "TFN", 1) <> 0 Then
                        slDate = RptSel!CSI_CalFrom.Text
                        slDateTo = slDate
                        If gValidDate(slDate) Then
                        Else
                            mReset
                            RptSel!CSI_CalFrom.SetFocus
                            Exit Function
                        End If
                    End If
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"

                slDateFrom = RptSel!CSI_CalFrom.Text                'Dates entered, acct hist may be blank
                slDateTo = RptSel!CSI_CalTo.Text
                If RptSel!rbcOutput(3).Value = False Then
                    If slDateFrom <> "" And slDateTo <> "" Then            'Start & end dates entered
                        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
                        'slDateTo = rptSel!edcSelCTo.Text
                        slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
                        If Not gSetFormula("InputDates", "'" & slDateFrom & "-" & slDateTo & "'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    ElseIf slDateFrom = "" And slDateTo = "" Then           'no dates entered
                        If Not gSetFormula("InputDates", "'All Dates'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    ElseIf slDateFrom = "" Then                             'only end date entred
                        slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
                        If Not gSetFormula("InputDates", "'Thru " & slDateTo & "'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else                                                    'only start date entred
                        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'makesure year included
                        If Not gSetFormula("InputDates", "'From " & slDateFrom & "'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    If ilListIndex = COLL_DISTRIBUTE And RptSel!rbcSelCSelect(2).Value Then 'if cash distribution by participant, exclude PO
                        If RptSel!ckcTrans.Value = vbChecked Then        'skip to new page
                            If Not gSetFormula("NewPage", "'Y'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("NewPage", "'N'") Then    'no page skips
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If                    '2-5-15 option to skip to new page each vehicle within participant
                        
                        slSelection = slSelection & " And {RVR_Receivables_Rept.rvrTranType} <> 'PO'"
                        'need to send the earliest date requsted to crystal for prior month distribution flag
                        gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
                        If Not gSetFormula("PriorDateFlag", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If
    
                    If ilListIndex = COLL_PAYHISTORY Then           'send title of report for cash, trade, merch, promotions
                        If Not mTitleCashTrMercProm() Then               'send Crystal reporttitle
                            gCmcGen = -1
                        End If
                        If Not mShowTransComments(RptSel!ckcSelC7) Then        'send formula to show/not show trans comments
                            gCmcGen = -1
                        End If
                    End If
                    If ilListIndex = COLL_ACCTHIST Then           'Account history
                        If Not mShowTransComments(RptSel!ckcSelC5(0)) Then        'send formula to show/not show trans comments
                            gCmcGen = -1
                        End If
    
    
                       If RptSel!ckcSelC7.Value = vbChecked Then    '5-7-04
                            If Not gSetFormula("ShowParticipant", "'Y'") Then 'Show sales source & participants
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("ShowParticipant", "'N'") Then 'dont show sales source & participant
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                        
                        '8-1-17 option include user name (important for the exports)
                        If RptSel!ckcOption.Value = vbChecked Then
                            If Not gSetFormula("IncludeUserName", "'Y'") Then 'Show username
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("IncludeUserName", "'N'") Then 'dont show username
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
    
                        If RptSel!rbcSelCSelect(0).Value Then       'advertiser
                            If Not gSetFormula("SortBy", "'A'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("SortBy", "'G'") Then    'agency
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
    
                        If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                            If Not gSetFormula("UsingTaxes", "'Y'") Then    'using taxes
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("UsingTaxes", "'N'") Then    'not using taxes
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
    
                        If RptSel!ckcSelC10(0).Value = vbChecked Then        'skip to new page
                            If Not gSetFormula("NewPage", "'Y'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("NewPage", "'N'") Then    'no page skips
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    'End If
    
                        '5-18-09 Show which file in header
                        If RptSel!rbcSelC12(0).Value Then       'history only
                            If Not gSetFormula("WhichFileHeader", "'History Only'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        ElseIf RptSel!rbcSelC12(1).Value Then       'receivables only
                            If Not gSetFormula("WhichFileHeader", "'Receivables Only'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("WhichFileHeader", "'History & Receivables'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    End If
                End If
            'All Ageings
            ElseIf (ilListIndex = COLL_AGEPAYEE) Or (ilListIndex = COLL_AGESLSP) Or (ilListIndex = COLL_AGEVEHICLE) Or (ilListIndex = COLL_AGEOWNER) Or (ilListIndex = COLL_AGESS) Or (ilListIndex = COLL_AGEPRODUCER) Then         '2-10-00 ageing
                '8-29-19 use csi calendar control vs edit box
'                If RptSel!edcSelCFrom.Text <> "" Then
                 If RptSel!CSI_CalFrom.Text <> "" Then
                   slDateFrom = RptSel!CSI_CalFrom.Text
                    If Not gValidDate(slDateFrom) Then
'                        slDateFrom = slDate
'                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                        8-27-15 all versions have a prepass
'                        If (ilListIndex = COLL_AGESLSP) Then   'All ageing versions except slsp are prepass
'                            slSelection = "{RVF_Receivables.rvfTranDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'                        End If
'                    Else
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                Else            'last bill date to include is mandatory
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    Exit Function
                End If
                
'                If RptSel!edcLatestCashDate.Text <> "" Then
                If RptSel!CSI_CalTo.Text <> "" Then
                    slCashFrom = RptSel!CSI_CalTo.Text
                    If Not gValidDate(slCashFrom) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus       '5-20-19
                        Exit Function
                    End If
                Else            'latest cash to include is mandatory
                    mReset
                    RptSel!CSI_CalTo.SetFocus       '5-20-19
                    Exit Function
                End If
                
                If RptSel!edcSelCFrom1.Text <> "" Then
                    ilRet = mVerifyMMYY(RptSel!edcSelCFrom1, slMonthCurr, slYearCurr)
                    If ilRet <> CP_MSG_NONE Then
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!edcSelCFrom1.SetFocus
                    Exit Function
                End If

                'no ageing period entered assumes 01/1970
                If RptSel!edcSelCTo.Text <> "" Then     'verify mm/yy entered for earliest ageing date
                    ilRet = mVerifyMMYY(RptSel!edcSelCTo, slEarliestMM, slEarliestYY)
                    If ilRet <> CP_MSG_NONE Then
                        Exit Function
                    End If
                Else
                    slEarliestMM = "1"
                    slEarliestYY = "1970"
                End If

                'no latest aging period assumes 12/2069
                If RptSel!edcSelCTo1.Text <> "" Then           'verify mm/yy entered for latest aging date
                    ilRet = mVerifyMMYY(RptSel!edcSelCTo1, slLatestMM, slLatestYY)
                    If ilRet <> CP_MSG_NONE Then
                        Exit Function
                    End If
                Else
                    slLatestMM = "12"
                    slLatestYY = "2069"
                End If
'               4-17-11 make ageing by slsp prepass
'                If (ilListIndex = COLL_AGESLSP) Then    'All versions except slsp are prepass
'                    slSelection = "(" & slSelection & ") And (({RVF_Receivables.rvfAgeYear} < " & slYearCurr & ") Or (({RVF_Receivables.rvfAgeYear} = " & slYearCurr & ") And ({RVF_Receivables.rvfAgePeriod} <= " & slMonthCurr & ")))"
'                    slSelection = slSelection & "and ({RVF_Receivables.rvfAgeYear} >= " & slEarliestYY & " and {RVF_Receivables.rvfAgeYear} <= " & slLatestYY & " and {RVF_Receivables.rvfAgePeriod} >= " & slEarliestMM & " and {RVF_Receivables.rvfAgePeriod} <= " & slLatestMM & ")"
'                    ilRet = mSelAirNTRHardCost(slSelection) ' 6/03/08 Dan M
'                    If ilRet <> 0 Then
'                        gCmcGen = -1
'                        Exit Function
'                    End If
'                End If
'
'                '2-10-00
'                If (ilListIndex = COLL_AGESLSP) Then 'Only ageing by slsp needs filtering for cash/trade because not a prepass
'                    mSelCashTrMercProm slSelection                     'send Crystal selection of Cash, Trade, Merchandise, Promotions
'                End If

                '7-8-02 If ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEVEHICLE Then
                'Slsp, VEhicle and Payee Ageings have 4 levels of totals as options
                If ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_AGEPAYEE Then       '7-8-02
                    If RptSel!rbcSelCSelect(0).Value Then       'Detail
                        If Not gSetFormula("TotalLevel", "'D'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not mShowTransComments(RptSel!ckcTrans) Then        'send formula to show/not show trans comments
                            gCmcGen = -1
                        End If
                    ElseIf RptSel!rbcSelCSelect(1).Value Then       'Tran Trype
                        If Not gSetFormula("TotalLevel", "'T'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not mShowTransComments(RptSel!ckcTrans) Then        'send formula to show/not show trans comments
                            gCmcGen = -1
                        End If
                    ElseIf RptSel!rbcSelCSelect(2).Value Then       'invoice
                        If Not gSetFormula("TotalLevel", "'I'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not gSetFormula("ShowTransComments", "'N'") Then 'dont show trans comments
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("TotalLevel", "'S'") Then 'summary by advt within slsp
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not gSetFormula("ShowTransComments", "'N'") Then 'dont show trans comments
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    
'                   4-17-11 make ageing by slsp prepass
'                    If ilListIndex = COLL_AGESLSP Then
'                        mSelCashTrMercProm slSelection                     'send Crystal selection of Cash, Trade, Merchandise, Promotions
'                        If Not RptSel!ckcAll.Value = vbChecked Then         'not all vehicles selected
'                            If slSelection <> "" Then
'                                slSelection = "(" & slSelection & ") " & " and ("
'                                slOr = ""
'                            Else
'                                slSelection = "("
'                                slOr = ""
'                            End If
'
'                            For ilLoop = 0 To RptSel!lbcSelection(5).ListCount - 1 Step 1
'                                If RptSel!lbcSelection(5).Selected(ilLoop) Then
'                                    slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
'                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                                    slSelection = slSelection & slOr & "{RVF_Receivables.rvfslfCode} = " & Trim$(slCode)
'                                    slOr = " Or "
'                                End If
'                            Next ilLoop
'                            'for selective slsp, include those trans. without a slsp reference (PO)
'                            slSelection = slSelection & slOr & "{RVF_Receivables.rvfslfCode} = 0"
'                           slSelection = slSelection & ")"
'                        End If
'
'                        '8-6-10 option for selective sales offices
'                        If Not RptSel!ckcAllGroups.Value = vbChecked Then         'not all offices selected
'                            If slSelection <> "" Then
'                                slSelection = "(" & slSelection & ") " & " and ("
'                                slOr = ""
'                            Else
'                                slSelection = "("
'                                slOr = ""
'                            End If
'
'                            For ilLoop = 0 To RptSel!lbcSelection(7).ListCount - 1 Step 1
'                                If RptSel!lbcSelection(7).Selected(ilLoop) Then
'                                    slNameCode = tgSOCode(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
'                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                                    slSelection = slSelection & slOr & "{SLF_Salespeople.slfsofCode} = " & Trim$(slCode)
'                                    slOr = " Or "
'                                End If
'                            Next ilLoop
'                            'for selective slsp, include those trans. without a slsp reference (PO)
'                            slSelection = slSelection & slOr & "{SLF_Salespeople.slfsofCode}  = 0"
'                           slSelection = slSelection & ")"
'                        End If
'                    End If
                End If

                If ilListIndex = COLL_AGEVEHICLE Then           'show owners share?
                    If RptSel!ckcSelC3(0).Value = vbChecked Then    '9-12-02
                        If Not gSetFormula("ShowOwnersShare", "'Y'") Then 'Show owners participation share
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowOwnersShare", "'N'") Then 'Show owners participation share
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    If RptSel!ckcSelC10(0).Value = vbChecked Then    '11-16-06 include slsp subtotals
                        If Not gSetFormula("IncludeSlspSubTotals", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("IncludeSlspSubTotals", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    '4-11-07 option to skip to new page each slsp when sorting by vehicle
                    If RptSel!ckcSelC10(1).Value = vbChecked Then        'skip to new page
                        If Not gSetFormula("NewPage", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("NewPage", "'N'") Then    'no page skips
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                End If

                'Ageings by Owner, Producer, SAlesSource only have options for Detail & Summary
                If ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Then  '7-8-02       '2-10-00 ageing by owner has filter in pre-pas
                    'send formula whether detail or summary version
                    If RptSel!rbcSelCSelect(0).Value Then       'detail
                        If Not gSetFormula("TotalLevel", "'D'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not mShowTransComments(RptSel!ckcTrans) Then        'send formula to show/not show trans comments
                            gCmcGen = -1
                        End If
                    ElseIf RptSel!rbcSelCSelect(1).Value Then       'summary by advt
                        If Not gSetFormula("TotalLevel", "'S'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        If Not gSetFormula("ShowTransComments", "'N'") Then 'dont show trans comments
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    If ilListIndex = COLL_AGESS Then        '2-10-00
                        If Not gSetFormula("ReportType", "'S'") Then
                            gCmcGen = -1
                            Exit Function
                        End If

                    ElseIf ilListIndex = COLL_AGEPRODUCER Then  '2-10-00
                        If Not gSetFormula("ReportType", "'P'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                End If

                '4-17-11 make ageing by slsp a prepass
                'If ilListIndex <> COLL_AGESLSP Then                     'only ageing by slsp isnt prepass, filter the other options to get correct date & time records genned
                'retrieve the matching date & time genrated records
                gCurrDateTime slHeader, slTime, slGenMonth, slGenDay, slGenYear
                'any selections gathered for any of the other agings are ignored here with the following selection
                slSelection = " ({RVR_Receivables_Rept.rvrGenDate} = Date(" & slGenYear & "," & slGenMonth & "," & slGenDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"

                'End If

                If Not gSetSelection(slSelection) Then  'send selections to any of the ageings
                    gCmcGen = -1
                    Exit Function
                End If
                If Not mTitleCTMPForCbc() Then
                    gCmcGen = -1
                End If
                gPDNToStr tgSpf.sRB, 2, slStr
                If Not gSetFormula("Balance", slStr) Then
                    gCmcGen = -1
                    Exit Function
                End If
                If tgSpf.sRRP = "C" Then    'Calendar
                    slBaseDate = gObtainEndCal(slMonthCurr & "/15/" & slYearCurr)
                ElseIf tgSpf.sRRP = "F" Then 'Corporate
                    slBaseDate = gObtainEndCorp(slMonthCurr & "/15/" & slYearCurr, True)
                Else
                    slBaseDate = gObtainEndStd(slMonthCurr & "/15/" & slYearCurr)
                End If
                gObtainYearMonthDayStr slBaseDate, True, slYear, slMonth, slDay
                If Not gSetFormula("Base Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                'latest billing tran date to include
                gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
                If Not gSetFormula("TranDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If

                'latest cash tran date to include
                gObtainYearMonthDayStr slCashFrom, True, slYear, slMonth, slDay
                If Not gSetFormula("CashTranDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGESLSP Then
                    If RptSel!ckcSelC5(0).Value = vbChecked Then    'sort in collection to the back
                        If Not gSetFormula("SortInCollect", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("SortInCollect", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    If ilListIndex = COLL_AGEPAYEE Then
                        If RptSel!ckcSelC7.Value = vbChecked Then    'separate Sales Source as major
                            If Not gSetFormula("SeparateSS", "'Y'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("SeparateSS", "'N'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    End If
                End If

                'gCurrDateTime slDate, slTime, slMonth, slDay, slYear   'comment out 2-10-00 (out of mem)
                'If Not gSetFormula("AsOfT", Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))) Then
                '    gCmcGen = -1
                '    Exit Function
               ' End If
                'for AGeing by Payee, send formula whether by detail or inv type total levels   '9-7-99
                'If RptSel!rbcSelCSelect(0).Value And ilListIndex = 1 Then
                '    If Not gSetFormula("TotalLevel", "'D'") Then    'detail
                '        gCmcGen = -1
                '        Exit Function
                '    End If
                '
                'ElseIf RptSel!rbcSelCSelect(2).Value And ilListIndex = COLL_AGEPAYEE Then
                '    If Not gSetFormula("TotalLevel", "'I'") Then        'inv totals
                '        gCmcGen = -1
                '        Exit Function
                '    End If
                'End If
            ElseIf ilListIndex = COLL_DELINQUENT Then 'delinquent accounts
                slDateFrom = ""
                'Date selection passed by formula
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date        8-29-19 use csi calendar control vs edit box
                If (slDate <> "") And (slDate <> "TFN") Then
                    If gValidDate(slDate) Then
                        slDateFrom = slDate
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slSelection = "{RVF_Receivables.rvfTranDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    Else
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                End If
                If slDateFrom <> "" Then
                    gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
                    If Not gSetFormula("Last Cash", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                slDateTo = ""
                'Date selection passed by formula
'                slDate = RptSel!edcSelCTo.Text   'Latest cash date
                slDate = RptSel!CSI_CalTo.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    Exit Function
                End If
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slBaseDate = slDate
                If Not gSetFormula("Base Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If
                'Only Cash
                If slSelection <> "" Then
                    slSelection = slSelection & " And {RVF_Receivables.rvfCashTrade} = 'C'"
                Else
                    slSelection = "{RVF_Receivables.rvfCashTrade} = 'C'"
                End If

                'test for airtime, ntr or both
                If RptSel!rbcSelC6(0).Value Then        'air time only
                    slSelection = slSelection & " and {RVF_Receivables.rvfmnfItem} = 0"
                ElseIf RptSel!rbcSelC6(1).Value Then    'ntr only
                    slSelection = slSelection & " and {RVF_Receivables.rvfmnfItem} <> 0"
                End If

                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

                ilRet = mTitleAirTimeNTRHdr(True, True, False) 'parameters:  Ok to combine hard cost with non hard-cost; Include hard costs
                If ilRet <> 0 Then
                    gCmcGen = -1
                End If
            ElseIf ilListIndex = COLL_STATEMENT Then
                '8-29-19 use csi cal control vs edit box
'                slDate = RptSel!edcSelCFrom.Text            'date mandatory
                slDate = RptSel!CSI_CalFrom.Text            'date mandatory
                If slDate <> "" Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    Exit Function
                End If
                
'                slDate = RptSel!edcSelCTo.Text          'date mandatory
                slDate = RptSel!CSI_CalTo.Text          'date mandatory
                If slDate <> "" Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    Exit Function
                End If


            
'                slDateFrom = RptSel!edcSelCFrom.Text   'Latest cash date
                slDateFrom = RptSel!CSI_CalFrom.Text   'Latest cash date
                llDate1 = gDateValue(slDateFrom)
                slDateFrom = Format$(llDate1, "ddddd")
'                slDateTo = RptSel!edcSelCTo.Text   'Latest bill date
                slDateTo = RptSel!CSI_CalTo.Text   'Latest bill date
                llDate1 = gDateValue(slDateTo)
                slDateTo = Format$(llDate1, "ddddd")

'                'Date selection passed by formula
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
'                If (slDate <> "") And (slDate <> "TFN") Then
'                    If gValidDate(slDate) Then
'                        slDateFrom = slDate
'                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                        slSelection = "({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'P') Or ({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'W')"
'                    Else
'                        mReset
'                        RptSel!edcSelCFrom.SetFocus
'                        Exit Function
'                    End If
'                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
'                If (slDate <> "") And (slDate <> "TFN") Then
'                    If gValidDate(slDate) Then
'                        slDateTo = slDate
'                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'                        If slSelection = "" Then
'                            slSelection = "({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'I') Or ({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'A')"
'                        Else
'                            slSelection = slSelection & " Or " & "({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'I') Or ({RVF_Receivables.rvfTranDate} <= Date(" & slYear & ", " & slMonth & ", " & slDay & ") And {RVF_Receivables.rvfTranType}[1] = 'A')"
'                        End If
'                    Else
'                        mReset
'                        RptSel!edcSelCTo.SetFocus
'                        Exit Function
'                    End If
'                End If
'                If slSelection = "" Then
'                    slSelection = "{RVF_Receivables.rvfCashTrade} = 'C' "
'                Else
'                    slSelection = "(" & slSelection & ") " & " And ({RVF_Receivables.rvfCashTrade} = 'C')"
'                End If
'                If Not RptSel!ckcAll.Value = vbChecked Then
'                    If slSelection <> "" Then
'                        slSelection = "(" & slSelection & ") " & " And ("
'                        slOr = ""
'                    Else
'                        slSelection = "("
'                        slOr = ""
'                    End If
'                    For ilLoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
'                        If RptSel!lbcSelection(2).Selected(ilLoop) Then
'                            slNameCode = RptSel!lbcAgyAdvtCode.List(ilLoop)
'                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                            If InStr(slNameCode, "/Direct") <> 0 Or InStr(slNameCode, "/Non-") <> 0 Then
'                                slSelection = slSelection & slOr & "{RVF_Receivables.rvfadfCode} = " & Trim$(slCode)
'                            Else
'                                slSelection = slSelection & slOr & "{RVF_Receivables.rvfagfCode} = " & Trim$(slCode)
'                            End If
'                            slOr = " Or "
'                        End If
'                    Next ilLoop
'                    slSelection = slSelection & ")"
'                End If
'
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"

                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

                If RptSel!edcSelCTo1.Text <> "" Then
                    ilRet = mVerifyMMYY(RptSel!edcSelCTo1, slMonthCurr, slYearCurr)
                    If ilRet <> CP_MSG_NONE Then
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!edcSelCTo1.SetFocus
                    Exit Function
                End If

                If tgSpf.sRRP = "C" Then    'Calendar
                    slBaseDate = gObtainEndCal(slMonthCurr & "/15/" & slYearCurr)
                ElseIf tgSpf.sRRP = "F" Then 'Corporate
                    slBaseDate = gObtainEndCorp(slMonthCurr & "/15/" & slYearCurr, True)
                Else
                    slBaseDate = gObtainEndStd(slMonthCurr & "/15/" & slYearCurr)
                End If
                gObtainYearMonthDayStr slBaseDate, True, slYear, slMonth, slDay

                'gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay

                If Not gSetFormula("Base Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGen = -1
                    Exit Function
                End If
                ilFormulaNo = 0
                If slDateFrom <> "" Then
                    gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
                    If Not gSetFormula("Last Cash", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    ilFormulaNo = ilFormulaNo + 1
                End If
                If slDateTo <> "" Then
                    gObtainYearMonthDayStr slDateTo, True, slYear, slMonth, slDay
                    If Not gSetFormula("Last Bill", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    ilFormulaNo = ilFormulaNo + 1
                End If
                '1-14-11 change to use Payee name vs Client name
                If Not gSetFormula("NA 1", "'" & Trim$(tgSpf.sBPayName) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(ilFormulaNo) = "NA 1= '" & tgSpf.sGClient & "'"
                ilFormulaNo = ilFormulaNo + 1
                '10-17-05 chg from tgSpf.sGAddr to tgSpf.sBPayAddr
                If Not gSetFormula("NA 2", "'" & Trim$(tgSpf.sBPayAddr(0)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(ilFormulaNo) = "NA 2= '" & tgSpf.sGAddr(0) & "'"
                ilFormulaNo = ilFormulaNo + 1
                '10-17-05 chg from tgspf.sgaddr to tgspf.sbpayaddr
                If Not gSetFormula("NA 3", "'" & Trim$(tgSpf.sBPayAddr(1)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                'Report!crcReport.Formulas(ilFormulaNo) = "NA 3= '" & tgSpf.sGAddr(1) & "'"
                ilFormulaNo = ilFormulaNo + 1
                If Not gSetFormula("NA 4", "'" & Trim$(tgSpf.sBPayAddr(2)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                ilRet = mSetupSpacingForm()           '6-28-05 send crystal report the spacing before & after logos
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If
                '9-5-03 Send blanks to show in header to align to fit in windowed envelope
                'If tgSpf.sExport = "0" Or tgSpf.sExport = "N" Or tgSpf.sExport = "Y" Or tgSpf.sExport = "" Then
                '    ilBlanksBeforeLogo = 0
                'Else
                '    ilBlanksBeforeLogo = Val(tgSpf.sExport)
                'End If
                'If Trim$(tgSpf.sImport) = "0" Or Trim$(tgSpf.sImport) = "N" Or Trim$(tgSpf.sImport) = "Y" Or Trim$(tgSpf.sImport) = "" Then
                '    ilBlanksAfterLogo = 0
                'Else
                '    ilBlanksAfterLogo = Val(tgSpf.sImport)
                'End If
                'If Not gSetFormula("BlanksBeforeLogo", ilBlanksBeforeLogo) Then
                '    gCmcGen = -1
                '    Exit Function
                'End If
                'If Not gSetFormula("BlanksAfterLogo", ilBlanksAfterLogo) Then
                '    gCmcGen = -1
                '    Exit Function
                'End If

                '2-12-04 Send the default TERMS incase not defined with Advt or Agy
                If Not gSetFormula("DefaultTerms", "'" & sgDefaultTerms & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                '3-17-10 If user entered override terms on input screen, show that on all statements printed
                'obtain the selected override terms
                slInclude = ""
                If RptSel!cbcSet1.ListIndex > 0 Then
                    slInclude = RptSel!cbcSet1.List(RptSel!cbcSet1.ListIndex)
                End If
                If Not gSetFormula("OverrideTerms", "'" & slInclude & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                '11-22-06
                If ((Asc(tgSpf.sUsingFeatures4) And LOCKBOXBYVEHICLE) = LOCKBOXBYVEHICLE) Then  'using lock boxes by vehicle vs payee
                    If Not gSetFormula("LockBoxByVehicle", "'V'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("LockBoxByVehicle", "'NP'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                '5/16/08 tax not as gross option added
                If RptSel!ckcSelC10(0).Value = 1 Then

                    If Not gSetFormula("taxflag", "false") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("taxflag", "true") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                If RptSel!rbcSelCInclude(0).Value Then            'Detail
                    '4-27-12  test for vehicle name word wrap (for forms only)
                    If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
                        If Not gSetFormula("WordWrapVehicle", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("WordWrapVehicle", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                End If

                
                'Report!crcReport.Formulas(ilFormulaNo) = "NA 4= '" & tgSpf.sGAddr(2) & "'"
            '1-23-18 add option by Sales Commission on Collections
            ElseIf ilListIndex = COLL_CASH Or ilListIndex = COLL_SALESCOMM_COLL Then                 'default to current billing period, in the past is in error
                '8-28-19 use csi calendar control vs edit box
'                slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from PRF or RVF
                slStr = RptSel!CSI_CalFrom.Text                'Earliest date to retrieve from PRF or RVF
                llDate1 = gDateValue(slStr)
                If llDate1 = 0 Then
                    llDate1 = gDateValue("1/1/1970")
                End If
'                slStr = RptSel!edcSelCTo.Text               'Latest date to retrieve from PRF or RVF
                slStr = RptSel!CSI_CalTo.Text               'Latest date to retrieve from PRF or RVF
                llDate2 = gDateValue(slStr)
                If llDate2 = 0 Then                    'if end date not entered, use all
                    llDate2 = gDateValue("12/29/2069")
                End If
                If llDate1 > llDate2 Then


                    If RptSel!CSI_CalFrom.Text = "" Then
                         mReset
                         RptSel!CSI_CalFrom.SetFocus
                         Exit Function
                    Else
                         mReset
                         RptSel!CSI_CalTo.SetFocus
                         Exit Function
                    End If
                End If

                'Send input dates to report
'                slDateFrom = RptSel!edcSelCFrom.Text
'                slDateTo = RptSel!edcSelCTo.Text
                slDateFrom = RptSel!CSI_CalFrom.Text        '8-28-19 use csi cal control
                slDateTo = RptSel!CSI_CalTo.Text

                If Trim$(slDateFrom) = "" And Trim$(slDateTo) = "" Then             'no dates entered
                    slStr = "All Dates"
                ElseIf Trim$(slDateFrom) <> "" And Trim$(slDateTo) <> "" Then       'both dates entered
                    slStr = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
                    slStr = slStr & " - " & Format$(gDateValue(slDateTo), "m/d/yy")
                ElseIf Trim$(slDateFrom) = "" Then
                    slStr = "Thru " & Format$(gDateValue(slDateTo), "m/d/yy")
                Else
                    slStr = "From " & Format$(gDateValue(slDateFrom), "m/d/yy")
                End If
                
                If ilListIndex = COLL_SALESCOMM_COLL Then
                    slStr = "For Commission Dates: " & slStr
                End If
                If Not gSetFormula("InputDates", "'" & slStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                'If Not mTitleCashTrMercProm() Then               'send Crystal reporttitle
                 If Not mTitleCashRpts() Then
                    gCmcGen = -1
                End If
                ' 6-04-08 changed to show history/receivables  Dan M
                ilRet = mTitleTypeAndReceivables(True, True)    'parameters:  Ok to combine hard cost with non hard-cost; Include hard costs
                If ilRet <> 0 Then
                    gCmcGen = -1
                End If

                If RptSel!ckcSelC3(0).Value = vbChecked And RptSel!ckcSelC3(1).Value = vbChecked And RptSel!ckcSelC3(2).Value = vbChecked Then
                    slStr = ""
                ElseIf RptSel!ckcSelC3(0).Value = vbChecked Then            'include payments only
                    slStr = "Payments(PI) Only"
                    If RptSel!ckcSelC3(1).Value = vbChecked Then
                        slStr = slStr & ", Payments(PO) Only"
                    ElseIf RptSel!ckcSelC3(2).Value = vbChecked Then
                        slStr = slStr & ", Journal Entries(W) Only"
                    End If
                ElseIf RptSel!ckcSelC3(1).Value = vbChecked Then            'journal entries only
                    slStr = "Payments(PO) Only"
                    If RptSel!ckcSelC3(2).Value = vbChecked Then
                        slStr = slStr & ", Journal Entries(W) Only"
                    End If
                ElseIf RptSel!ckcSelC3(2).Value = vbChecked Then
                    slStr = "Journal Entries(W) Only"
                Else
                    slStr = "No Transactions Types selected"
                End If
                If Not gSetFormula("TranTypeHdr", "'" & slStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If ilListIndex = COLL_SALESCOMM_COLL Then
                    If RptSel!rbcSelC12(0).Value Then           'totals by contract
                        If Not gSetFormula("TotalsBy", "'C'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("TotalsBy", "'T'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                    
                    If tgSpf.sSubCompany = "Y" Then
                        If Not gSetFormula("SubCompany", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("SubCompany", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                Else                'Cash receipts
                    If Not mShowTransComments(RptSel!ckcSelC7) Then        'send formula to show/not show trans comments
                        gCmcGen = -1
                    End If
    
                    If RptSel!rbcSelC4(2).Value = True Then             '11-17-05 major sort by vehicle group
                        If Not gSetFormula("MajorSortByVehicleGroup", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                        illoop = RptSel!cbcSet1.ListIndex           'Determine vehicle group selected for heading
                        ilRet = gFindVehGroupInx(illoop, tgVehicleSets1())
    
                        If Not gSetFormula("GroupSelectedDesc", ilRet) Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    ElseIf RptSel!rbcSelC4(0).Value = True Then         'by date
                        If Not gSetFormula("MajorSortByVehicleGroup", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    'send no formula to option by salesperson, it has a different crystl report
                    End If
                    
                    '11-8-19 regardless of sort option, if NTR printed, ask to show the description
                    If (RptSel!rbcSelC6(0).Value) Or ((Not RptSel!rbcSelC6(0).Value) And (RptSel!ckcOption.Value = vbUnchecked)) Then   'air time only or not air time only and they do not want to see ntr item type
                        If Not gSetFormula("ShowNTRItem", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowNTRItem", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                End If

                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

            ElseIf ilListIndex = 7 Then 'agency and advertiser Credit Status
                slSelection = ""
                If igJobRptNo = 1 Then  'Agency
                    'If Not RptSel!ckcSel1(0).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{AGF_Agencies.agfCreditRestr} <>" & "'N'"
                    '    Else
                    '        slSelection = "{AGF_Agencies.agfCreditRestr} <>" & "'N'"
                    '    End If
                    'End If
                    'If Not RptSel!ckcSel1(1).Value = vbChecked Then
                   '     If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    '    Else
                    '        slSelection = "{@Credit Used} <>" & "0"
                    '    End If
                    'End If
                    'If RptSel!ckcADate.Value = vbChecked Then          'see if only overdue accounts required
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
                    If RptSel!ckcSel1(1).Value = vbChecked Then         'include zero balance clients

                    Else
                        slSelection = "(" & slSelection & " and ({@Credit Used} <> 0) )"
                    End If

                    'include no new orders
                    If RptSel!ckcADate.Value = vbChecked Then
                        slSelection = slSelection & " or ( {AGF_Agencies.agfCreditRestr} = 'P') "
                    End If

                    If RptSel!ckcSel1(0).Value = vbChecked Then         'include unrestricted
                        slSelection = slSelection & " or ( {AGF_Agencies.agfCreditRestr} = 'N') "
                    End If
                    
                    'TTP 9893
                    slSelection = slSelection & " And {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                Else    'Advertiser
                    'If Not RptSel!ckcSel1(0).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{ADF_Advertisers.adfCreditRestr} <>" & "'N'"
                    '    Else
                    '        slSelection = "{ADF_Advertisers.adfCreditRestr} <>" & "'N'"
                    '    End If
                    'End If
                    'If Not RptSel!ckcSel1(1).Value = vbChecked Then
                    '    If slSelection <> "" Then
                    '        slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    '    Else
                    '        slSelection = "{@Credit Used} <>" & "0"
                    '    End If
                    'End If
                    'If RptSel!ckcADate.Value = vbChecked Then          'see if only overdue accounts required
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
                    If RptSel!ckcSel1(1).Value = vbChecked Then         'include zero balance clients

                    Else
                        slSelection = "(" & slSelection & " and ({@Credit Used} <> 0) )"
                    End If

                    'include no new orders
                    If RptSel!ckcADate.Value = vbChecked Then
                        slSelection = slSelection & " or ( {ADF_Advertisers.adfCreditRestr} = 'P') "
                    End If

                    If RptSel!ckcSel1(0).Value = vbChecked Then         'include unrestricted
                        slSelection = slSelection & " or ( {ADF_Advertisers.adfCreditRestr} = 'N') "
                    End If

                    'TTP 9893
                    slSelection = slSelection & " And {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                End If

                '2-24-05 determine to show action comments
                If RptSel!ckcInclCommentsA.Value = vbChecked Then
                    'yes, include comments.  If no date entered, show all comments
                    'force date entered as the earliest date possible
'                    If RptSel!edcSelA.Text = "" Then
                    If RptSel!CSI_CalDateA.Text = "" Then       '8-29-19 csi calendar control vs edit box
                        slYear2 = "1970"
                        slMonth2 = "1"
                        slDay2 = "1"
                    Else
                        slDateFrom = RptSel!CSI_CalDateA.Text
                        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
                        gObtainYearMonthDayStr slDateFrom, True, slYear2, slMonth2, slDay2

                    End If
                Else                    'dont show any comments
                    slYear2 = "2070" 'JW 9/30/21 - TTP 10286: Advertiser and Agency Credit Status Report - comments appearing when "include comments" is unchecked"
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

                If RptSel!ckcDelinquentOnly.Value = vbChecked Then      'delinquents (overdue) only
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

            'D.S. 8/15/01
            ElseIf ilListIndex = COLL_CASHSUM Then
'                slDateFrom = RptSel!edcSelCFrom.Text
                slDateFrom = RptSel!CSI_CalFrom.Text            '8-29-19 csi cal control vs edit box
                slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
'                slDateTo = RptSel!edcSelCTo.Text
                slDateTo = RptSel!CSI_CalTo.Text
                slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
                If Not gSetFormula("InputDates", "'" & slDateFrom & "-" & slDateTo & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                '3-21-05 parameters:  Ok to combine hard cost with non hard-cost; Include hard costs
                ilRet = mTitleAirTimeNTRHdr(True, True, False)  'tell crystal if requesting air time, ntr or both
                If ilRet <> 0 Then
                    gCmcGen = -1
                End If

                'If Not mTitleCashTrMercProm() Then               'send Crystal reporttitle
                 If Not mTitleCashRpts() Then
                    gCmcGen = -1
                End If

                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

            ElseIf ilListIndex = COLL_MERCHANT Then
                slStr = ""
                illoop = Val(RptSel!edcSelCFrom1.Text)
                If illoop = 1 Then
                    slStr = "1st Quarter "
                ElseIf illoop = 2 Then
                    slStr = "2nd Quarter "
                ElseIf illoop = 3 Then
                    slStr = "3rd Quarter "
                Else
                    slStr = "4th Quarter "
                End If
                slName = RptSel!edcSelCFrom.Text
                If Not gSetFormula("QtrHeader", "'" & slStr & slName & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If RptSel!rbcSelCSelect(0).Value Then       'vehicle
                    If Not gSetFormula("SortBy", "'V'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("SortBy", "'A'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                If RptSel!rbcSelC4(0).Value Then       'vehicle
                    If Not gSetFormula("MrchOrPromo", "'M'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("MrchOrPromo", "'P'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                slStr = RptSel!edcSelCTo.Text      'user input Pct from & to
                gFormatStr slStr, FMTPERCENTSIGN, 2, slName         '4/2/99 chg from 3 to 2 dec places
                ilPos1 = gStrDecToInt(slStr, 2)
                slStr = RptSel!edcSelCTo1.Text      'user input Pct from & to
                gFormatStr slStr, FMTPERCENTSIGN, 2, slNameCode
                ilPos2 = gStrDecToInt(slStr, 2)
                If RptSel!rbcSelC4(0).Value Then    'merchandising
                    slCity = "Merchandising"
                Else
                    slCity = "Promotion"
                End If
                If ilPos1 = 0 And ilPos2 = 0 Then
                    'all percents
                    If Not gSetFormula("PctFromTo", "'" & "All " & slCity & " Percents'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If ilPos2 = 0 Then
                        'ilPos2 = 30000                 'assume highest allowed
                        slNameCode = "30.000%"
                    End If
                    If Not gSetFormula("PctFromTo", "'" & slCity & " Pcts " & slName & " - " & slNameCode & "'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                'gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                'slSelection = "({GRF Receivables Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'slSelection = slSelection & " And {RVR_Receivables_Rept.rvrGenTime} = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
'10/30/20  - TTP # 10008 - No Longer used in MchVehDt, MchVehSm, MchAdvDt, and MchAdvSm Crystal Reports
'                gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
'                If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGen = -1
'                    Exit Function
'                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = COLL_MERCHRECAP Then
                ilRet = mVerifyDate(RptSel!edcSelCFrom, llDate1, False)
                ilRet = mVerifyDate(RptSel!edcSelCFrom1, llDate2, False)
                If llDate1 = 0 And llDate2 = 0 Then             'no dates entered
                    slStr = "All Dates"
                ElseIf llDate1 <> 0 And llDate2 <> 0 Then       'both dates entered
                    slStr = Format$(llDate1, "m/d/yy")
                    slStr = slStr & " - " & Format$(llDate2, "m/d/yy")
                ElseIf llDate1 = 0 Then
                    slStr = "Thru " & slStr & Format$(llDate2, "m/d/yy")
                Else
                    slStr = "From " & Format$(llDate1, "m/d/yy")
                End If
                If Not gSetFormula("DateHdr", "'" & slStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
'10/30/20 - TTP # 10008 - No Longer used in MchRecap Crystal Report
'                gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
'                If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGen = -1
'                    Exit Function
'                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = COLL_POAPPLY Then
                '8-30-19 use csi calendar control vs edit box
                'Verify input dates for Entered Date selectivity
'                ilRet = mVerifyDate(RptSel!edcSelCFrom, llDate1, True)
                ilRet = mVerifyDate(RptSel!CSI_CalFrom, llDate1, True)
                If ilRet <> 0 Then
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    Exit Function
                End If
'                ilRet = mVerifyDate(RptSel!edcSelCFrom1, llDate2, True)
                ilRet = mVerifyDate(RptSel!CSI_CalTo, llDate2, True)
                If ilRet <> 0 Then
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    Exit Function
                End If
                If llDate1 <> 0 And llDate2 <> 0 Then
                    If llDate1 > llDate2 Then
                        mReset
                        RptSel!edcSelCFrom.SetFocus
                        Exit Function
                    End If
                End If

                'Verify input dates for Transaction Date selectivity
'                ilRet = mVerifyDate(RptSel!edcSelCTo, llDate1, True)
                 ilRet = mVerifyDate(RptSel!CSI_CalFrom2, llDate1, True)
               If ilRet <> 0 Then
                    mReset
                    RptSel!CSI_CalFrom2.SetFocus
                    Exit Function
                End If
'                ilRet = mVerifyDate(RptSel!edcSelCTo1, llDate2, True)
                ilRet = mVerifyDate(RptSel!CSI_CalTo2, llDate2, True)
                If ilRet <> 0 Then
                    mReset
                    RptSel!CSI_CalTo2.SetFocus
                    Exit Function
                End If
                If llDate1 <> 0 And llDate2 <> 0 Then
                    If llDate1 > llDate2 Then
                        mReset
                        RptSel!CSI_CalFrom2.SetFocus
                        Exit Function
                    End If
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

                '8-30-19 send the calendar control input date
'                slDateFrom = RptSel!edcSelCFrom.Text        'transaction dates
'                slDateTo = RptSel!edcSelCFrom1.Text
                slDateFrom = RptSel!CSI_CalFrom.Text        'transaction dates
                slDateTo = RptSel!CSI_CalTo.Text
                ilRet = mDateHdr(slDateFrom, slDateTo, "CreationDates")
'                slDateFrom = RptSel!edcSelCTo.Text        'creation date (date entered)
'                slDateTo = RptSel!edcSelCTo1.Text
                slDateFrom = RptSel!CSI_CalFrom2.Text        'creation date (date entered)
                slDateTo = RptSel!CSI_CalTo2.Text
                ilRet = mDateHdr(slDateFrom, slDateTo, "TransDates")

            End If

            If (ilListIndex = COLL_AGEPAYEE) Or (ilListIndex = COLL_AGESLSP) Or (ilListIndex = COLL_AGEVEHICLE) Or (ilListIndex = 4) Or (ilListIndex = 5) Or (ilListIndex = COLL_AGEOWNER) Or (ilListIndex = COLL_AGEPRODUCER) Or (ilListIndex = COLL_AGESS) Then
                If RptSel!ckcSelC11(0).Value = vbChecked Then       'extended ageing column
                    If mExtendedAgeing(slBaseDate) = -1 Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    'Set 5 previous end date values from base date
                    If tgSpf.sRRP = "C" Then    'Calendar
                        slDate = slBaseDate
                        slDate = gObtainStartCal(slDate)
                        For illoop = 2 To 6 Step 1
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                            'slDate = gObtainEndCal(slDate)
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                            slDate = gObtainStartCal(slDate)
                        Next illoop
                    ElseIf tgSpf.sRRP = "F" Then 'Corporate
                        slDate = slBaseDate
                        slDate = gObtainStartCorp(slDate, True)
                        For illoop = 2 To 6 Step 1
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                            'slDate = gObtainEndCorp(slDate)
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                            slDate = gObtainStartCorp(slDate, True)
                        Next illoop
                    Else    'Standard
                        slDate = slBaseDate
                        slDate = gObtainStartStd(slDate)
                        For illoop = 2 To 6 Step 1
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                            'slDate = gObtainEndStd(slDate)
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                            slDate = gObtainStartStd(slDate)
                        Next illoop
                    End If
                End If                  'ckcSelC11(0)
            End If
            'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
            If ilListIndex = COLL_AGEMONTH And RptSel!rbcOutput(3).Value = False Then
                If RptSel!ckcSelC6Add(0).Value = vbChecked Then         'including hardcost
                    slStr = "Including Cash Air Time, NTR & Hard Cost"
                Else
                    slStr = "Including Cash Air Time & NTR; Excluding Hard Cost"
                End If
                If Not gSetFormula("IncludeExclude", "'" & slStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            End If
    
        Case POSTLOGSJOB
            If ilListIndex = 0 Or ilListIndex = 1 Then                 'log posting or missing isci
                '8-20-19 change to use csi calendar control vs text box
'                If mVerifyDate(RptSel!edcSelCFrom, llDate1, True) Then
                If mVerifyDate(RptSel!CSI_CalFrom, llDate1, True) Then
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If
                If mVerifyDate(RptSel!CSI_CalTo, llDate1, True) Then
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If
            End If
            If ilListIndex = PL_LIVELOG Then            '12-8-05
                gCmcGen = mPLDatesFormula()     'get the PL dates and send to crystal reports
                If gCmcGen <> 0 Then
                    Exit Function
                End If

                'obtain current date and time filtering
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = PL_STATION_POSTING Then        '1-23-19
                '8-20-19 change to use csi calendar control vs text box
'                If mVerifyDate(RptSel!edcSelCFrom, llDate1, True) Then
                If mVerifyDate(RptSel!CSI_CalFrom, llDate1, True) Then
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCFrom1, llDate2, True) Then
                If mVerifyDate(RptSel!CSI_CalTo, llDate2, True) Then
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo, llDate1, True) Then
                If mVerifyDate(RptSel!CSI_CalFrom2, llDate1, True) Then
                    mReset
                    RptSel!CSI_CalFrom2.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo1, llDate2, True) Then
                If mVerifyDate(RptSel!CSI_CalTo2, llDate2, True) Then
                    mReset
                    RptSel!CSI_CalTo2.SetFocus
                    'gCmcGen = -1
                    Exit Function
                End If

'                ilRet = mInpDateFormula(RptSel!edcSelCFrom, RptSel!edcSelCFrom1, "InvDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "InvDates")
                If gCmcGen <> 0 Then
                    Exit Function
                End If
'                ilRet = mInpDateFormula(RptSel!edcSelCTo, RptSel!edcSelCTo1, "LogInDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom2, RptSel!CSI_CalTo2, "LogInDates")
                If gCmcGen <> 0 Then
                    Exit Function
                End If
                
                If RptSel!cbcSet1.ListIndex = 0 Then
                    slSortStr = "I"
                ElseIf RptSel!cbcSet1.ListIndex = 1 Then
                    slSortStr = "L"
                Else
                    slSortStr = "V"
                End If
                If Not gSetFormula("SortByMajor", "'" & slSortStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                If RptSel!cbcSet2.ListIndex = 0 Then
                    slSortStr = "I"
                ElseIf RptSel!cbcSet2.ListIndex = 1 Then
                    slSortStr = "L"
                Else
                    slSortStr = "V"
                End If
                If Not gSetFormula("SortByMinor", "'" & slSortStr & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                If RptSel!ckcSelC10(0).Value = vbChecked Then
                    If Not gSetFormula("SkipPageMajor", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("SkipPageMajor", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                If RptSel!ckcSelC10(1).Value = vbChecked Then
                    If Not gSetFormula("SkipPageMinor", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("SkipPageMinor", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                'obtain current date and time filtering
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            Else

                slExclude = ""
                slInclude = ""
                If ilListIndex = 0 Then         'log posting status
                    gIncludeExcludeCkc RptSel!ckcSelC3(0), slInclude, slExclude, "Billed"
                    gIncludeExcludeCkc RptSel!ckcSelC3(1), slInclude, slExclude, "Unbilled"

                    gIncludeExcludeCkc RptSel!ckcSelC3(3), slInclude, slExclude, "PSA/Promo"
                    gIncludeExcludeCkc RptSel!ckcSelC3(4), slInclude, slExclude, "Missed"
                    gIncludeExcludeCkc RptSel!ckcSelC3(5), slInclude, slExclude, "Cancelled"
                    gIncludeExcludeCkc RptSel!ckcSelC3(6), slInclude, slExclude, "Hidden"
                Else                'missing isci
                    gIncludeExcludeCkc RptSel!ckcSelC3(7), slInclude, slExclude, "+Fill"
                    gIncludeExcludeCkc RptSel!ckcSelC3(8), slInclude, slExclude, "-Fill"
                End If

                'only show inclusion/exclusion of cntr/feed spots if not included
                If tgSpf.sSystemType = "R" Then         'radio system are only ones that can have feed spots
                    If Not RptSel!ckcSelC10(0).Value = vbChecked Then
                        gIncludeExcludeCkc RptSel!ckcSelC10(0), slInclude, slExclude, "Contract spots"
                    End If
                    If Not RptSel!ckcSelC10(1).Value = vbChecked Then
                        gIncludeExcludeCkc RptSel!ckcSelC10(1), slInclude, slExclude, "Feed spots"
                    End If
                End If
                If Len(slInclude) > 0 Then
                    If Not gSetFormula("Included", "'" & slInclude & "'") Then
                        gCmcGen = -1
                    End If
                Else
                    If Not gSetFormula("Included", "'" & " " & "'") Then
                        gCmcGen = -1
                    End If
                End If
                If Len(slExclude) > 0 Then
                    If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                        gCmcGen = -1
                    End If
                Else
                    If Not gSetFormula("Excluded", "'" & " " & "'") Then
                        gCmcGen = -1
                    End If
                End If

                gCmcGen = mPLDatesFormula()     'get the PL dates and send to crystal reports
                If gCmcGen <> 0 Then
                    Exit Function
                End If

                '4-28-11 option to show summary day is complete flags
                slStr = ""
                'determine detail or summary (detail = spots, summary is day is incomplete flags)
                If RptSel!rbcSelC8(1).Value = True Then         'summary
                    If RptSel!ckcTrans.Value = vbChecked Then   'show discreps only on day is incomplete
                        slStr = "DS"        'day incomplete flags only
                    Else
                        slStr = "AS"        'all day is complete flags
                    End If
                End If
                If Not gSetFormula("SummaryOnly", "'" & slStr & "'") Then
                    gCmcGen = -1
                End If

                'Dates are checked within mGenReport
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear

                slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            End If
            'If ilIndex = 0 Then
            '    If Not gSetFormula("ReportHeader", "'" & "Log Posting Status" & "'") Then
            '        Exit Function
            '    End If
            'Else
            '    If Not gSetFormula("ReportHeader", "'" & "Missing ISCI Codes" & "'") Then
            '        Exit Function
            '    End If
            'End If

        Case COPYJOB
            If Not gSetSelection(slSelection) Then
                gCmcGen = -1
                Exit Function
            End If
            If ilListIndex = 11 Then                                'playlist by isci
                If RptSel!rbcSelCSelect(1).Value = True Then        'check for playlist
                    'If RptSel!ckcSelC5(0).Value = vbChecked Then    'include split copy?
                    If RptSel!rbcSelC8(1).Value = True Then
                        If Not gSetFormula("ShowSplitCopy", "'Y'") Then
                            gCmcGen = -15
                            Exit Function
                        End If
                    ElseIf RptSel!rbcSelC8(2).Value Then            'split copy only
                        If Not gSetFormula("ShowSplitCopy", "'O'") Then
                            gCmcGen = -15
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("ShowSplitCopy", "'N'") Then     'no split copy, just generic
                            gCmcGen = -15
                            Exit Function
                        End If
                    End If

                    If RptSel!ckcTrans.Value = vbChecked Then       'show rotation dates
                        If Not gSetFormula("ShowRotationDates", "'Y'") Then     'show the rotation dates from inventory
                            gCmcGen = -15
                            Exit Function
                        End If
                    Else
                     If Not gSetFormula("ShowRotationDates", "'N'") Then     'hide the rotation dates from inventory
                            gCmcGen = -15
                            Exit Function
                        End If
                    End If
                    
                    slHeader = RptSel!edcSelCTo1.Text
                    ilPos1 = InStr(slHeader, "'")
                    If ilPos1 = 0 Then      'no apostrophe in text
                        If Not gSetFormula("Descrip", "'" & slHeader & "'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                         slNameCode = Mid$(slHeader, 1, ilPos1 - 1)   'get string upto the apostrophe"
                         'slName = "+" & """'""" & "+"
                         slCode = Mid$(slHeader, ilPos1 + 1)
                         slOr = slNameCode & slCode
                         If Not gSetFormula("Descrip", "'" & slOr & "'") Then
                            gCmcGen = -15
                            Exit Function
                        End If
                    End If
                    'If Not gSetFormula("AsOfT", Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))) Then
                    '    gCmcGen = -1
                    '    Exit Function
                    'End If
                End If
             End If
             
            ilRet = mCopyJob(ilListIndex, slLogUserCode)
            If ilRet = -1 Then
                gCmcGen = -1
                Exit Function
            ElseIf ilRet = 0 Then
                gCmcGen = 0
                Exit Function
            ElseIf ilRet = 2 Then
                gCmcGen = 2
                Exit Function
            End If
        Case INVOICESJOB
            slSortStr = "IDGSVEONC"        'I =inv,D = advt,G = agy,S = slsp,V = bill veh, E = air veh, N = NTR Item billing, C = Sales Source
            If ilListIndex = INV_REGISTER Or ilListIndex = 2 Then      'inv reg or billing distribution.
                slDateFrom = ""
                slDateTo = ""
                'Date selection passed by formula
'                slDate = RptSel!edcSelCFrom.Text   'Start date
                slDate = RptSel!CSI_CalFrom.Text   '8-15-19 Start date
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                
                'TTP 10118 -Billing Distribution Export to CSV
                If RptSel!rbcOutput(3).Value = False Then
                    slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                    'TTP 10395 - Invoice Register report: data from a previous, terminated run of the report was included due to report Gen Date/Time not being used
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If
    '                slDate = RptSel!edcSelCFrom.Text   'Start date
                    slDate = RptSel!CSI_CalFrom.Text   '8-15-19 Start date
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("StartDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
    '                slDate = RptSel!edcSelCTo.Text   'End date
                    slDate = RptSel!CSI_CalTo.Text    '8-15-19 End date
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("EndDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                If ilListIndex = INV_REGISTER Then                 'invoice register, send Crystal which sort
                    If RptSel!rbcSelCSelect(9).Value Then       'sort by sales origin
                        'skip to new page each vehicle (if applicable by air/bill vehicle)
                        If RptSel!ckcTrans.Value = vbChecked Then        'skip page each vehicle
                            If Not gSetFormula("SkipPage", "'Y'") Then        'page skip
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("SkipPage", "'N'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                        If RptSel!rbcSelC8(0).Value Then        'no vehicle major totals
                            If Not gSetFormula("SortBy", "'O'") Then        'invoice
                                gCmcGen = -1
                                Exit Function
                            End If
                        ElseIf RptSel!rbcSelC8(1).Value Then    'bill vehicle major totals, then sales origin
                            If Not gSetFormula("SortBy", "'B'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else                                    'air vehicle major totals , then sales origin
                            If Not gSetFormula("SortBy", "'A'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    Else
                        For illoop = 0 To 8 Step 1                  'send invoice register the sort option code
                            If RptSel!rbcSelCSelect(illoop).Value Then
                                slStr = Mid$(slSortStr, illoop + 1, 1)
                                If Trim$(slStr) = "E" Then      'sort by airing vehicle, see if user requested vehicle groups
                                    If RptSel!cbcSet1.ListIndex > 0 Then
                                        slStr = "P"         'let Crystal know its vehicle group & airing vehicle
                                    End If
                                ElseIf Trim$(slStr) = "D" Then      'sort by advt, see if subsort by vehicle gorup
                                    If RptSel!cbcSet1.ListIndex > 0 Then
                                        slStr = "T"
                                    End If
                                End If
                                If Not gSetFormula("SortBy", "'" & slStr & "'") Then        'invoice
                                    gCmcGen = -1
                                    Exit Function
                                End If
                                Exit For
                            End If
                        Next illoop
                        If RptSel!rbcSelCSelect(8).Value Then           'sales source option
                            If RptSel!rbcSelCInclude(1).Value Then    'summary for sales sources needs the landscape version, tell crystal to hide/show show total levels
                                slStr = "Y"
                            Else
                                slStr = "N"
                            End If
                            If Not gSetFormula("SumForSSOption", "'" & Trim$(slStr) & "'") Then        'invoice
                                gCmcGen = -1
                                Exit Function
                            End If
                        ElseIf RptSel!rbcSelCInclude(0).Value And Not RptSel!rbcSelCSelect(0).Value Then         'if detail version on any sort except Invoice sort, tell crystal to hide/show total levels
                            If Not gSetFormula("SumForSSOption", "'N'") Then        'invoice
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    End If
                    
                    '2-2-12 if invoice sort and not summary version, or any other sort that is in detail, allow comments to be an option for AN transactions
                    If (RptSel!rbcSelCSelect(0).Value = True And RptSel!rbcSelCInclude(2).Value = False) Or (RptSel!rbcSelCSelect(0).Value = False And RptSel!rbcSelCInclude(0).Value = True) Then
                        If RptSel!ckcOption.Value = vbChecked Then        'skip page each vehicle
                            If Not gSetFormula("ShowANComments", "'Y'") Then        'page skip
                                gCmcGen = -1
                                Exit Function
                            End If
                        Else
                            If Not gSetFormula("ShowANComments", "'N'") Then
                                gCmcGen = -1
                                Exit Function
                            End If
                        End If
                    End If
                    

                    '2-2-03 Show selectivity for Air Time, NTR or both in report heading
                    '3-21-05 parameters:  Ok to combine hard cost with non hard-cost; Include hard costs (true or false)
                    ilInclHardCost = False
                    If RptSel!ckcSelC7.Value = vbChecked Then
                        ilInclHardCost = True
                    End If

                    ilRet = mTitleAirTimeNTRHdr(False, ilInclHardCost, True)
                    If ilRet <> 0 Then
                        gCmcGen = -1
                    End If
                    '3-19-03 show selectivity for tran types (IN/AN/HI)
                    ilRet = mTitleRecHistBoth()
                    If ilRet <> 0 Then
                        gCmcGen = -1
                    End If

                    '7-15-08 Politicals/non-Politicals
                    ilRet = mTitlePoliticals(RptSel!ckcSelC11(0), RptSel!ckcSelC11(1))
                    If ilRet <> 0 Then
                        gCmcGen = -1
                    End If
                'TTP 10118 -Billing Distribution Export to CSV
                ElseIf ilListIndex = INV_DISTRIBUTE And RptSel!rbcOutput(3).Value = False Then             '1-22-15; Note:  formula send to Summary version but is not used.  Retain formula in crystal report
                    'skip to new page each vehicle option
                    If RptSel!ckcTrans.Value = vbChecked Then
                        If Not gSetFormula("PageSkipVehicle", "'Y'") Then        'skip new page each vehicle
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("PageSkipVehicle", "'N'") Then        'do not skip new page each vehicle
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
    
                    If Not gSetSelection(slSelection) Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
'*****************
'                '1-14-12 show AN transaction comments
'                If RptSel!ckcOption.Value = vbChecked Then       'show comments for AN transactions
'                    If Not gSetFormula("ShowANComments", "'Y'") Then
'                        gCmcGen = -1
'                        Exit Function
'                    Else
'                        If Not gSetFormula("ShowANComments", "'N'") Then
'                            gCmcGen = -1
'                            Exit Function
'                        End If
'                    End If
'                End If
                
            ElseIf ilListIndex = INV_VIEWEXPORT Then
                If (Not igUsingCrystal) Then
                    If RptSel!rbcOutput(0).Value Then
                        ilPreview = True
                    ElseIf RptSel!rbcOutput(1).Value Then
                        ilPreview = False
                    End If
                    'gDumpFileRpt 1, ilPreview, "DumpFile.Lst", 0, "View Invoice Export"
                    gDumpFileRpt 1, 0, "View Invoice Export"

                    'gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = INV_CREDITMEMO Then        '10-8-03
                '8-16-19 change using csi_calendar control vs edit box
'                If mVerifyDate(RptSel!edcSelCFrom, llDate1, False) Then
                If mVerifyDate(RptSel!CSI_CalFrom, llDate1, False) Then
                    gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCFrom1, llDate2, False) Then
                If mVerifyDate(RptSel!CSI_CalTo, llDate2, False) Then
                    gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo, llDate1, False) Then
                If mVerifyDate(RptSel!CSI_CalFrom2, llDate1, False) Then
                    gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo1, llDate2, False) Then
                If mVerifyDate(RptSel!CSI_CalTo2, llDate2, False) Then
                    gCmcGen = -1
                    Exit Function
                End If

'                ilRet = mInpDateFormula(RptSel!edcSelCFrom, RptSel!edcSelCFrom1, "InvDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "InvDates")
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If

'                ilRet = mInpDateFormula(RptSel!edcSelCTo, RptSel!edcSelCTo1, "CreationDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom2, RptSel!CSI_CalTo2, "CreationDates")
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If

                ilRet = mSetupSpacingForm()       '6-28-05 send crystal report the spacing before & after logos
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                'send to report which vehicle to show (billing or airing)
                If RptSel!rbcSelC4(0).Value Then            'show billing
                    If Not gSetFormula("UseBillOrAirVehicle", "'B'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else                                        'show airing
                    If Not gSetFormula("UseBillOrAirVehicle", "'A'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                'show detail or summarize by vehicle
                If RptSel!rbcSelC6(0).Value Then            'show detail
                    If Not gSetFormula("TotalsBy", "'D'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else                                        'summarize
                    If Not gSetFormula("TotalsBy", "'S'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If


                '9-5-03 Send blanks to show in header to align to fit in windowed envelope
                'If tgSpf.sExport = "0" Or tgSpf.sExport = "N" Or tgSpf.sExport = "Y" Or tgSpf.sExport = "" Then
                '    ilBlanksBeforeLogo = 0
                'Else
                '    ilBlanksBeforeLogo = Val(tgSpf.sExport)
                'End If
                'If Trim$(tgSpf.sImport) = "0" Or Trim$(tgSpf.sImport) = "N" Or Trim$(tgSpf.sImport) = "Y" Or Trim$(tgSpf.sImport) = "" Then
                '    ilBlanksAfterLogo = 0
                'Else
                '    ilBlanksAfterLogo = Val(tgSpf.sImport)
                'End If
                'If Not gSetFormula("BlanksBeforeLogo", ilBlanksBeforeLogo) Then
                '    gCmcGen = -1
                '    Exit Function
                'End If
                'If Not gSetFormula("BlanksAfterLogo", ilBlanksAfterLogo) Then
                '    gCmcGen = -1
                '    Exit Function
                'End If

                '2-12-04 Send the default TERMS incase not defined with Advt or Agy
                If Not gSetFormula("DefaultTerms", "'" & sgDefaultTerms & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If Not gSetFormula("NA 1", "'" & Trim$(tgSpf.sGClient) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                If Not gSetFormula("NA 2", "'" & Trim$(tgSpf.sBPayAddr(0)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                If Not gSetFormula("NA 3", "'" & Trim$(tgSpf.sBPayAddr(1)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                If Not gSetFormula("NA 4", "'" & Trim$(tgSpf.sBPayAddr(2)) & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                '3-7-07
                If ((Asc(tgSpf.sUsingFeatures4) And LOCKBOXBYVEHICLE) = LOCKBOXBYVEHICLE) Then  'using lock boxes by vehicle vs payee
                    If Not gSetFormula("LockBoxByVehicle", "'V'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("LockBoxByVehicle", "'P'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                '4-26-12  test for vehicle name word wrap (for forms only)
                If ((Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE) Then
                    If Not gSetFormula("WordWrapVehicle", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("WordWrapVehicle", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

            ElseIf ilListIndex = INV_SUMMARY Then        '6-28-05
                 'validity check the input dates, allow month text jan, feb, etc or month # (1-12)

                slStr = RptSel!edcSelCTo.Text             'month in text form (jan..dec)
                gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
                If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                    ilSaveMonth = Val(slStr)
                    ilRet = gVerifyInt(slStr, 1, 12)
                    If ilRet = -1 Then
                        mReset
                        RptSel!edcSelCTo.SetFocus                 'invalid # periods
                        Exit Function
                    End If
                End If

                slStr = RptSel!edcSelCTo1.Text
                ilRet = gVerifyYear(slStr)
                'If Val(slStr) < 1990 Or Val(slStr) > 2020 Then
                If ilRet = 0 Then
                    mReset
                    RptSel!edcSelCTo1.SetFocus                 'invalid year
                    gCmcGen = False
                    Exit Function
                End If

                ilRet = mSetupSpacingForm()       '6-28-05 send crystal report the spacing before & after logos
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If

                If Not gSetFormula("DefaultTerms", "'" & sgDefaultTerms & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                If RptSel!rbcSelCInclude(0).Value Then       'Detail
                    If Not gSetFormula("DetailSummary", "'D'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    
                    '1-12-11 determine whether to hide or show the transaction net amounts
                    If RptSel!ckcSelC10(0).Value = vbChecked Then       'hide net amounts
                        If Not gSetFormula("HideTransNet", "'Y'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    Else
                        If Not gSetFormula("HideTransNet", "'N'") Then
                            gCmcGen = -1
                            Exit Function
                        End If
                    End If
                Else
                    If Not gSetFormula("DetailSummary", "'S'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    If Not gSetFormula("HideTransNet", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If

                If RptSel!ckcSelC7.Value = vbChecked Then       'Include ANs
                    If Not gSetFormula("InclAdjustments", "'Y'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("InclAdjustments", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                '5-23-13 show lock box by vehicle based on site
                If ((Asc(tgSpf.sUsingFeatures4) And LOCKBOXBYVEHICLE) = LOCKBOXBYVEHICLE) Then  'using lock boxes by vehicle vs payee
                    If Not gSetFormula("LockBoxByVehicle", "'V'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("LockBoxByVehicle", "'NP'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = INV_TAXREGISTER Then           '1-30-07
                 If mVerifyDate(RptSel!CSI_CalFrom, llDate1, True) Then        '8-15-19
'                If mVerifyDate(RptSel!edcSelCFrom, llDate1, True) Then
                    gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo, llDate2, True) Then
                 If mVerifyDate(RptSel!CSI_CalTo, llDate2, True) Then        '8-15-19
                    gCmcGen = -1
                    Exit Function
                End If
'                ilRet = mInpDateFormula(RptSel!edcSelCFrom, RptSel!edcSelCTo, "InvDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "InvDates")
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If
                slInclude = ""
                slExclude = ""
                gIncludeExcludeCkc RptSel!ckcSelC3(0), slInclude, slExclude, "Invoices"
                gIncludeExcludeCkc RptSel!ckcSelC3(1), slInclude, slExclude, "Inv Adjustments"
                gIncludeExcludeCkc RptSel!ckcSelC3(2), slInclude, slExclude, "Payments"
                gIncludeExcludeCkc RptSel!ckcSelC3(3), slInclude, slExclude, "Journal Entries"
                If Not gSetFormula("Included", "'" & slInclude & "'") Then
                    gCmcGen = -1
                End If

                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "({RVR_Receivables_Rept.rvrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({RVR_Receivables_Rept.rvrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = INV_RECONCILE Then        '11-30-07
'                If mVerifyDate(RptSel!edcSelCFrom, llDate2, False) Then
                If mVerifyDate(RptSel!CSI_CalFrom, llDate2, False) Then
                    gCmcGen = -1
                    Exit Function
                End If
'                If mVerifyDate(RptSel!edcSelCTo, llDate1, False) Then
                If mVerifyDate(RptSel!CSI_CalTo, llDate1, False) Then
                    gCmcGen = -1
                    Exit Function
                End If

'                ilRet = mInpDateFormula(RptSel!edcSelCFrom, RptSel!edcSelCTo, "ActiveDates")
                ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "ActiveDates")
                If ilRet <> 0 Then
                    gCmcGen = -1
                    Exit Function
                End If
                
                gUnpackDateLong tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), llDate1
                slBaseDate = Format$(llDate1, "m/d/yy")
                If Not gSetFormula("LastBillDate", "'" & slBaseDate & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If


                If RptSel!ckcSelC5(0).Value = vbChecked Then        '10-14-16 chged control (was ckcselc5(0); discrep only
                    If Not gSetFormula("DiscrepText", "'Discrepancy Only'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("DiscrepText", "' '") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                '10-14-16 gross or net option
                If RptSel!rbcSelC4(0).Value = True Then        'Gross
                    If Not gSetFormula("GrossOrNet", "'G'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("GrossOrNet", "'N'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                
                'obtain current date and time filtering
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            ElseIf ilListIndex = INV_UNPOSTED_STATIONS Then             '8-12-15
            
                'validity check the input dates, allow month text jan, feb, etc or month # (1-12)

                slStr = RptSel!edcSelCFrom.Text             'month in text form (jan..dec)
                gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
                If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                    ilSaveMonth = Val(slStr)
                    ilRet = gVerifyInt(slStr, 1, 12)
                    If ilRet = -1 Then
                        mReset
                        RptSel!edcSelCFrom.SetFocus                 'invalid # periods
                        Exit Function
                    End If
                End If

                slStr = RptSel!edcSelCTo.Text
                ilRet = gVerifyYear(slStr)
                If ilRet = 0 Then
                    mReset
                    RptSel!edcSelCTo.SetFocus                 'invalid year
                    gCmcGen = False
                    Exit Function
                End If
            
                slStr = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(RptSel!edcSelCTo.Text)
                slDateFrom = gObtainStartStd(slStr)
                slDateTo = gObtainEndStd(slStr)
                slDateFrom = slDateFrom & " - " & slDateTo
                If Not gSetFormula("DatesRequested", "'" & slDateFrom & "'") Then
                    gCmcGen = -1
                    Exit Function
                End If

                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If
            End If
       Case CHFCONVMENU
            slSelection = ""
            If Not RptSel!ckcAll.Value = vbChecked Then
                For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                    If RptSel!lbcSelection(0).Selected(illoop) Then
                        slName = RptSel!lbcSelection(0).List(illoop)
                        ilPos1 = InStr(slName, " on ")
                        ilPos2 = InStr(slName, " at ")
                        slDate = Mid$(slName, ilPos1 + 4, ilPos2 - ilPos1 - 4)
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slTime = Mid$(slName, ilPos2 + 4)
                        slTime = Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                        If slSelection <> "" Then
                            slSelection = slSelection & " Or " & "({ICF_Import_Contracts.icfDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")" & " And " & "{ICF_Import_Contracts.icfTime} =" & slTime & ")"
                        Else
                            slSelection = "({ICF_Import_Contracts.icfDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")" & " And " & "{ICF_Import_Contracts.icfTime} =" & slTime & ")"
                        End If
                    End If
                Next illoop
            End If
            If Not gSetSelection(slSelection) Then
                gCmcGen = -1
                Exit Function
            End If
        Case GENERICBUTTON


        Case CMMLCHG
            If (Not igUsingCrystal) Then
                If RptSel!rbcOutput(0).Value Then
                    ilPreview = True
                ElseIf RptSel!rbcOutput(1).Value Then
                    ilPreview = False
                End If

                gCmcGen = 2         'successful return
                Exit Function
            Else
                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGen = -1
                    Exit Function
                End If

            End If
        Case EXPORTAFFSPOTS
            If ilListIndex = 0 Then
               If Not gSetFormula("DumpRptName", "'Export Affiliate Spots'") Then
                   gCmcGen = -1
                   Exit Function
               End If
            Else         'NY Error Log
                If Not gSetFormula("DumpRptName", "'Export Affiliate Spots'") Then
                    gCmcGen = -1
                    Exit Function
                End If
                'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", 2, "New York Error Log"
            End If

            'obtain current date and time for record retrieval
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                gCmcGen = -1
                Exit Function
            End If
            If Not gSetSelection(slSelection) Then
                gCmcGen = -1
                Exit Function
            End If
        Case BULKCOPY
            slSelection = ""
            'If (Not igUsingCrystal) And (ilListIndex = 0) Then
            '    If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If
            '    gListFileRpt 0, ilPreview, "ListFile.Lst", ilListIndex, "Copy Bulk Feed"

            '    gCmcGen = 2         'successful return
            '    Exit Function
            'End If
            'If (Not igUsingCrystal) And (ilListIndex = 1) Then
            '    If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If

            '    gListFileRpt 0, ilPreview, "ListFile.Lst", ilListIndex, "Copy Bulk Feed Cross Reference"

            '    gCmcGen = 2         'successful return
            '    Exit Function
            'End If
            If ilListIndex = 0 Or ilListIndex = 1 Then

                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))


            ElseIf (ilListIndex = 2) Or (ilListIndex = 3) Or (ilListIndex = 4) Or (ilListIndex = 5) Then
                slSelection = "({CIF_Copy_Inventory.cifPurged} <> 'H')"
                If Not RptSel!ckcSelC3(0).Value = vbChecked Then
                    slSelection = slSelection & " And ({CIF_Copy_Inventory.cifPurged} <> 'A')"
                End If
                If Not RptSel!ckcSelC3(1).Value = vbChecked Then
                    slSelection = slSelection & " And ({CIF_Copy_Inventory.cifPurged} <> 'P')"
                End If
            End If
            If ilListIndex = 2 Then 'Affiliate BF by Cart #
                slStr = RptSel!edcSelCFrom.Text
                If (slStr <> "") Then
                    If Not gSetFormula("Low Cart", "'" & slStr & "'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    If slSelection = "" Then
                        slSelection = "({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) >= '" & slStr & "'"
                    Else
                        slSelection = slSelection & " And ({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) >= '" & slStr & "'"
                    End If
                End If
                slStr = RptSel!edcSelCTo.Text
                If (slStr <> "") Then
                    If Not gSetFormula("Hi Cart", "'" & slStr & "'") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                    If slSelection = "" Then
                        slSelection = "({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) <= '" & slStr & "'"
                    Else
                        slSelection = slSelection & " And ({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) <= '" & slStr & "'"
                    End If
                End If
            End If
            If ilListIndex = 3 Then 'Affiliate BF by Vehicle
                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If gValidDate(slDate) Then
                        slDateFrom = slDate
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If slSelection = "" Then
                            slSelection = "{CIF_Copy_Inventory.cifRotEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                        Else
                            slSelection = slSelection & " And {CIF_Copy_Inventory.cifRotEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                        End If
                    Else
                        mReset
                        RptSel!edcSelCFrom.SetFocus
                        Exit Function
                    End If
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Earliest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                If Not RptSel!ckcAll.Value = vbChecked Then

                    If slSelection = "" Then
                        slOr = ""
                    Else
                        slOr = " And ("
                    End If
                    For illoop = 0 To RptSel!lbcSelection(3).ListCount - 1 Step 1
                        If RptSel!lbcSelection(3).Selected(illoop) Then
                            slNameCode = tgVehicle(illoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            slSelection = slSelection & slOr & "{CYF_Copy_Feed_Dates.cyfvefCode} = " & Trim$(slCode)
                            slOr = " Or "
                        End If
                    Next illoop
                    slSelection = slSelection & ")"
                End If
            End If
            If ilListIndex = 4 Then 'Affiliate BF by Feed Date
                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If gValidDate(slDate) Then
                        slDateFrom = slDate
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If slSelection = "" Then
                            slSelection = "{CYF_Copy_Feed_Dates.cyfFeedDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                        Else
                            slSelection = slSelection & " And {CYF_Copy_Feed_Dates.cyfFeedDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                        End If
                    Else
                        mReset
                        RptSel!edcSelCFrom.SetFocus
                        Exit Function
                    End If
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Earliest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If gValidDate(slDate) Then
                        slDateTo = slDate
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        If slSelection = "" Then
                            slSelection = "{CYF_Copy_Feed_Dates.cyfFeedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                        Else
                            slSelection = slSelection & " And {CYF_Copy_Feed_Dates.cyfFeedDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        End If
                    Else
                        mReset
                        Beep
                        RptSel!edcSelCTo.SetFocus
                        Exit Function
                    End If
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Latest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGen = -1
                        Exit Function
                    End If
                End If
            End If
            If ilListIndex = 5 Then 'Affiliate BF by Advertiser
                If Not RptSel!ckcAll.Value = vbChecked Then

                    If slSelection = "" Then
                        slOr = ""
                    Else
                        slOr = " And ("
                    End If
                    For illoop = 0 To RptSel!lbcSelection(5).ListCount - 1 Step 1
                        If RptSel!lbcSelection(5).Selected(illoop) Then
                            slNameCode = tgAdvertiser(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            slSelection = slSelection & slOr & "{CIF_Copy_Inventory.cifAdfCode} = " & Trim$(slCode)
                            slOr = " Or "
                        End If
                    Next illoop
                    slSelection = slSelection & ")"
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
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:4/18/96       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Duplicated from gCmcGen due       *
'*         to procedure too large                      *
'       8-27-01 Fix date selectivity to Blackout reports
'*******************************************************
Function gCmcGenMore(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  blSortOrder                                                                           *
'******************************************************************************************


'   ilRet = gCmcGenMore(ilListIndex)
'
'   ilRet (O)-  -1=Crystal failure of gSetformula
'               0 = Invalid input data, stay in
'               1 = Successful Crystal
'               2 = Successful Bridge
'
    Dim illoop As Integer
    Dim slSelection As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slOr As String
    Dim slDate As String
    Dim slDateFrom As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slTime As String
    Dim ilPreview As Integer
    Dim slBaseDate As String
    Dim ilMnfCode As Integer
    Dim slNameYear As String
    Dim slBudgetName As String
    Dim slMoreBudgetNames As String
    'ReDim lmStartDates(1 To 13) As Long
    'ReDim lmEndDates(1 To 13) As Long
    Dim slStartTime As String
    Dim slEndTime As String
    Dim slSort As String * 1
    Dim slInclude As String     'Date: 4/17/2020 A-Active, D-Dormant, B-Both
    Dim slSortBy As String      'Date: 4/17/2020 sort by "Sign on Name" or "City"
    
    gCmcGenMore = 0
    Select Case igRptCallType

        Case VEHICLESLIST
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Or ilListIndex = 1 Or ilListIndex = 5 Then
            'ElseIf ilListIndex = 1 Then
                slSelection = ""
                If Not RptSel!ckcAll.Value = vbChecked Then
                    For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                        If RptSel!lbcSelection(0).Selected(illoop) Then
                            slNameCode = tgVehicle(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get vehicle code
                            If slSelection <> "" Then
                                slSelection = slSelection & " Or " & "{VEF_Vehicles.VEFCode} =" & Trim$(slCode)
                            Else
                                slSelection = "{VEF_Vehicles.VEFCode} =" & Trim$(slCode)
                            End If
                        End If
                    Next illoop
                Else
                End If
                If slSelection = "" Then
                    slSelection = "({VEF_Vehicles.VEFType} <> 'P')"
                Else
                    slSelection = "(" & slSelection & ") and ({VEF_Vehicles.VEFType} <> 'P')"
                End If

                '4-6-04 include dormant vehicles
                If RptSel!ckcSelC3(0).Value = vbUnchecked Then  'exclude the dormant vehicles
                    slSelection = slSelection & " and {VEF_Vehicles.VEFState} <> 'D'"
                End If
                'slSelection = slSelection & ")"
                slDate = RptSel!edcCheck.Text
                If slDate <> "" Then
                    If gValidDate(slDate) Then
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                        slDate = Format$(gDateValue(slDate), "m/d/yy")    'make sure year included
                        slSelection = slSelection & " and ({PIF_Participant_Info.pifStartDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        slSelection = slSelection & " and {PIF_Participant_Info.pifEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        slSelection = slSelection & " or {PIF_Participant_Info.pifStartDate} >= Date(" & slYear & "," & slMonth & "," & slDay & "))"
                        If Not gSetFormula("EffectiveDate", "'Effective Date: " & slDate & "'") Then
                            gCmcGenMore = -1
                            Exit Function
                        End If

                    Else
                        mReset
                        RptSel!edcCheck.SetFocus
                        Exit Function
                    End If
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            End If
            ' sort by type or owner
            If ilListIndex = 5 Then 'vehicle participant only
                If RptSel!rbcSelC4(0).Value = True Then
                    If Not gSetFormula("sortByTypeOrOwner", "{@type}") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("sortByTypeOrOwner", "{@OwnerName}") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If
            End If
            If ilListIndex = LIST_STDPKG Then
                'obtain current date and time for report headers
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            End If
            
        Case ADVERTISERSLIST
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Then
                If RptSel!rbcType(1).Value = True Then     'summary
                    slSelection = "{@First Letter} >=" & "'" & Trim$(RptSel!cbcFrom.Text) & "'"
                    slSelection = slSelection & " And " & "{@First Letter} <=" & "'" & Trim$(RptSel!cbcTo.Text) & "'"
                    If RptSel!ckcRepInv(0).Value <> vbChecked Or RptSel!ckcRepInv(1).Value <> vbChecked Then
                        If RptSel!ckcRepInv(0).Value = vbChecked Then           'external
                            slSelection = slSelection & " and ({ADF_Advertisers.adfRepInvGen} = " & "'" & "E" & "')"
                        Else
                            If RptSel!ckcRepInv(1).Value Then           'internal
                                slSelection = slSelection & " and ({ADF_Advertisers.adfRepInvGen} <>" & "'" & "E" & "')"
                            End If
                        End If
                    End If
'                    slDate = RptSel!edcAsOfDate.Text
                    slDate = RptSel!CSI_CalAsOfDate.Text        '8-16-19 use csi calendar control
                    If slDate <> "" Then
    
                        If gValidDate(slDate) Then
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            slSelection = slSelection & " and {ADF_Advertisers.adfDateEntrd} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        Else
                            mReset
'                            RptSel!edcdateAsOf.SetFocus
                            RptSel!CSI_CalAsOfDate.SetFocus
                            Exit Function
                        End If
                    End If
    
                    If Not gSetSelection(slSelection) Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else
            'Advertiser Detail
            'ElseIf ilListIndex = 1 Then
                    slStr = ""
                    If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) = GETPAIDEXPORT) Then
                        slStr = "P"     'get paid export
                    Else
                        If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS Then
                            slStr = "L"     'great plains
                        End If
                    End If
    
                    If Not gSetFormula("ShowID", "'" & slStr & "'") Then
                       gCmcGenMore = -1
                        Exit Function
                    End If
    
                    slSelection = ""
                    If Not RptSel!ckcAll.Value = vbChecked Then
                        For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                            If RptSel!lbcSelection(0).Selected(illoop) Then
                                slNameCode = tgAdvertiser(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get adf code
                                If slSelection <> "" Then
                                    slSelection = slSelection & " Or " & "{ADF_Advertisers.adfCode} =" & Trim$(slCode)
                                Else
                                    slSelection = "{ADF_Advertisers.adfCode}=" & Trim$(slCode)
                                End If
    
                            End If
                        Next illoop
                    End If
                    If Not gSetSelection(slSelection) Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If
            End If
        Case AGENCIESLIST
            If ilListIndex = 0 Then
                If RptSel!rbcType(1).Value = True Then     'summary

                    slSelection = "{@First Letter} >=" & "'" & Trim$(RptSel!cbcFrom.Text) & "'"
                    slSelection = slSelection & " And " & "{@First Letter} <=" & "'" & Trim$(RptSel!cbcTo.Text) & "'"
    
'                    slDate = RptSel!edcAsOfDate.Text
                    slDate = RptSel!CSI_CalAsOfDate.Text    '8-16-19 use csi calendar control
                    If slDate <> "" Then
    
                        If gValidDate(slDate) Then
                            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                            slSelection = slSelection & " and {AGY_Agencies.agfDateEntrd} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                        Else
                            mReset
'                            RptSel!edcAsOfDate.SetFocus
                            RptSel!CSI_CalAsOfDate.SetFocus
                            Exit Function
                        End If
                    End If
                Else

            'ElseIf ilListIndex = 1 Then
                    slStr = ""
                    If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) = GETPAIDEXPORT) Then
                        slStr = "P"     'get paid export
                    Else
                        If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS Then
                            slStr = "L"     'great plains
                        End If
                    End If
    
                    If Not gSetFormula("ShowID", "'" & slStr & "'") Then
                       gCmcGenMore = -1
                        Exit Function
                    End If
                    slSelection = ""
                    If Not RptSel!ckcAll.Value = vbChecked Then
                        For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                            If RptSel!lbcSelection(0).Selected(illoop) Then
                                slNameCode = tgAgency(illoop).sKey
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get agf code
                                If slSelection <> "" Then       'Dan M 9/17/08 changed "AGY Agencies" to "AGY_Agencies" Also changed report.
                                    slSelection = slSelection & " Or " & "{AGY_Agencies.agfCode} =" & Trim$(slCode)
                                Else
                                    slSelection = "{AGY_Agencies.agfCode}=" & Trim$(slCode)
                                End If
                            End If
                        Next illoop
                    End If
                End If

            ElseIf ilIndex = 2 Then         'mailing labels
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = ""
                gUnpackDate igNowDate(0), igNowDate(1), slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
                slSelection = "{IVR_Invoice_Rpt.ivrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({IVR_Invoice_Rpt.ivrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            End If
            If Not gSetSelection(slSelection) Then
                gCmcGenMore = -1
                Exit Function
            End If
        Case BUDGETSJOB
            slSelection = ""
            'obtain current date and time for report headers
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            If ilListIndex = 1 Then         'budget comparison, send date & time generated for filter to grf file
                slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            End If

            If RptSel!rbcSelC4(0).Value Then      'test Calendar or Standard
                If Not gSetFormula("CalType", "'C'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("CalType", "'S'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            End If

            If RptSel!rbcSelCInclude(0).Value Then      'test quarter, month, week
                If Not gSetFormula("ReportType", "'Q'") Then
                   gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf RptSel!rbcSelCInclude(1).Value Then
                If Not gSetFormula("ReportType", "'M'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ReportType", "'W'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            End If
            slNameCode = tgRptSelBudgetCode(RptSel!lbcSelection(4).ListIndex).sKey 'RptSel!lbcBudgetCode.List(RptSel!lbcSelection(4).ListIndex)
            ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
            ilRet = gParseItem(slNameYear, 2, "/", slBudgetName)    'Save to pass in formula for budget name
            ilRet = gParseItem(slNameYear, 1, "/", slBaseDate)
            slBaseDate = gSubStr("9999", slBaseDate)
            igYear = Val(slBaseDate)        'save for filtering
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMnfCode = Val(slCode)
            If ilListIndex = 1 Then                 'concatenate base budget to comparison
                                                    'budget names
                    slBudgetName = slBudgetName & " vs "
                    slMoreBudgetNames = ""
                    For illoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                        If RptSel!lbcSelection(2).Selected(illoop) Then
                            'slNameCode = RptSel!lbcBudgetCode.List(RptSel!lbcSelection(2).ListIndex)
                            slNameCode = tgRptSelBudgetCode(illoop).sKey   'RptSel!lbcBudgetCode.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slCode)    'Get application name
                            ilRet = gParseItem(slCode, 2, "/", slBaseDate)
                            If slMoreBudgetNames = "" Then
                                slMoreBudgetNames = slMoreBudgetNames & slBaseDate
                            Else
                                slMoreBudgetNames = slMoreBudgetNames & ", " & slBaseDate
                            End If
                      End If
                    Next illoop
                    slBudgetName = slBudgetName & slMoreBudgetNames
            End If
            If Not gSetFormula("BudgetName", "'" & slBudgetName & "'") Then
                gCmcGenMore = -1
                Exit Function
            End If
            slMonth = "01"          'set to calc std or corp start date of year
            slDay = "15"
            slYear = Trim(str$(igYear))     'convert year to string for month conversion

            If RptSel!rbcSelC4(0).Value Then         'corporate month?  (vs std)
                ilRet = gGetCorpCalIndex(igYear)
                'gUnpackDate tgMCof(ilRet).iStartDate(0, 1), tgMCof(ilRet).iStartDate(1, 1), slDate         'convert last bdcst billing date to string
                gUnpackDate tgMCof(ilRet).iStartDate(0, 0), tgMCof(ilRet).iStartDate(1, 0), slDate         'convert last bdcst billing date to string
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slDate = gObtainStartCorp(slMonth & "/" & slDay & "/" & Trim(str$(igYear - 1)), True)
                slDate = Format$(gDateValue(slDate), "m/d/yy")    'Start of the corporate year
            Else
                slDate = gObtainStartStd(slMonth & "/" & slDay & "/" & slYear)
                slDate = Format$(gDateValue(slDate), "m/d/yy")    'start of stdyear
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            'Budget report needs the start date of the budget year
            If ilListIndex = 0 Then         'budget report only (vs comparison)
                If Not gSetFormula("StartOfYear", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            End If
            If RptSel!rbcSelCInclude(2).Value Then  'week option- set last sunday of first week
                If RptSel!edcSelCTo.Text <> "" Then
                    slDate = RptSel!edcSelCTo.Text
                End If
                slDate = Format$(gDateValue(slDate), "m/d/yy")
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                For illoop = 1 To 13 Step 1
                    If illoop = 1 And RptSel!rbcSelC4(0).Value Then              'first time thru loop insure Corp week ends on Sun
                        slDate = gObtainNextSunday(slDate)
                    Else
                        slDate = gIncOneWeek(slDate)
                    End If
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Next illoop

            ElseIf RptSel!rbcSelCInclude(1).Value Then   'set last date of 12 std or corp months periods
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                For illoop = 1 To 13 Step 1
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                    If RptSel!rbcSelC4(0).Value Then                    'corp months
                        slDate = gObtainEndCorp(slDate, True)
                    Else                                                'std months
                        slDate = gObtainEndStd(slDate)
                    End If
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If illoop = 13 Then
                        slYear = "0"
                        slMonth = "0"
                        slDay = "0"
                    End If
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            ElseIf RptSel!rbcSelCInclude(0).Value Then   'std or corp quarters
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If Not gSetFormula("P0", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                For illoop = 1 To 12 Step 1
                    For ilIndex = 1 To 3 Step 1
                        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")   'Start date next month
                        If RptSel!rbcSelC4(0).Value Then                    'corp months
                            slDate = gObtainEndCorp(slDate, True)
                        Else                                                'std months
                            slDate = gObtainEndStd(slDate)
                        End If
                        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    Next ilIndex
                    If illoop > 4 Then
                        slYear = "0"
                        slMonth = "0"
                        slDay = "0"
                    End If
                    If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                    'Report!crcReport.Formulas(ilLoop + 1) = "P" & Trim$(Str$(ilLoop)) & "= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                Next illoop
            End If
            If ilListIndex = 0 Then             'budgets (vs comprisons)
                'Only send time printed, date can be retrieved from Crystal function
                'If Not gSetFormula("AsOfDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                '    gCmcGenMore = -1
                '    Exit Function
                'End If
'10/30/20 - TTP # 10008 - No Longer used in TextDump Crystal Reports
'                If Not gSetFormula("AsOfTime", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGenMore = -1
'                    Exit Function
'                End If
                slSelection = "{BVF_Budgets_by_Veh.bvfmnfBudget} = " & Trim$(str$(ilMnfCode)) & " And " & "{BVF_Budgets_by_Veh.bvfYear} = " & slBaseDate
            'End If
                'send selective vehicle or offices to Crystal only if not comparison selection.
                'Comparison prepass filters them out.
                If Not RptSel!ckcAll.Value = vbChecked Then         'not all vehicles or offices selected
                    If slSelection <> "" Then
                        slSelection = "(" & slSelection & ") " & " and ("
                        slOr = ""
                    Else
                        slSelection = "("
                        slOr = ""
                    End If
                    'setup selective vehicles
                    If RptSel!rbcSelCSelect(1).Value Then              'vehicle selection
                        For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                            If RptSel!lbcSelection(0).Selected(illoop) Then
                                slNameCode = tgCSVNameCode(illoop).sKey    'RptSel!lbcCSVNameCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                If ilListIndex = 0 Then                                'budget,  match on veh file
                                    slSelection = slSelection & slOr & "{BVF_Budgets_by_Veh.bvfvefCode} = " & Trim$(slCode)
                                Else                                        'budget comparisons, match on GRF file
                                    slSelection = slSelection & slOr & "{GRF_Generic_Report.grfvefCode} = " & Trim$(slCode)
                                End If
                                slOr = " Or "
                            End If
                        Next illoop
                    Else            'office selection
                        For illoop = 0 To RptSel!lbcSelection(1).ListCount - 1 Step 1
                            If RptSel!lbcSelection(1).Selected(illoop) Then
                                slNameCode = tgSOCode(illoop).sKey 'RptSel!lbcSOCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                If ilListIndex = 0 Then                 'budget report, match on SOF file
                                    slSelection = slSelection & slOr & "{BVF_Budgets_by_Veh.bvfsofCode} = " & Trim$(slCode)
                                Else                                    'budget comparisons, match on GRF file
                                    slSelection = slSelection & slOr & "{GRF_Generic_Report.grfsofCode} = " & Trim$(slCode)
                                End If
                                slOr = " Or "
                            End If
                        Next illoop
                    End If
                    slSelection = slSelection & ")"
                End If
            End If
            If Not gSetSelection(slSelection) Then
                gCmcGenMore = -1
                Exit Function
            End If
        Case RATECARDSJOB
            If ilListIndex = RC_RCITEMS Then
                If RptSel!rbcSelC4(0).Value Then      'test Calendar or Standard
                    If Not gSetFormula("CalType", "'C'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("CalType", "'S'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If

                If RptSel!rbcSelCInclude(0).Value Then      'test quarter, month, week
                    If Not gSetFormula("ReportType", "'Q'") Then
                       gCmcGenMore = -1
                        Exit Function
                    End If
                ElseIf RptSel!rbcSelCInclude(1).Value Then
                    If Not gSetFormula("ReportType", "'M'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ReportType", "'W'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If
                
                 If RptSel!rbcSelC6(0).Value Then      'avg r/c vs acq barter fee
                    If Not gSetFormula("AvgOrBarter", "'A'") Then
                       gCmcGenMore = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("AvgOrBarter", "'B'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If
                
                slSelection = ""
                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                If ilListIndex = 0 Then         'rate card prices
                    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf ilListIndex = RC_DAYPARTS Then
'10/30/20 - TTP # 10008 - No Longer used in TextDump, bofsupp, bofreplc Crystal Reports
'                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                'send generated time to report
'                If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGenMore = -1
'                    Exit Function
'                End If
            End If
        Case DALLASFEED
            'obtain current date and time for record retrieval
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

            If ilListIndex = 0 Then    'Dallas Feed
                'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", ilListIndex, "New York Feed"

                If Not gSetFormula("DumpRptName", "'Dallas Feed'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf ilListIndex = 2 Then       'Dallas Error Log
                If Not gSetFormula("DumpRptName", "'Dallas Error Log'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", 2, "New York Error Log"
            End If


            If Not gSetSelection(slSelection) Then
                gCmcGenMore = -1
                Exit Function
            End If

            'If (Not igUsingCrystal) And (ilListIndex = 0) Then
            '    If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If

            '    gDumpFileRpt 0, ilListIndex, "Dallas Feed"

            '    gCmcGenMore = 2         'successful return
            '    Exit Function
            'End If

            'If (Not igUsingCrystal) And (ilListIndex = 1) Then
            '    If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If
            '    gStudioLogRpt ilPreview, "StudioLg.Lst", Val(slLogUserCode)
            '
            '    gCmcGenMore = 2
            '    Exit Function
            'End If
            'If (Not igUsingCrystal) And (ilListIndex = 2) Then
           '     If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If

            '    gDumpFileRpt 0, 2, "Dallas Error Log"

            '   gCmcGenMore = 2         'successful return
            '    Exit Function
            'End If
        Case PHOENIXFEED
            'obtain current date and time for record retrieval
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

            If ilListIndex = 0 Then    'Phoenix Feed
                'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", ilListIndex, "New York Feed"

                If Not gSetFormula("DumpRptName", "'Commercial Studio Log'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf ilListIndex = 1 Then       'Phoenix Error Log
                If Not gSetFormula("DumpRptName", "'Phoenix Error Log'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", 2, "New York Error Log"
            End If


            If Not gSetSelection(slSelection) Then
                gCmcGenMore = -1
                Exit Function
            End If

            'If (Not igUsingCrystal) And ((ilListIndex = 0) Or (ilListIndex = 1)) Then
            '    If RptSel!rbcOutput(0).Value Then
            '        ilPreview = True
            '    ElseIf RptSel!rbcOutput(1).Value Then
            '        ilPreview = False
            '    End If
            '    If ilListIndex = 0 Then
            '        gListFileRpt 0, ilPreview, "ListFile.Lst", ilListIndex, "Phoenix Studio Log"
            '    Else
            '        gDumpFileRpt 0, 2, "Phoenix Error Log"
            '    End If
            '    gCmcGenMore = 2         'successful return
            '    Exit Function
            'End If
        Case NYFEED
            'obtain current date and time for record retrieval
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{TXR_Text_Report.txrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({TXR_Text_Report.txrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If (Not igUsingCrystal) Then
                If RptSel!rbcOutput(0).Value Then
                    ilPreview = True
                ElseIf RptSel!rbcOutput(1).Value Then
                    ilPreview = False
                End If
                gCmcGenMore = 2     'successful return from bridge
                Exit Function

            ElseIf ilListIndex = 0 Or ilListIndex = 1 Then
                If ilListIndex = 0 Then
                    'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", ilListIndex, "New York Feed"

                    If Not gSetFormula("DumpRptName", "'Engineering Feed'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else         'NY Error Log
                    If Not gSetFormula("DumpRptName", "'Engineering Error Log'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                    'gDumpFileRpt 0, ilPreview, "DumpFile.Lst", 2, "New York Error Log"
                End If
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf ilListIndex = 2 Or ilListIndex = 3 Then          'blackout replacement or suppress
                gCurrDateTime slStr, slTime, slMonth, slDay, slYear
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'                If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'                    gCmcGenMore = -1
'                    Exit Function
'                End If
'                If RptSel!edcSelCFrom.Text <> "" Then
                If RptSel!CSI_CalFrom.Text <> "" Then       '1-6-20 use csi calendar control
'                    slDateFrom = RptSel!edcSelCFrom.Text
                    slDateFrom = RptSel!CSI_CalFrom.Text
                    slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
                    slDate = slDateFrom
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("ActiveDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("ActiveDate", "'All Records'") Then
                        gCmcGenMore = -1
                        Exit Function
                    End If

                End If
                If ilListIndex = 2 Then
'                    If RptSel!edcSelCFrom.Text = "" Then
                    If RptSel!CSI_CalFrom.Text = "" Then            '1-6-20 use csi cal control
                        slSelection = "{BOF_Blackout.bofType} = 'S'"
                    Else
                        slSelection = "{BOF_Blackout.bofType} = 'S'"
                        slSelection = slSelection & "And( ( Date(" & slYear & "," & slMonth & "," & slDay & ") >= {BOF_Blackout.bofStartDate} "
                        slSelection = slSelection & "And {BOF_Blackout.bofEndDate} = Date(0,0,0) )  "
                        slSelection = slSelection & " or ( {BOF_Blackout.bofStartDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")) "
                        slSelection = slSelection & " or {BOF_Blackout.bofEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")) "
                    End If
                    If Not gSetSelection(slSelection) Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                ElseIf ilListIndex = 3 Then
'                    If RptSel!edcSelCFrom.Text = "" Then
                    If RptSel!CSI_CalFrom.Text = "" Then                '1-6-20 use csi cal control
                        slSelection = "{BOF_Blackout.bofType} = 'R'"
                    Else
                        slSelection = "{BOF_Blackout.bofType} = 'R'"
                        slSelection = slSelection & "And( ( Date(" & slYear & "," & slMonth & "," & slDay & ") >= {BOF_Blackout.bofStartDate} "
                        slSelection = slSelection & "And {BOF_Blackout.bofEndDate} = Date(0,0,0) )  "
                        slSelection = slSelection & " or ( {BOF_Blackout.bofStartDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")) "
                        slSelection = slSelection & " or {BOF_Blackout.bofEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ")) "
                    End If
                    If Not gSetSelection(slSelection) Then
                        gCmcGenMore = -1
                        Exit Function
                    End If
                End If
            Else
                gCmcGenMore = -1
                Exit Function
            End If
       Case USERLIST            '9-28-09
            If ilListIndex = USER_OPTIONS Then
                If Not gSetFormula("User", tgUrf(0).iCode) Then  'only show internal code if Guide
                    gCmcGenMore = -1
                    Exit Function
                End If
                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{UOR_User_Option_Rpt.uorGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({UOR_User_Option_Rpt.uorGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
            ElseIf ilListIndex = USER_ACTIVITY Then         '5-9-11 User Activity
                slSelection = ""
                'obtain current date and time for report headers
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{afr.afrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({afr.afrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                slDateFrom = Format$(gDateValue(RptSel!edcSelCFrom.Text), "m/d/yy")    'makesure year included
                slDate = Format$(gDateValue(RptSel!edcSelCFrom1.Text), "m/d/yy")    'makesure year included
                
                If RptSel!edcSelCTo.Text = "" Then
                    slStartTime = "12M"
                Else
                    slStartTime = Trim$((RptSel!edcSelCTo.Text))
                End If
                If RptSel!edcSelCTo1.Text = "" Then
                    slEndTime = "12M"
                Else
                    slEndTime = Trim$((RptSel!edcSelCTo1.Text))
                End If
                
                If Not gSetFormula("InputDates", "'" & slDateFrom & "-" & slDate & ", " & slStartTime & "-" & slEndTime & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                'sort selections
                mUserActivitySortSelect RptSel!cbcSet1, False, slSort
                If Not gSetFormula("Sort1", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                mUserActivitySortSelect RptSel!cbcSet2, True, slSort
                If Not gSetFormula("Sort2", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                mUserActivitySortSelect RptSel!cbcSet3, True, slSort
                If Not gSetFormula("Sort3", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                'skip selections
                slSort = gSetCheckStr(RptSel!ckcSelC10(0).Value)       'Skip, sort field 1
                If Not gSetFormula("Skip1", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                slSort = gSetCheckStr(RptSel!ckcSelC10(1).Value)       'Skip, sort field 2
                If Not gSetFormula("Skip2", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                slSort = gSetCheckStr(RptSel!ckcSelC10(2).Value)       'Skip, sort field 3
                If Not gSetFormula("Skip3", "'" & slSort & "'") Then
                    gCmcGenMore = -1
                    Exit Function
                End If
               
            ElseIf ilListIndex = USER_SUMMARY Then                      'User Summary
'                If Not gSetFormula("User", tgUrf(0).iCode) Then        'only show internal code if Guide
'                    gCmcGenMore = -1
'                    Exit Function
'                End If
                
                'Date: 4/17/2020 set report parameter "Include"
                slInclude = "B"
                If RptSel!rbcSelC6(0).Value = True Then                                 'include Active only
                    slInclude = "A"
                ElseIf RptSel!rbcSelC6(1).Value = True Then   'include Dormant only
                    slInclude = "D"
                End If
                If Not gSetFormula("Include", Chr(39) & slInclude & Chr(39)) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                'set report parameter SystemMajorSort (Sytem Type as Major Sort)
                If Not gSetFormula("SystemMajorSort", IIF(RptSel!ckcSelC7.Value = vbChecked, "'Y'", "'N'")) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                'set report parameter "SortBy"
                slSortBy = "C"                                  'City
                If RptSel!rbcSelCSelect(0).Value = True Then
                    slSortBy = "S"                              'Sign on Name
                End If
                If Not gSetFormula("SortBy", Chr(39) & slSortBy & Chr(39)) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
                
                'obtain current date and time for record retrieval
                gCurrDateTime slDate, slTime, slMonth, slDay, slYear
                slSelection = "{UOR_User_Option_Rpt.uorGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
                slSelection = slSelection & " And Round({UOR_User_Option_Rpt.uorGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
                If Not gSetSelection(slSelection) Then
                    gCmcGenMore = -1
                    Exit Function
                End If
    
            End If

       End Select
    gCmcGenMore = 1
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

    ilListIndex = RptSel!lbcRptType.ListIndex
    Select Case igRptCallType
        Case VEHICLESLIST
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Then
                If Not gOpenPrtJob("Faf.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            'ElseIf rbcRptType(1).Value Then
            ElseIf ilListIndex = 1 Then
                If Not gOpenPrtJob("Vpf.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
                'Report!crcReport.ReportFileName = sgRptPath & "Vpf.Rpt"
            ElseIf ilListIndex = 2 Then
                If Not gOpenPrtJob("Vv.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 3 Then
                If Not gOpenPrtJob("vehgrps.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 4 Then
                If Not gOpenPrtJob("logscrn.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 5 Then
                If Not gOpenPrtJob("Participants.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 6 Then
                If Not gOpenPrtJob("StdPkg.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
            
        Case ADVERTISERSLIST
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Then
                If RptSel!rbcType(1).Value = True Then
                    If Not gOpenPrtJob("Adf.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                'Report!crcReport.ReportFileName = sgRptPath & "Adf.Rpt"
            'ElseIf ilListIndex = 1 Then
                Else
                'If slSelection = "" Then
                    If Not gOpenPrtJob("AdfDet.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                    'Report!crcReport.ReportFileName = sgRptPath & "AdfDet.Rpt"
                'Else6-7-02 remove this version for selective which doesnt skip to new page each new alphabet
                '    If Not gOpenPrtJob("AdfSel.Rpt") Then
                '        gGenReport = False
                '        Exit Function
                '    End If
                    'Report!crcReport.ReportFileName = sgRptPath & "AdfSel.Rpt"
                'End If
                End If
            End If
        Case AGENCIESLIST
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Then
                If RptSel!rbcType(1).Value = True Then
                    If Not gOpenPrtJob("Agf.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                'Report!crcReport.ReportFileName = sgRptPath & "Agf.Rpt"
            'ElseIf ilListIndex = 1 Then
            
                'If slSelection = "" Then
                    If Not gOpenPrtJob("AgfDet.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                    'Report!crcReport.ReportFileName = sgRptPath & "AgfDet.Rpt"
                'Else    6-7-02 remove this version for selective which doesnt skip to new page each new alphabet
                '    If Not gOpenPrtJob("AgfSel.Rpt") Then
                '        gGenReport = False
                '        Exit Function
                '    End If
                    'Report!crcReport.ReportFileName = sgRptPath & "AgfSel.Rpt"
                'End If
            End If
            ElseIf ilListIndex = 2 Then         'mailing labels
                If RptSel!rbcSelCSelect(0).Value Then       '2-across (avery #5161/5261)
                                                        '3-across (avery #5160/5260)
                    If Not gOpenPrtJob("Label2.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("Label3.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            End If
        Case SALESPEOPLELIST
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            'If rbcRptType(0).Value Then
            If ilListIndex = 0 Then
                If Not gOpenPrtJob("Slf.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
                'Report!crcReport.ReportFileName = sgRptPath & "Slf.Rpt"
            'ElseIf rbcRptType(1).Value Then

'            ElseIf ilListIndex = 1 Then
'                If Not gOpenPrtJob("SlfFin.Rpt") Then
'                    gGenReport = False
'                    Exit Function
'                End If
                'Report!crcReport.ReportFileName = sgRptPath & "SlfFin.Rpt"
            End If
        Case EVENTNAMESLIST
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If Not gOpenPrtJob("Enf.Rpt") Then
                gGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Enf.Rpt"
        Case BUDGETSJOB
            If (ilListIndex = 0) Then   'Budgets by office or vehicle
                If (RptSel!ckcSelC3(0).Value = vbChecked) Then               'summary only
                    If (RptSel!rbcSelCSelect(0).Value) Then
                        If Not gOpenPrtJob("bgtofcsm.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("bgtvehsm.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                Else                                            'detail
                    If (RptSel!rbcSelCSelect(0).Value) Then
                        If Not gOpenPrtJob("bgtofc.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("bgtveh.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
            ElseIf (ilListIndex = 1) Then    'budget comparisons
                'slstr = RptSel!edcSelCFrom.Text
                'ilIndex = Val(slstr)
                'If ilIndex < 1 Or ilIndex > 4 Then
                '    mReset
                '   RptSel!edcSelCFrom.SetFocus
                '    Exit Function
                'End If
                If (RptSel!ckcSelC3(0).Value = vbChecked) Then              'summary only
                    If (RptSel!rbcSelCSelect(0).Value) Then
                        If Not gOpenPrtJob("bgtcmosm.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("bgtcmvsm.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                Else
                    If (RptSel!rbcSelCSelect(0).Value) Then
                        If Not gOpenPrtJob("bgtcmpof.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("bgtcmpvh.rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Case RATECARDSJOB
            If (ilListIndex = 0) Then   'RC Flights
                If Not gOpenPrtJob("rcflts.rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf (ilListIndex = RC_DAYPARTS) Then
                If Not gOpenPrtJob("rdf.rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
        Case PROGRAMMINGJOB
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If igRptType = 3 Then           'programming reports (vs links repts)
                If (ilListIndex = 0) Then   'library report option
                    If Not gOpenPrtJob("library.rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            Else                            'links
                If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Then   'Selling to airing or Conflict
                    slSelection = ""
                    'If rbcRptType(2).Value Then 'Conflict
                    If ilListIndex = 2 Then
                        If Not gOpenPrtJob("Vcf.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                        'Report!crcReport.ReportFileName = sgRptPath & "Vcf.Rpt"
                    Else
                        'If rbcRptType(0).Value Then 'Selling to Airing
                        '7-25-14
                        If (ilListIndex = 0 Or ilListIndex = 1) And RptSel!ckcInclCommentsA.Value = vbChecked Then      'selling to air links or air to sell- include avail lengths
                            If Not gOpenPrtJob("LinksAvailLen.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                    
                        ElseIf ilListIndex = 0 Then
                            If Not gOpenPrtJob("SellAir.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                            'Report!crcReport.ReportFileName = sgRptPath & "SellAir.Rpt"
                        Else    'Airing to Selling
                            If Not gOpenPrtJob("AirSell.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                            'Report!crcReport.ReportFileName = sgRptPath & "AirSell.Rpt"
                        End If
                    End If
                ElseIf (ilListIndex = 3) Or (ilListIndex = 4) Then   'delivery
                    If Not gOpenPrtJob("Delivery.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                    'Report!crcReport.ReportFileName = sgRptPath & "Delivery.Rpt"
                ElseIf (ilListIndex = 5) Or (ilListIndex = 6) Then   'delivery
                    If Not gOpenPrtJob("EngrFeed.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                ElseIf (ilListIndex = PRG_AIRING_INV) Then          '3-31
                    If RptSel!ckcInclCommentsA.Value = vbUnchecked Then           'airing vehicles inventory (without selling inventory)
                        If Not gOpenPrtJob("AirInv.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else                                                        'airing vheicle inventory with selling inventory
                        If Not gOpenPrtJob("AirInvWithSell.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
            End If                      'program reports (vs links)
        Case COLLECTIONSJOB
            If ilListIndex = 0 Then         'Date: 9/6/2018 Cash Receipt: added Contract and Invoice filters    FYM
                If (RptSel!edcContract.Text <> "") Then
                    If Not (IsNumeric(RptSel!edcContract.Text)) Then
                        mReset
                        RptSel!edcContract.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
                If (RptSel!edcInvoice.Text <> "") Then
                    If Not (IsNumeric(RptSel!edcInvoice.Text)) Then
                        mReset
                        RptSel!edcInvoice.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = 7 Then 'Credit Status
'                slDate = RptSel!edcSelA.Text   'Latest cash date
                slDate = RptSel!CSI_CalDateA.Text   'Latest cash date        8-29-19 use csi cal control vs edit box
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalDateA.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
            End If
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            'If rbcRptType(0).Value Then 'Cash receipts
            If ilListIndex = COLL_PAYHISTORY Then
                '12/17/06-Change to tax by agency or vehicle
                'If tgSpf.iBTax(0) <> 0 Or tgSpf.iBTax(1) <> 0 Then
                If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                    If RptSel!rbcSelC8(0).Value Then        'detail
                        If Not gOpenPrtJob("PymHistx.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PymHsmtx.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                Else
                    If RptSel!rbcSelC8(0).Value Then        'detail
                        If Not gOpenPrtJob("PymtHist.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("Pymhissm.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
                'Report!crcReport.ReportFileName = sgRptPath & "Cash.Rpt"
            'ElseIf rbcRptType(1).Value Then 'Ageing
            ElseIf (ilListIndex = COLL_AGEPAYEE) Then
                If RptSel!rbcSelCSelect(1).Value Then    'Tran level
                    If Not gOpenPrtJob("AgePayTr.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("AgePay.Rpt") Then   'detal, invoice, summary
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf (ilListIndex = COLL_AGESLSP) Then
                If RptSel!ckcOption.Value = vbChecked Then          '2-25-19 split slsp option added
                    If Not gOpenPrtJob("AgeSlsSplit.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("AgeSls.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
                '7-11-02 4 total levels added to aging by slsp.  Incorporate all into agesls.rpt
                'If RptSel!rbcSelCSelect(0).Value Then   'detail version of slsp ageing
                '    If Not gSetFormula("DetailSummary", "'D'") Then
                '        gGenReport = False
                '        Exit Function
                '    End If
                'Else
                '    If Not gSetFormula("DetailSummary", "'S'") Then
                '        gGenReport = False
                '        Exit Function
                '    End If
                'End If
                '1-8-02 Summary version of Slsp ageing has been combined with the Detail version
                'Report!crcReport.ReportFileName = sgRptPath & "AgePay.Rpt"
                'Else
                '    If Not gOpenPrtJob("AgeSlSum.Rpt") Then
                '        gGenReport = False
                '        Exit Function
                '    End If
                '    'Report!crcReport.ReportFileName = sgRptPath & "AgeSls.Rpt"
               'End If
            ElseIf (ilListIndex = COLL_AGEVEHICLE) Then
                'If Not RptSel!ckcSelC3(0).Value = vbChecked Then    'vehicle group checked?
                    '1-14-02 make detail & summary one crystal report by use of suppressing the detail for summary version
                    'If RptSel!rbcSelCSelect(0).Value Then           'detail
                    If RptSel!ckcSelC11(0).Value = vbChecked Then       'extended ageing columns
                        If Not gOpenPrtJob("AgeVehYr.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("AgeVeh.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If

                    'Else    'summary
                    '    If Not gOpenPrtJob("AgeVeSum.Rpt") Then
                    '        gGenReport = False
                    '        Exit Function
                    '    End If

                    'End If
                'Else                        'include vehicle group totals
                '    If RptSel!rbcSelCSelect(0).Value Then
                '        If Not gOpenPrtJob("AgeVehGp.Rpt") Then
                '            gGenReport = False
                '            Exit Function
                '        End If
                '
                '    Else
                '        If Not gOpenPrtJob("AgeVSGp.Rpt") Then
                '            gGenReport = False
                '            Exit Function
                '        End If
                '
                '    End If
                'End If
            'ElseIf rbcRptType(2).Value Then 'Delinquent
            ElseIf ilListIndex = 4 Then
                If Not gOpenPrtJob("Delinq.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            'ElseIf rbcRptType(3).Value Then 'Statements
            ElseIf ilListIndex = COLL_STATEMENT Then
                blCustomizeLogo = True
                If RptSel!rbcSelCInclude(0).Value Then            '8-2-00 detail
                    If Not gOpenPrtJob("StateC.Rpt", , blCustomizeLogo) Then
                        gGenReport = False
                        Exit Function
                    End If
                ElseIf RptSel!rbcSelCInclude(1).Value Then     'tran type
                    If Not gOpenPrtJob("StateCSM.Rpt", , blCustomizeLogo) Then
                        gGenReport = False
                        Exit Function
                    End If
                Else                                            'invoice #
                    If Not gOpenPrtJob("StateINV.Rpt", , blCustomizeLogo) Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = COLL_CASH Then
                If RptSel!rbcSelC4(0).Value Or RptSel!rbcSelC4(2).Value Then             '11-17-05 sort by date or vehicle group
                    '12/17/06-Change to tax by agency or vehicle
                    'If tgSpf.iBTax(0) <> 0 Or tgSpf.iBTax(1) <> 0 Then
                    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                        If Not gOpenPrtJob("Cashtax.Rpt") Then
                                gGenReport = False
                                Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("Cash.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                ElseIf RptSel!rbcSelC4(3).Value Then                '9-20-17 sort by ageing period
                    If Not gOpenPrtJob("CashByAgeing.Rpt") Then     'created to aid in balancing Participant Payables reort
                        gGenReport = False
                        Exit Function
                    End If
                Else                                        'sort by slsp
                    '12/17/06-Change to tax by agency or vehicle
                    'If tgSpf.iBTax(0) <> 0 Or tgSpf.iBTax(1) Then
                    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                        If RptSel!rbcSelC12(1).Value = True Then        'subtotals by advt,tax version
                            If Not gOpenPrtJob("CashAdvTax.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                        Else                                            'subtotals by check #
                            If Not gOpenPrtJob("CashsTax.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                        End If
                    Else
                        If RptSel!rbcSelC12(1).Value Then               'subtotals by advt, non tax version
                            If Not gOpenPrtJob("CashAdv.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                        Else
                            If Not gOpenPrtJob("Cashsls.Rpt") Then
                                gGenReport = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            ElseIf ilListIndex = COLL_SALESCOMM_COLL Then
                If Not gOpenPrtJob("CommColl.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If                'Report!crcReport.ReportFileName = sgRptPath & "Phf.Rpt"
            ElseIf ilListIndex = 7 Then 'Credit Status
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
            ElseIf ilListIndex = COLL_DISTRIBUTE Then             'cash distribution by participants
                If RptSel!rbcOutput(3).Value = False Then             'not Export -> use crystal
                    If RptSel!rbcSelCSelect(0).Value Then             'by inv
                        If Not gOpenPrtJob("DistInv.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    ElseIf RptSel!rbcSelCSelect(1).Value Then             'by check
                        If Not gOpenPrtJob("Distchk.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("Distpart.Rpt") Then          'by participation
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
            ElseIf ilListIndex = COLL_CASHSUM Then                      'Cash summary
                gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr       'previous billing period
                llDate = gDateValue(slStr)
'                slStr = RptSel!edcSelCFrom.Text
                slStr = RptSel!CSI_CalFrom.Text         '8-29-19 csi calendar control vs edit box
                If (Not gValidDate(slStr)) Then '7-14-03 allow to get into the past
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    Exit Function
                End If
                slStr = RptSel!CSI_CalTo.Text
                If Not gValidDate(slStr) Then
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    Exit Function
                End If
                'D.S. 08/15/01 veh or offfice choice
                If RptSel!rbcSelCSelect(0).Value = True Then
                    If Not gOpenPrtJob("CashSum.Rpt") Then           'cash summary by vehicle
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("CashSumS.Rpt") Then           'cash summary by sales office
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = COLL_ACCTHIST Then                'Account History
                If Not gOpenPrtJob("AcctHist.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COLL_MERCHANT Then
                slStr = RptSel!edcSelCFrom.Text
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSel!edcSelCFrom1.SetFocus                 'invalid year
                    gGenReport = False
                    Exit Function
                End If
                'igYear = Val(RptSelCt!edcSelCFrom1.Text)
                slStr = RptSel!edcSelCFrom1.Text
                ilRet = gVerifyInt(slStr, 1, 4)
                If ilRet = -1 Then
                    mReset
                    RptSel!edcSelCFrom.SetFocus                 'invalid qtr
                    gGenReport = False
                    Exit Function
                End If

                slStr = RptSel!edcCheck.Text
                llTemp = gVerifyLong(slStr, 0, 999999999)
                If llTemp = -1 Then                     'error
                    mReset
                    RptSel!edcCheck.SetFocus
                    Exit Function
                End If

                If RptSel!rbcSelCSelect(0).Value Then       'Sort by Vehicle (vs advt)
                    If RptSel!rbcSelCInclude(0).Value Then  'Detail (vs summary)
                        If Not gOpenPrtJob("MchVehDt.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("MchVehSm.Rpt") Then     'summary
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                Else
                    If RptSel!rbcSelCInclude(0).Value Then  'Detail
                        If Not gOpenPrtJob("MchAdvDt.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("MchAdvSm.Rpt") Then     'sort by advt, summary
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
                slStr = RptSel!edcSelCFrom.Text
                igYear = gVerifyYear(slStr)
                If igYear = 0 Then
                    mReset
                    RptSel!edcSelCTo.SetFocus                 'invalid year
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COLL_MERCHRECAP Then
                If mVerifyDate(RptSel!edcSelCFrom, llDate, False) Then
                    gGenReport = False
                    Exit Function
                End If
                If mVerifyDate(RptSel!edcSelCFrom1, llDate2, False) Then
                    gGenReport = False
                    Exit Function
                End If
                If llDate2 < llDate And llDate2 <> 0 Then
                    gGenReport = False
                    mReset
                    RptSel!edcSelCFrom.SetFocus
                    Exit Function
                End If
                If Not gOpenPrtJob("MchRecap.Rpt") Then         'Merch/Promo Recap
                    gGenReport = False
                End If
            ElseIf ilListIndex = COLL_AGEOWNER Then
                If Not gOpenPrtJob("AgeOwner.Rpt") Then         'detail & summary use same version
                    gGenReport = False
                End If
            ElseIf ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Then
                If Not gOpenPrtJob("AgeSS.Rpt") Then         'detail & summary use same version
                    gGenReport = False
                End If
            ElseIf ilListIndex = COLL_POAPPLY Then      '9-11-03
                If Not gOpenPrtJob("POApply.Rpt") Then         'POs Applied
                    gGenReport = False
                End If
            ElseIf ilListIndex = COLL_AGEMONTH Then
                'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
                If RptSel!rbcOutput(3).Value = False Then
                    If Not gOpenPrtJob("AgeMonth.Rpt") Then
                        gGenReport = False
                    End If
                End If
            End If
        Case POSTLOGSJOB
'            slDateFrom = ""
'            slDateTo = ""
'            'Date selection passed by formula
'            slDate = RptSel!edcSelCFrom.Text
'            If (slDate <> "") Then
'                If Not gValidDate(slDate) Then
'                    mReset
'                    RptSel!edcSelCFrom.SetFocus
'                    gGenReport = False
'                    Exit Function
'                End If
'            Else
'                mReset
'                RptSel!edcSelCFrom.SetFocus
'                gGenReport = False
'                Exit Function
'            End If
'            slDate = RptSel!edcSelCTo.Text 'Latest billing date
'            If (slDate <> "") Then
'                If Not gValidDate(slDate) Then
'                    mReset
'                    RptSel!edcSelCTo.SetFocus
'                    gGenReport = False
'                    Exit Function
'                End If
'            Else
'                mReset
'                RptSel!edcSelCTo.SetFocus
'                gGenReport = False
'                Exit Function
'            End If
'            If Not igUsingCrystal Then
'                gGenReport = True
'                Exit Function
'            End If
            'If rbcRptType(0).Value Then
            If (ilListIndex = 0) Or (ilListIndex = 1) Then
                If Not gOpenPrtJob("PLog.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = PL_LIVELOG Then            '12-8-05
                If Not gOpenPrtJob("LiveLogAct.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = PL_STATION_POSTING Then        '1-23-19
                If Not gOpenPrtJob("StationPostActivity.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
        Case COPYJOB
            '0 - copy status by date, 1 = copy status by advt, 2= cnts missing copy, 18 = script
            If (ilListIndex = 0) Or (ilListIndex = 1) Or (ilListIndex = 2) Or (ilListIndex = COPY_SCRIPTAFFS) Then          '4-9-12 add script affs
            '8-22-19 use csi calendar control vs edit box
'                gGenReport = gVerifyDate(RptSel, RptSel!edcSelCFrom)
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalFrom)
                If Not gGenReport Then
                    Exit Function
                End If

'                gGenReport = gVerifyDate(RptSel, RptSel!edcSelCTo)
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalTo)
                If Not gGenReport Then
                    Exit Function
                End If
      
            ElseIf ilListIndex = COPY_ROT Then 'Rotation by Advertiser
                'verify active date span
                '8-23-19 following changed to use csi calendar control vs text box
'                gGenReport = gVerifyDate(RptSel, RptSel!edcSelCFrom)
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalFrom)
                If Not gGenReport Then
                    Exit Function
                End If
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalTo)
                If Not gGenReport Then
                    Exit Function
                End If
                'verify rot date entered date span
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalFrom2)
                If Not gGenReport Then
                    Exit Function
                End If
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalTo2)
                If Not gGenReport Then
                    Exit Function
                End If
            ElseIf ilListIndex = 4 Then 'Inventory by Number
            ElseIf ilListIndex = 5 Then 'Inventory by ISCI
            ElseIf ilListIndex = 6 Then 'Inventory by Advertiser
            ElseIf ilListIndex = 7 Then 'Inventory by Start Date
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                slDate = RptSel!CSI_CalTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 8 Then 'Inventory by Expiration
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                slDate = RptSel!CSI_CalTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 9 Then 'Inventory by Purge
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                slDate = RptSel!CSI_CalTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 10 Then 'Inventory by Entry Date
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                slDate = RptSel!CSI_CalTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COPY_INVPRODUCER Then '4-10-13
'                gGenReport = gVerifyDate(RptSel, RptSel!edcSelCFrom)
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalFrom)
                If Not gGenReport Then
                    Exit Function
                End If

'                gGenReport = gVerifyDate(RptSel, RptSel!edcSelCTo)
                gGenReport = gVerifyDate(RptSel, RptSel!CSI_CalTo)
                If Not gGenReport Then
                    Exit Function
                End If

                If Not gOpenPrtJob("CifProd.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 11 Or ilListIndex = 13 Or ilListIndex = 14 Then    'Play List
'                slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   'Latest cash date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!CSI_CalFrom.SetFocus
                    gGenReport = False
                    Exit Function
                End If
'                slDate = RptSel!edcSelCTo.Text 'Latest billing date
                slDate = RptSel!CSI_CalTo.Text 'Latest billing date
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
                    RptSel!edcSelCTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COPY_INVUNAPPROVED Then       'unapproved copy
                slDateFrom = ""
                slDateTo = ""
'                slDate = RptSel!edcSelCTo.Text
                slDate = RptSel!CSI_CalTo.Text
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                End If
            End If
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If ilListIndex = 0 Then
                If Not gOpenPrtJob("CpyStsDt.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 1 Then
                If Not gOpenPrtJob("CpyStsAd.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 2 Then
                If RptSel!rbcSelCInclude(0).Value Then      'show by contract (vs line)
                    If Not gOpenPrtJob("CopyCntr.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("CopyCntrLine.Rpt") Then 'show by line
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = COPY_ROT Then
                If RptSel!ckcSelC7.Value = vbChecked Then       'show extra assign dates
                    If Not gOpenPrtJob("RotDetail.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("Rot.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = 4 Then
                If Not gOpenPrtJob("CifNum.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 5 Then
                If Not gOpenPrtJob("CifISCI.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 6 Then
                If Not gOpenPrtJob("CifAdv.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 7 Then
                If Not gOpenPrtJob("CifStr.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 8 Then
                If Not gOpenPrtJob("CifExp.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 9 Then
                If Not gOpenPrtJob("CifPur.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 10 Then
                If Not gOpenPrtJob("CifEntry.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 11 Then        'playlist by isci
'                If RptSel!rbcSelCSelect(0).Value = True Then
'                    If Not gOpenPrtJob("Ply.Rpt") Then         'no longer exists
'                        gGenReport = False
'                        Exit Function
'                    End If
'                Else
                    If RptSel!rbcSelC8(0).Value = True Then         'no split copy
                        If Not gOpenPrtJob("PlyISCI.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("PlyISCIRegion.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                'End If
            ElseIf ilListIndex = 12 Then
                If Not gOpenPrtJob("Cifunapr.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 13 Then                'playlist by vehicle
                If tgSpf.sUseCartNo <> "N" Then             'Playlist:  using cart #
                    If Not gOpenPrtJob("L09.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else                                        'not using cart #s, but retrieve it from reel # field if they are put in
                    If Not gOpenPrtJob("L17.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf ilListIndex = 14 Then            'playlist by advt
                If Not gOpenPrtJob("PlyAdvt.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
'            ElseIf ilListIndex = COPY_REGIONS Then      '2-12-09 chged to splitregionlist
'                If Not gOpenPrtJob("raf.Rpt") Then
'                    gGenReport = False
'                    Exit Function
'                    End If
            ElseIf ilListIndex = COPY_BOOK Then      '8-30-05
                If Not gOpenPrtJob("copybook.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COPY_SPLITROT Then '1-30-09
                If Not gOpenPrtJob("SplitBlackRot.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = COPY_SCRIPTAFFS Then '4-9-12
                If Not gOpenPrtJob("ScriptAffs.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
        Case INVOICESJOB
            If ilListIndex = INV_REGISTER Or ilListIndex = INV_DISTRIBUTE Then  'Invoice register
 '               slDate = RptSel!edcSelCFrom.Text   'Latest cash date
                slDate = RptSel!CSI_CalFrom.Text   ' 8-15-19
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
'                        RptSel!edcSelCFrom.SetFocus
                        RptSel!CSI_CalFrom.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
'                    RptSel!edcSelCFrom.SetFocus
                    RptSel!CSI_CalFrom.SetFocus
                    gGenReport = False
                    Exit Function
                End If
'                slDate = RptSel!edcSelCTo.Text   'Latest cash date
                slDate = RptSel!CSI_CalTo.Text   '8-15-19
                If (slDate <> "") Then
                    If Not gValidDate(slDate) Then
                        mReset
'                        RptSel!edcSelCTo.SetFocus
                        RptSel!CSI_CalTo.SetFocus
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    mReset
'                    RptSel!edcSelCTo.SetFocus
                    RptSel!CSI_CalTo.SetFocus
                    gGenReport = False
                    Exit Function
                End If
            End If
            'If Not igUsingCrystal Then
             '   gGenReport = True
             '   Exit Function
            'End If
            If ilListIndex = INV_REGISTER Then
                If RptSel!rbcSelCSelect(0).Value Then               'invoice option
                    If RptSel!rbcSelCInclude(0).Value Then          'detail with airing vehicles
                        If Not gOpenPrtJob("InvRegal.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    ElseIf RptSel!rbcSelCInclude(1).Value Then      'detail with billing vehicles
                        If Not gOpenPrtJob("InvRegIn.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("InvRegIt.Rpt") Then         'summary
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                ElseIf RptSel!rbcSelCSelect(9).Value Then               'sales origin, unique sort
                    If RptSel!rbcSelCInclude(0).Value Then          'detail
                        If Not gOpenPrtJob("InvRegSO.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("InvRegSOSum.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                Else                                                    'advt,agy slsp, bill & airing vehicle options
                    If (RptSel!rbcSelCInclude(0).Value) Or (RptSel!rbcSelCInclude(1).Value And RptSel!rbcSelCSelect(8).Value) Then   'detail or Sales Source w/summary version
                        If Not gOpenPrtJob("InvRegDt.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    Else
                        If Not gOpenPrtJob("InvRegSm.Rpt") Then
                            gGenReport = False
                            Exit Function
                        End If
                    End If
                End If
            'TTP 10118 -Billing Distribution Export to CSV
            ElseIf (ilListIndex = INV_DISTRIBUTE) And RptSel!rbcOutput(3).Value = False Then
                If RptSel!rbcSelCInclude(0).Value Then          'detail
                    If Not gOpenPrtJob("Invowndt.Rpt") Then
                        gGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("Invownsm.Rpt") Then         'summary
                        gGenReport = False
                        Exit Function
                    End If
                End If
            ElseIf (ilListIndex = INV_VIEWEXPORT) Then
                If Not gOpenPrtJob("TextDump.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
                If Not gSetFormula("DumpRptName", "'Invoice Export'") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf (ilListIndex = INV_CREDITMEMO) Then       '10-8-03
                If Not gOpenPrtJob("Inv_Memo.Rpt", , True) Then          'this report uses custom logo
                    gGenReport = False
                    Exit Function
                End If
            ElseIf (ilListIndex = INV_SUMMARY) Then       '6-28-05
                If Not gOpenPrtJob("InvSummary.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = INV_TAXREGISTER Then
                If Not gOpenPrtJob("TaxReg.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = INV_RECONCILE Then         '11-30-07
                If Not gOpenPrtJob("InstallRecon.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = INV_UNPOSTED_STATIONS Then     '8-12-15
                If Not gOpenPrtJob("UnpostedStations.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
        Case CHFCONVMENU
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If Not gOpenPrtJob("Icf.Rpt") Then
                gGenReport = False
                Exit Function
            End If
        Case GENERICBUTTON
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            slStr = RptSel!edcSelA.Text
            If InStr(slStr, ".") = 0 Then
                slStr = slStr & ".Rpt"
            End If
            If Not gOpenPrtJob("Generic\" & slStr) Then
                gGenReport = False
                Exit Function
            End If
        Case DALLASFEED
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If ilListIndex = 0 Or ilListIndex = 2 Then
                If Not gOpenPrtJob("TextDump.rpt") Then
                    gGenReport = False
                End If
            ElseIf ilListIndex = 1 Then         'Studio Log
                If Not gOpenPrtJob("StudioLg.rpt") Then
                    gGenReport = False
                End If
            End If
        Case NYFEED
            If ilListIndex = 0 Or ilListIndex = 1 Then         'NY Feed or NY Error Log
                If Not gOpenPrtJob("TextDump.rpt") Then
                    gGenReport = False
                End If
            ElseIf ilListIndex = 2 Then
                If Not gOpenPrtJob("bofsupp.rpt") Then
                    gGenReport = False
                End If
            ElseIf ilListIndex = 3 Then
                If Not gOpenPrtJob("bofreplc.rpt") Then
                    gGenReport = False
                End If
            ElseIf Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
        Case PHOENIXFEED
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            Else
                If Not gOpenPrtJob("TextDump.rpt") Then
                    gGenReport = False
                End If
            End If
        Case CMMLCHG
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            Else
                If Not gOpenPrtJob("TextList.rpt") Then  'no logos on this version
                    gGenReport = False
                End If
            End If
        Case EXPORTAFFSPOTS
            If Not gOpenPrtJob("TextDump.rpt") Then
                gGenReport = False
            End If

        Case BULKCOPY
            If Not igUsingCrystal Then
                gGenReport = True
                Exit Function
            End If
            If ilListIndex = 0 Or ilListIndex = 1 Then
                If Not gOpenPrtJob("TextList.rpt") Then  'no logos on this version
                    gGenReport = False
                End If
            ElseIf ilListIndex = 2 Then
                If Not gOpenPrtJob("CyfCart.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 3 Then
                If Not gOpenPrtJob("CyfVeh.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 4 Then
                If Not gOpenPrtJob("CyfDate.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            ElseIf ilListIndex = 5 Then
                If Not gOpenPrtJob("CyfAdv.Rpt") Then
                    gGenReport = False
                    Exit Function
                End If
            End If
         Case USERLIST           '9-28-09
            If ilListIndex = USER_OPTIONS Then
                If Not gOpenPrtJob("UserOptions.rpt") Then
                    gGenReport = False
                End If
            ElseIf ilListIndex = USER_ACTIVITY Then     '5-9-11
                'validity check start date
                slStr = RptSel!edcSelCFrom.Text
                If Not gValidDate(slStr) Then
                    mReset
                    RptSel!edcSelCFrom.SetFocus
                    Exit Function
                End If
                'validity check end date
                slStr = RptSel!edcSelCFrom1.Text
                If Not gValidDate(slStr) Then
                    mReset
                    RptSel!edcSelCFrom1.SetFocus
                    Exit Function
                End If
                'validity check start time
                slStr = RptSel!edcSelCTo.Text
                If Not gValidTime(slStr) Then
                    mReset
                    RptSel!edcSelCTo.SetFocus
                    Exit Function
                End If
                'validity check end time
                slStr = RptSel!edcSelCTo1.Text
                If Not gValidTime(slStr) Then
                    mReset
                    RptSel!edcSelCTo1.SetFocus
                    Exit Function
                End If
                
                If Not gOpenPrtJob("UserActivity.rpt") Then
                    gGenReport = False
                End If
            ElseIf ilListIndex = USER_SUMMARY Then
                If Not gOpenPrtJob("UserSummary.rpt") Then
                    gGenReport = False
                End If
            End If
    End Select
    gGenReport = True
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
    If Not RptSel!ckcSelC3(0).Value = vbChecked Then 'VBC NR
        'exclude trades
        slExclude = "Trades" 'VBC NR
        slSelection = slSelection & " And {CHF_Contract_Header.chfPctTrade} <> 10000" 'VBC NR
    Else                                'show trades as inclusion 'VBC NR
        slInclude = "Trades" 'VBC NR
    End If 'VBC NR
    If Not RptSel!ckcSelC3(1).Value = vbChecked Then 'VBC NR
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
    If Not RptSel!ckcSelC3(2).Value = vbChecked Then 'VBC NR
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
    If Not RptSel!ckcSelC3(3).Value = vbChecked Then 'VBC NR
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

    If Not RptSel!ckcSelC3(4).Value = vbChecked Then 'VBC NR
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

    If Not RptSel!ckcSelC3(5).Value = vbChecked Then 'VBC NR
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
'
'
'               Verify input start and end dates in edcselcFrom & edcSelCTo edit boxes
'
'               Send formula in DateRG with dates requested description
'
'               <input> - ilDatesReqd: true if dates required, else false for entire span
'               Return - true (-1) if error, ELSE 0
'               7-18-00
'
'
Function mCopyDates(ilDatesReqd As Integer) As Integer
Dim slDateFrom As String
Dim slDateTo As String
Dim slDate As String
        mCopyDates = 0
        slDateFrom = ""
        slDateTo = ""
        'Date selection passed by formula
'        slDate = RptSel!edcSelCFrom.Text   'Latest cash date
        slDate = RptSel!CSI_CalFrom.Text
        slDateFrom = slDate
        If (slDate <> "") Then
            If gValidDate(slDate) Then
                slDateFrom = Format$(gDateValue(slDate), "m/d/yy")
            Else
                mReset
                RptSel!CSI_CalFrom.SetFocus
                Exit Function
            End If
        Else
            If ilDatesReqd Then
                mReset
                RptSel!CSI_CalFrom.SetFocus
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCTo.Text 'Latest billing date
        slDate = RptSel!CSI_CalTo.Text 'Latest billing date
        slDateTo = slDate
        If (slDate <> "") Then
            If gValidDate(slDate) Then
                slDateTo = Format$(gDateValue(slDate), "m/d/yy")
            Else
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
        Else
            If ilDatesReqd Then
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
        End If
        If slDateFrom = "" And slDateTo = "" Then
            If Not gSetFormula("DateRg", "'" & "All Dates" & "'") Then
                mCopyDates = -1
                Exit Function
            End If
        ElseIf slDateFrom = "" And slDateTo <> "" Then
            If Not gSetFormula("DateRg", "'" & "thru " & slDateTo & "'") Then
                mCopyDates = -1
                Exit Function
            End If
        ElseIf slDateFrom <> "" And slDateTo = "" Then
            If Not gSetFormula("DateRg", "'" & "from  " & slDateFrom & "'") Then
                mCopyDates = -1
                Exit Function
            End If
        Else
            If gDateValue(slDateFrom) <> gDateValue(slDateTo) Then
                If Not gSetFormula("DateRg", "'" & slDateFrom & " - " & slDateTo & "'") Then
                    mCopyDates = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("DateRg", "'" & "for " & slDateFrom & "'") Then
                    mCopyDates = -1
                    Exit Function
                End If
            End If
        End If
End Function
'
'
'           5-2-01 dh Selection of records was not sent as a formula to Crystal for
'                      Copy STatus by advt or Date
Function mCopyJob(ilListIndex As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slOr                                                                                  *
'******************************************************************************************

    Dim slDateFrom As String
    Dim slSelection As String
    Dim slDate As String
    Dim slDateTo As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTime As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim illoop As Integer
    Dim ilPreview As Integer
    Dim ilOrAnd As Integer      '0 =do OR equal filter testing, 1 = do AND not equal filter testing
    Dim slInclude As String
    Dim slExclude As String

    mCopyJob = 0
    If (ilListIndex = 0) Or (ilListIndex = 1) Then  'copy status by date, copy status by advt
        If RptSel!rbcOutput(0).Value Then
            ilPreview = True
        ElseIf RptSel!rbcOutput(1).Value Then
            ilPreview = False
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If ilListIndex = 0 Then
            gCopyDateRpt
        Else
            gCopyAdvtRpt
        End If
        If Not gSetSelection(slSelection) Then   '5-2-01
            mCopyJob = -1
            Exit Function
        End If

        mCopyJob = 1
        Exit Function
    ElseIf (ilListIndex = 2) Then       'contracts missing copy
        'Date selection passed by formula
'        slDateFrom = RptSel!edcSelCFrom.Text   'Latest cash date
'        slDateTo = RptSel!edcSelCTo.Text 'Latest billing date
        slDateFrom = RptSel!CSI_CalFrom.Text
        slDateTo = RptSel!CSI_CalTo.Text

        gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
        If Not gSetFormula("StartDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            mCopyJob = -1
            Exit Function
        End If
        gObtainYearMonthDayStr slDateTo, True, slYear, slMonth, slDay
        If Not gSetFormula("EndDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            mCopyJob = -1
            Exit Function
        End If
        If RptSel!ckcSelC5(0).Value = vbChecked Then    'show whether fills included/excluded in header
            slStr = ""
        Else
            slStr = "(Fill excluded)"
        End If
        If Not gSetFormula("Subhead", "'" & slStr & "'") Then
            mCopyJob = -1
            Exit Function
        End If

        gIncludeExcludeCkc RptSel!ckcSelC5(0), slInclude, slExclude, "Fill spots"
        gIncludeExcludeCkc RptSel!ckcSelC3(0), slInclude, slExclude, "Copy Unassigned"
        gIncludeExcludeCkc RptSel!ckcSelC3(1), slInclude, slExclude, "Copy to Reassign "
        gIncludeExcludeCkc RptSel!ckcSelC3(2), slInclude, slExclude, "Missed Spots"

        'only show inclusion/exclusion of cntr/feed spots if not included

        If tgSpf.sSystemType = "R" Then         'radio system are only ones that can have feed spots
            If Not RptSel!ckcSelC10(0).Value = vbChecked Then
                gIncludeExcludeCkc RptSel!ckcSelC10(0), slInclude, slExclude, "Contract spots"
            End If
            If Not RptSel!ckcSelC10(1).Value = vbChecked Then
                gIncludeExcludeCkc RptSel!ckcSelC10(1), slInclude, slExclude, "Feed spots"
            End If
        End If
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                mCopyJob = -1
            End If
        Else
            If Not gSetFormula("Included", "'" & " " & "'") Then
                mCopyJob = -1
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
               mCopyJob = -1
            End If
        Else
            If Not gSetFormula("Excluded", "'" & " " & "'") Then
                mCopyJob = -1
            End If
        End If


        If RptSel!rbcSelCSelect(0).Value Then           'sort by advertiser
            If Not gSetFormula("SortBy", "'A'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else                                            'sort by vehicle
            If Not gSetFormula("SortBy", "'V'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

        If RptSel!ckcSelC7.Value Then           'New Page each vehicle? (if by advertiser, it has been forced to NO)
            If Not gSetFormula("NewPage", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
         '****************************ds
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

        'mCopyJob = 2
        'Exit Function
    ElseIf ilListIndex = COPY_ROT Then 'Rotation by Advertiser

'        ilRet = mInpDateFormula(RptSel!edcSelCFrom, RptSel!edcSelCFrom1, "RptDates")
'        ilRet = mInpDateFormula(RptSel!edcSelCTo, RptSel!edcSelCTo1, "EntDates")
        ilRet = mInpDateFormula(RptSel!CSI_CalFrom, RptSel!CSI_CalTo, "RptDates")
        ilRet = mInpDateFormula(RptSel!CSI_CalFrom2, RptSel!CSI_CalTo2, "EntDates")

        If RptSel!ckcSelC5(0).Value = vbChecked Then
            If Not gSetFormula("ShowInventory", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowInventory", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        
        '8-3-10 option to show rotation comments
        If RptSel!ckcTrans.Value = vbChecked Then
            If Not gSetFormula("IncludeComments", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("IncludeComments", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    ElseIf ilListIndex = COPY_INVBYNUMBER Then 'Inventory by Number
        slStr = RptSel!edcSelCFrom.Text
        If (slStr <> "") Then
            If Not gSetFormula("Low Cart", "'" & slStr & "'") Then
                mCopyJob = -1
                Exit Function
            End If
            slSelection = "({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) >= '" & slStr & "'"
        End If
        slStr = RptSel!edcSelCTo.Text
        If (slStr <> "") Then
            If Not gSetFormula("Hi Cart", "'" & slStr & "'") Then
                mCopyJob = -1
                Exit Function
            End If
            If slSelection = "" Then
                slSelection = "({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) <= '" & slStr & "'"
            Else
                slSelection = slSelection & " And ({MCF_Media_Code.mcfName} + {CIF_Copy_Inventory.cifName}) <= '" & slStr & "'"
            End If
        End If
    ElseIf ilListIndex = COPY_INVBYISCI Then 'Inventory by ISCI
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        slStr = RptSel!edcSelCFrom.Text
        If (slStr <> "") Then
            If Not gSetFormula("Low Cart", "'" & slStr & "'") Then
                mCopyJob = -1
                Exit Function
            End If
            slSelection = "{CPF_Copy_Prodct_ISCI.cpfISCI} >= '" & slStr & "'"
        End If
        slStr = RptSel!edcSelCTo.Text
        If (slStr <> "") Then
            If Not gSetFormula("Hi Cart", "'" & slStr & "'") Then
                mCopyJob = -1
                Exit Function
            End If
            If slSelection = "" Then
                slSelection = "{CPF_Copy_Prodct_ISCI.cpfISCI} <= '" & slStr & "'"
            Else
                slSelection = slSelection & " And {CPF_Copy_Prodct_ISCI.cpfISCI} <= '" & slStr & "'"
            End If
        End If
        If slSelection = "" Then
            slSelection = "{CPF_Copy_Prodct_ISCI.cpfISCI} <> ''"
        Else
            slSelection = slSelection & " And {CPF_Copy_Prodct_ISCI.cpfISCI} <> ''"
        End If
    ElseIf ilListIndex = COPY_INVBYADVT Then 'Inventory by Advertiser
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        If Not RptSel!ckcAll.Value = vbChecked Then
            ilOrAnd = gGetCountSelected(0)
            slSelection = ""
            If ilOrAnd = 0 Then                         'selected less than half the list box
                'filter using standard OR equal testing
                For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                    If RptSel!lbcSelection(0).Selected(illoop) Then
                        slNameCode = tgAdvertiser(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If slSelection <> "" Then
                            slSelection = slSelection & " Or " & "{CIF_Copy_Inventory.cifadfCode} = " & Trim$(slCode)
                        Else
                            slSelection = "{CIF_Copy_Inventory.cifadfCode} = " & Trim$(slCode)
                        End If
                    End If
                Next illoop
            Else            'selected more than half, filter using AND not equal testing
                            'so that limits are not exceeded
                For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                    If Not RptSel!lbcSelection(0).Selected(illoop) Then
                        slNameCode = tgAdvertiser(illoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If slSelection <> "" Then
                            slSelection = slSelection & " And " & "{CIF_Copy_Inventory.cifadfCode} <> " & Trim$(slCode)
                        Else
                            slSelection = "{CIF_Copy_Inventory.cifadfCode} <> " & Trim$(slCode)
                        End If
                    End If
                Next illoop
            End If
        End If
    '11-22-06 changed touse prepass
    ElseIf ilListIndex = COPY_INVBYSTARTDATE Then 'Inventory by Start Date
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

        If RptSel!rbcSelC6(0).Value = True Then         'carted
            If Not gSetFormula("CartedFlag", "'C'") Then
                mCopyJob = -1
                Exit Function
            End If
        ElseIf RptSel!rbcSelC6(1).Value = True Then     'uncarted
            If Not gSetFormula("CartedFlag", "'U'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        ' 6-04-08 Dan M added check box to show salesperson
        If RptSel!ckcSelC10(0).Value = 1 Then
            If Not gSetFormula("HideSalesperson", "false") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("HideSalesperson", "true") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCFrom.Text        'earliest rot start date
        slDate = RptSel!CSI_CalFrom.Text        'earliest rot start date

        If (slDate <> "") Then
            If Not gValidDate(slDate) Then
                'slDateFrom = slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'slSelection = "{CIF_Copy_Inventory.cifRotStartDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
            'Else
                mReset
                RptSel!CSI_CalFrom.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Earliest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCTo.Text          'latest rotation start date
        slDate = RptSel!CSI_CalTo.Text          'latest rotation start date

        If (slDate <> "") Then
            If Not gValidDate(slDate) Then
                'slDateTo = slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'If slSelection = "" Then
                '    slSelection = "{CIF_Copy_Inventory.cifRotStartDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                'Else
                '    slSelection = slSelection & " And {CIF_Copy_Inventory.cifRotStartDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'End If
            'Else
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Latest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If


        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

    '11-22-06 changed to useprepass
    ElseIf ilListIndex = COPY_INVBYEXPDATE Then 'Inventory by Expiration
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCFrom.Text   'earliest rot end date
        slDate = RptSel!CSI_CalFrom.Text   'earliest rot end date
        If (slDate <> "") Then
            If Not gValidDate(slDate) Then
               ' slDateFrom = slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'slSelection = "{CIF_Copy_Inventory.cifRotEndDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
            'Else
                mReset
                RptSel!CSI_CalFrom.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Earliest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCTo.Text 'Latest rotation end date
        slDate = RptSel!CSI_CalTo.Text 'Latest rotation end date
        If (slDate <> "") Then
            If Not gValidDate(slDate) Then
                'slDateTo = slDate
                'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                'If slSelection = "" Then
                '    slSelection = "{CIF_Copy_Inventory.cifRotEndDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                'Else
                '    slSelection = slSelection & " And {CIF_Copy_Inventory.cifRotEndDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                'End If
            'Else
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Latest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

    ElseIf ilListIndex = COPY_INVBYPURGE Then 'Inventory by Purge
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCFrom.Text
        slDate = RptSel!CSI_CalFrom.Text
        If (slDate <> "") Then
            If gValidDate(slDate) Then
                slDateFrom = slDate
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slSelection = "{CIF_Copy_Inventory.cifPurgeDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
            Else
                mReset
                RptSel!CSI_CalFrom.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Earliest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
'        slDate = RptSel!edcSelCTo.Text
        slDate = RptSel!CSI_CalTo.Text
        If (slDate <> "") Then
            If gValidDate(slDate) Then
                slDateTo = slDate
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                If slSelection = "" Then
                    slSelection = "{CIF_Copy_Inventory.cifPurgeDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                Else
                    slSelection = slSelection & " And {CIF_Copy_Inventory.cifPurgeDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If
            Else
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Latest", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        slSelection = slSelection & " And {CIF_Copy_Inventory.cifPurged} = 'P'"
    ElseIf ilListIndex = COPY_INVBYENTRYDATE Then 'Inventory by Entry Date
        If tgSpf.sUseCartNo = "N" Then
            If Not gSetFormula("UseCartNo", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseCartNo", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        
'        slDate = RptSel!edcSelCFrom.Text        'earliset entry date
        slDate = RptSel!CSI_CalFrom.Text        'earliset entry date
        slDateFrom = slDate
        If (slDate <> "") Then
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            slSelection = "{CIF_Copy_Inventory.cifEntryDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
        End If
        
'        slDate = RptSel!edcSelCTo.Text          'latest entry date
        slDate = RptSel!CSI_CalTo.Text          'latest entry date
        slDateTo = slDate
        If slDate <> "" Then
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If slSelection = "" Then
                slSelection = "{CIF_Copy_Inventory.cifEntryDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
            Else
                slSelection = slSelection & " And {CIF_Copy_Inventory.cifEntryDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
            End If
        End If
        
        ilRet = mDateHdr(slDateFrom, slDateTo, "EntryDatesRequested")
'       4-10-13 Implemented and never released; create new report for Inventory Producer instead
'       comment out in case this option needed later.  Option to enter Sent Dates of Inventory Item
'        slDate = RptSel!edcText1.Text       'earlest date sent
'        slDateFrom = slDate
'        If (slDate <> "") Then
'            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'            slSelection = "{CIF_Copy_Inventory.cifInvSentDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
'        End If
'
'
'        slDate = RptSel!edcText2.Text             'latest date sent
'        slDateTo = slDate
'        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'        If slDate <> "" Then
'            If slSelection = "" Then
'                slSelection = "{CIF_Copy_Inventory.cifInvSentDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ") "
'            Else
'                slSelection = slSelection & " And {CIF_Copy_Inventory.cifInvSentDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'            End If
'        End If
'
'        ilRet = mDateHdr(slDateFrom, slDateTo, "DatesSentRequested")
'
        '4-12-05 Show only printables?
        If RptSel!ckcSelC5(0).Value = vbChecked Then
            slSelection = "(" & slSelection & ") and " & "({CIF_Copy_Inventory.cifPrint} <> 'P')"
            If Not gSetFormula("PrintablesOnly", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("PrintablesOnly", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

        '4-13-05 include carted, uncarted, or both when Site TapeShowForm is set to "C" (vs "A" for approved)
        If RptSel!rbcSelC6(0).Value = True Then         'carted
            slSelection = slSelection & " and ({CIF_Copy_Inventory.cifCleared} = 'Y')"
            If Not gSetFormula("CartedFlag", "'C'") Then
                mCopyJob = -1
                Exit Function
            End If
        ElseIf RptSel!rbcSelC6(1).Value = True Then     'uncarted
            slSelection = slSelection & " and ({CIF_Copy_Inventory.cifCleared} <> 'Y')"
            If Not gSetFormula("CartedFlag", "'U'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If

    ElseIf ilListIndex = COPY_INVPRODUCER Then              '4-10-13
'        slDate = RptSel!edcSelCFrom.Text        'earliset entry date
        slDate = RptSel!CSI_CalFrom.Text        'earliset entry date
        slDateFrom = slDate
        'If (slDate <> "") Then
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'End If
        
'        slDate = RptSel!edcSelCTo.Text          'latest entry date
        slDate = RptSel!CSI_CalTo.Text          'latest entry date
        slDateTo = slDate
        'If slDate <> "" Then
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'End If
        ilRet = mDateHdr(slDateFrom, slDateTo, "ActionDatesRequested")
        
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        
    ElseIf ilListIndex = 11 Or ilListIndex = 13 Or ilListIndex = 14 Then    'Play list
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))


        ilRet = mCopyDates(True)        '7-18-00 verify input dates and set formula of dates requested to Crystal
        If ilRet = -1 Then
            mCopyJob = -1
            Exit Function
        End If
        
        If ilListIndex = 11 Then            '7-23-12 playlist by ISCI
            If RptSel!ckcSelC7.Value = vbChecked Then           'include vehicles ordered
                slSelection = "(" & slSelection & ") and " & "({CPR_Copy_Report.cprReady} = 0)"      'records in cpr flagged with cprREady = -1 are for subreport only if vehicle ordered requested
                If Not gSetFormula("InclVehiclesOrdered", "'Y'") Then
                    mCopyJob = -1
                    Exit Function
                End If
            Else
                slSelection = "(" & slSelection & ") and " & "({CPR_Copy_Report.cprReady} = 0)"
                If Not gSetFormula("InclVehiclesOrdered", "'N'") Then
                    mCopyJob = -1
                    Exit Function
                End If
            End If
        End If
        
        '7-18-00 make subrotuine out of following, comment this out
        'slDateFrom = ""
        'slDateTo = ""
        'Date selection passed by formula
        'slDate = RptSel!edcSelCFrom.Text   'Latest cash date
        'slDateFrom = slDate
        'If (slDate <> "") Then
        '    If gValidDate(slDate) Then
        '        slDateFrom = Format$(gDateValue(slDate), "m/d/yy")
        '    Else
        '        mReset
        '        RptSel!edcSelCFrom.SetFocus
        '        Exit Function
        '    End If
        'Else
        '    mReset
        '    RptSel!edcSelCFrom.SetFocus
        '    Exit Function
        'End If
        'slDate = RptSel!edcSelCTo.Text 'Latest billing date
        'slDateTo = slDate
        'If (slDate <> "") Then
        '    If gValidDate(slDate) Then
        '        slDateTo = Format$(gDateValue(slDate), "m/d/yy")
        '    Else
        '        mReset
        '        RptSel!edcSelCTo.SetFocus
        '        Exit Function
        '    End If
        'Else
         '   mReset
         '   RptSel!edcSelCTo.SetFocus
        '    Exit Function
        'End If
        'If gDateValue(slDateFrom) <> gDateValue(slDateTo) Then
        '    If Not gSetFormula("DateRg", "'" & slDateFrom & " - " & slDateTo & "'") Then
        '        mCopyJob = -1
        '        Exit Function
        '    End If
        'Else
        '    If Not gSetFormula("DateRg", "'" & slDateFrom & "'") Then
        '        mCopyJob = -1
        '        Exit Function
        '    End If
        'End If

    ElseIf ilListIndex = COPY_INVUNAPPROVED Then 'unapproved copy
        slDateFrom = ""
        slDateTo = ""
'        slDate = RptSel!edcSelCTo.Text
        slDate = RptSel!CSI_CalTo.Text          '8-23-19 use csi calendar control vs edit box
        If (slDate <> "") Then
            If gValidDate(slDate) Then
                slDateTo = slDate
                gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                slSelection = "{CIF_Copy_Inventory.cifEntryDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
                slSelection = slSelection & " And {CIF_Copy_Inventory.cifCleared} <> 'Y' "
            Else
                mReset
                RptSel!CSI_CalTo.SetFocus
                Exit Function
            End If
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("Entry Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        '2-12-09 moved to splitregionlist
'    ElseIf ilListIndex = COPY_REGIONS Then      '7-18-00
'        ilRet = mCopyDates(False)        '7-18-00 verify input dates and set formula of dates requested to Crystal
'        If RptSel!rbcSelCSelect(0).Value Then       'advertisr or region
'            If Not gSetFormula("SortBy", "'A'") Then
'                mCopyJob = -1
'                Exit Function
'            End If
'        Else
'            If Not gSetFormula("SortBy", "'R'") Then
'                mCopyJob = -1
'                Exit Function
'            End If
'        End If
'
'        If ilRet = -1 Then
'            mCopyJob = -1
'            Exit Function
'        End If
'        slSelection = ""
'        If Not RptSel!ckcAll.Value = vbChecked Then
'            For ilLoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
'                If RptSel!lbcSelection(0).Selected(ilLoop) Then
'                     slNameCode = tgAdvertiser(ilLoop).sKey 'Traffic!lbcAdvertiser.List(ilLoop)
'                     ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
'                     slSelection = slSelection & slOr & "{RAF_Region_Area.rafAdfCode} = " & Trim$(slCode)
'                     slOr = " Or "
'                 End If
'            Next ilLoop
'        End If
'
'        If RptSel!edcSelCFrom.Text = "" Then
'            slDateFrom = "1/1/1970"
'        Else
'            slDateFrom = RptSel!edcSelCFrom.Text
'        End If
'        If RptSel!edcSelCTo.Text = "" Then
'            slDateTo = "12/31/2069"
'        Else
'            slDateTo = RptSel!edcSelCTo.Text
'        End If
'        gObtainYearMonthDayStr slDateFrom, True, slYear, slMonth, slDay
'        slStr = Format$(gDateValue(slDateFrom), "m/d/yy")
'        If slSelection <> "" Then
'            slSelection = slSelection & " and "
'        End If
'        slSelection = slSelection & " ({RAF_Region_Area.rafDateEntrd} >= Date(" & slYear & "," & slMonth & "," & slDay & ")"
'        gObtainYearMonthDayStr slDateTo, True, slYear, slMonth, slDay
'        slSelection = slSelection & " And {RAF_Region_Area.rafDateEntrd} <= Date(" & slYear & "," & slMonth & "," & slDay & "))"
    ElseIf ilListIndex = COPY_BOOK Then
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{ODF_One_Day_Log.odfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({ODF_One_Day_Log.odfGenTime} ) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        slSelection = slSelection & " and  {CIF_Copy_Inventory.cifcsfCode} > 0 and ( {ODF_One_Day_Log.odfZone}[1 to 3] = 'EST' or {ODF_One_Day_Log.odfZone}[1 to 3] = '   ' )  "
    ElseIf ilListIndex = COPY_SPLITROT Then
        ilRet = mCopyDates(False)        '7-18-00 verify input dates and set formula of dates requested to Crystal
        If ilRet = -1 Then
            mCopyJob = -1
            Exit Function
        End If

        If RptSel!rbcSelCSelect(0).Value Then       'split copy
            If Not gSetFormula("RptOption", "'S'") Then
                mCopyJob = -1
                Exit Function
            End If
        ElseIf RptSel!rbcSelCSelect(1).Value Then
            If Not gSetFormula("RptOption", "'B'") Then     'blackouts only
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("RptOption", "'A'") Then     'both split copy and blackout
                mCopyJob = -1
                Exit Function
            End If
        End If

        If RptSel!ckcSelC7.Value = vbChecked Then       'include dormant
            If Not gSetFormula("Dormant", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("Dormant", "'N'") Then     'exclude dormant
                mCopyJob = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    ElseIf ilListIndex = COPY_SCRIPTAFFS Then           '4-9-12
        If RptSel!ckcSelC7.Value = vbChecked Then       'include notarization lines
            If Not gSetFormula("showNotarization", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("showNotarization", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        
        If RptSel!ckcTrans.Value = vbChecked Then       'include inventory detail
            If Not gSetFormula("ShowInvDetail", "'Y'") Then
                mCopyJob = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowInvDetail", "'N'") Then
                mCopyJob = -1
                Exit Function
            End If
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CPR_Copy_Report.cprGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    End If

    'All inventory reports are to ignore History CIF
    'Include  Report                      Index
    ' A & P   Inventory by #                4
    ' A       Inventory by ISCI             5
    ' A       Inventory by Advertiser       6
    ' A       Inventory by Start Date       7
    ' A       Inventory by Expiration Date  8
    '     P   Inventory by Purged Date      9
    ' A       Inventory by Entry Date      10
    If (ilListIndex >= 4) And (ilListIndex <= 8) Or (ilListIndex = 10) Then
        If slSelection = "" Then
            slSelection = "{CIF_Copy_Inventory.cifPurged} <> 'H'"
        Else
            slSelection = slSelection & " And {CIF_Copy_Inventory.cifPurged} <> 'H'"
        End If
        If ilListIndex <> 4 Then
            slSelection = slSelection & " And {CIF_Copy_Inventory.cifPurged} <> 'P'"
        End If
    End If
    If Not gSetSelection(slSelection) Then
        mCopyJob = -1
        Exit Function
    End If
    mCopyJob = 1
    Exit Function

'mCopyJob = 1
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
    RptSel!frcOutput.Enabled = igOutput
    RptSel!frcCopies.Enabled = igCopies
    'RptSel!frcWhen.Enabled = igWhen
    RptSel!frcFile.Enabled = igFile
    RptSel!frcOption.Enabled = igOption
    'rptsel!frcRptType.Enabled = igReportType
    Beep
End Sub
'*****************************************************************************
'
'                       mSelCashTrMercProm - Select receivables
'                           types Cash, Trade, Merchandise, Promotions
'                           Pass Selection formula to Crystal
'                           Slsp Ageing rpt
'
'                       Created: 4/97 D.Hosaka
'
'
'*****************************************************************************
Sub mSelCashTrMercProm(slSelection As String)
'Dim slSelection As String
    If RptSel!rbcSelCInclude(0).Value Then
        If slSelection <> "" Then
            slSelection = slSelection & " And ({RVF_Receivables.rvfCashTrade} = 'C')"
        Else
            slSelection = "{RVF_Receivables.rvfCashTrade} = 'C'"
        End If
    ElseIf RptSel!rbcSelCInclude(1).Value Then
        If slSelection <> "" Then
            slSelection = slSelection & " And ({RVF_Receivables.rvfCashTrade} = 'T')"
        Else
            slSelection = "{RVF_Receivables.rvfCashTrade} = 'T'"
        End If
    ElseIf RptSel!rbcSelCInclude(2).Value Then
        If slSelection <> "" Then
            slSelection = slSelection & " And ({RVF_Receivables.rvfCashTrade} = 'M')"
        Else
            slSelection = "{RVF_Receivables.rvfCashTrade} = 'M'"
        End If
    ElseIf RptSel!rbcSelCInclude(3).Value Then
        If slSelection <> "" Then
            slSelection = slSelection & " And ({RVF_Receivables.rvfCashTrade} = 'P')"
        Else
            slSelection = "{RVF_Receivables.rvfCashTrade} = 'P'"
        End If
    Else
        slSelection = slSelection & " and (Trim({MNF_Multi_Names.mnfCodeStn}) = 'Y')"
    End If
End Sub
'***************************************************************************
'
'                       mTitleCashTrMercProm - send Crystal
'                           formula for report title
'                           on Cash, Trade, Merchandise, or Promotions
'
'                       Return: True if error
'                       Created:  4/97  D.Hosaka
'
'***************************************************************************
Function mTitleCashTrMercProm() As Integer
    mTitleCashTrMercProm = False                     'assume formula ok to send
    If RptSel!rbcSelCInclude(0).Value Then          'cash
        If RptSel!rbcSelC6(0).Value Then            'airtime, ntr or both
            If Not gSetFormula("Name", "'Cash'") Then
                mTitleCashTrMercProm = True
                Exit Function
            End If
        ElseIf RptSel!rbcSelC6(1).Value Then
            If Not gSetFormula("Name", "'NTR'") Then
                    mTitleCashTrMercProm = True
                    Exit Function
            End If
        Else
            If Not gSetFormula("Name", "'Cash & NTR'") Then
                mTitleCashTrMercProm = True
                Exit Function
            End If
        End If
        'Report!crcReport.Formulas(0) = "Name= 'Cash'"
    ElseIf RptSel!rbcSelCInclude(1).Value Then          'trade
        If RptSel!rbcSelC6(0).Value Then            'NTR
            If Not gSetFormula("Name", "'Trade'") Then
                mTitleCashTrMercProm = True
                Exit Function
            End If
        ElseIf RptSel!rbcSelC6(1).Value Then
            If Not gSetFormula("Name", "'NTR'") Then
                    mTitleCashTrMercProm = True
                    Exit Function
            End If
        Else
            If Not gSetFormula("Name", "'Trade & NTR'") Then
                mTitleCashTrMercProm = True
                Exit Function
            End If
        End If
        'Report!crcReport.Formulas(0) = "Name= 'Trade'"
    ElseIf RptSel!rbcSelCInclude(2).Value Then      'merchandising
        If Not gSetFormula("Name", "'Merchandise Usage'") Then
            mTitleCashTrMercProm = True
            Exit Function
        End If
    ElseIf RptSel!rbcSelCInclude(3).Value Then          'promotions
        If Not gSetFormula("Name", "'Promotions Usage'") Then
            mTitleCashTrMercProm = True
            Exit Function
        End If
    Else                                                'hard cost
        If Not gSetFormula("Name", "'Hard Cost Usage'") Then
            mTitleCashTrMercProm = True
            Exit Function
        End If
    End If
End Function
'
'
'           mVerifyDate - verify date entered as valid
'           <input> edcDate - control field (edit box) containing date string
'                   ilDateReqd - true if date required, else false (no date ok)
'           <output> llDate - date converted as Long
'           <return> 0 = OK, 1 = invalid date entered
'
'           3-19-03 added new paramter to test date reqd
Function mVerifyDate(edcDate As Control, llValidDate As Long, ilDateReqd As Integer) As Integer
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
End Function '
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
        ilYear = RptSel!edcSelCTo.Text                'starting year
        ilMonth = RptSel!edcSelCTo1.Text              'month
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
'
'       Format AirTime, NTR or both for report heading
'
'       <input> ilOKToCombine - ok to combine hard cost & non-hard cost
'               ilInclHardCost - true if hard cost included
'               ilInstallMsg as integer - True if using installments and separating Billing & revenue, show that in report header
Function mTitleAirTimeNTRHdr(ilOkToCombine As Integer, ilInclHardCost As Integer, ilInstallMsg As Integer) As Integer
Dim slStatus As String
Dim slInstallMsg As String

    mTitleAirTimeNTRHdr = Correct
    slInstallMsg = ""
    
    If igRptCallType = INVOICESJOB Then
        'test if installment feature turned on and separating billing and revenue
        If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) = INSTALLMENT And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then
            If RptSel!rbcSelC12(0).Value Then           'billing
                slInstallMsg = "Billing for "
            Else
                slInstallMsg = "Revenue for "
            End If
        End If
    End If

    'test for airtime, ntr or both
    If RptSel!rbcSelC6(0).Value Then        'air time only
        slStatus = "Air Time Only"
    ElseIf RptSel!rbcSelC6(1).Value Then    'ntr only
        If ilOkToCombine Then              'OK to combine hard cost with non-hard costs?
            If ilInclHardCost Then         'yes include hard cost
                slStatus = "NTR & Hard Cost Only"
            Else
                slStatus = "NTR Only"
            End If
        Else
            If ilInclHardCost Then
                slStatus = "Hard Cost Only"
            Else
                slStatus = "NTR Only"
            End If
        End If
    Else
        If slInstallMsg = "" Then
            slStatus = ""           '"Air Time & NTR"
        Else
            slStatus = "Air Time & NTR"
        End If
    End If
    If Not gSetFormula("ReportHdr", "'" & slInstallMsg & slStatus & "'") Then
        mTitleAirTimeNTRHdr = Incorrect
        Exit Function
    End If

End Function

Function mTitleTypeAndReceivables(ilOkToCombine As Integer, ilInclHardCost As Integer) As Integer
Dim slStatus As String
'copy of mtitleairtimentrhardcost to add recievable/history to header  6-04-08 Dan M
   mTitleTypeAndReceivables = Correct

    'test for airtime, ntr or both
    If RptSel!rbcSelC6(0).Value Then        'air time only
        slStatus = "Air Time Only"
    ElseIf RptSel!rbcSelC6(1).Value Then    'ntr only
        If ilOkToCombine Then              'OK to combine hard cost with non-hard costs?
            If ilInclHardCost Then         'yes include hard cost
                slStatus = "NTR & Hard Cost Only"
            Else
                slStatus = "NTR Only"
            End If
        Else
            If ilInclHardCost Then
                slStatus = "Hard Cost Only"
            Else
                slStatus = "NTR Only"
            End If
        End If
    Else
        slStatus = "Air Time & NTR"           '"Air Time & NTR"
    End If
    If RptSel!rbcSelC8(2).Value Then
        slStatus = slStatus & " from History & Receivables"
    ElseIf RptSel!rbcSelC8(1).Value Then
        slStatus = slStatus & " from Receivables"
    Else
        slStatus = slStatus & " from History"
    End If
    If Not gSetFormula("ReportHdr", "'" & slStatus & "'") Then
        mTitleTypeAndReceivables = Incorrect
        Exit Function
    End If

End Function
Public Function mVerifyMMYY(InputEdit As Control, slMonth As String, slYear As String) As Integer
Dim slDate As String
Dim ilRet As Integer
Dim slStr As String
Dim ilYear As Integer

        mVerifyMMYY = CP_MSG_NONE
        slDate = InputEdit.Text
        ilRet = gParseItem(slDate, 1, "/", slMonth)
        If ilRet <> CP_MSG_NONE Then
            mReset
            InputEdit.SetFocus
            mVerifyMMYY = CP_MSG_PARSE
            Exit Function
        End If
        If Val(slMonth) < 1 Or Val(slMonth) > 12 Then
            mReset
            InputEdit.SetFocus
            mVerifyMMYY = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slDate, 2, "/", slYear)
        If ilRet <> CP_MSG_NONE Then
            mReset
            InputEdit.SetFocus
            mVerifyMMYY = CP_MSG_PARSE
            Exit Function
        End If
        ilYear = Val(slYear)
        If (ilYear >= 0) And (ilYear <= 69) Then
            ilYear = 2000 + ilYear
        ElseIf (ilYear >= 70) And (ilYear <= 99) Then
            ilYear = 1900 + ilYear
        End If
        If ilYear < 1970 Or ilYear > 2069 Then
            mReset
            InputEdit.SetFocus
            mVerifyMMYY = CP_MSG_PARSE
            Exit Function
        End If

        ilRet = gParseItem(slDate, 3, "/", slStr)
        If ilRet = CP_MSG_NONE Then
            mReset
            InputEdit.SetFocus
            Exit Function
        End If
        slYear = Trim$(str$(ilYear))
End Function
'
'
'   Determine which tran types included
'   for Invoice Register

Public Function mTitleRecHistBoth() As Integer
Dim slStatus As String
Dim ilHowMany As Integer
    mTitleRecHistBoth = 0
    slStatus = ""
    
    'test for Cash / trade
    If RptSel!rbcSelC4(0).Value = True Then     'cash only
        slStatus = "Cash "
    ElseIf RptSel!rbcSelC4(1).Value = True Then
        slStatus = "Trade "
    Else
        slStatus = "Cash/Trade "
    End If
    'test for airtime, ntr or both
    If RptSel!ckcSelC3(0).Value Then        'IN
        slStatus = slStatus & "Invoices"
        ilHowMany = 1
    End If
    If RptSel!ckcSelC3(1).Value Then    'AN
        ilHowMany = ilHowMany + 1
        If slStatus = "" Then
            slStatus = "Adjustments"
        Else
            slStatus = slStatus & "/Adjustments"
        End If
    End If
    If RptSel!ckcSelC3(2).Value Then    'HI
        ilHowMany = ilHowMany + 1
        If slStatus = "" Then
            slStatus = "History"
        Else
            slStatus = slStatus & "/History"
        End If
    End If
    If ilHowMany = 1 Then
        slStatus = slStatus & " Only"
    End If
    If Not gSetFormula("RecHisHdr", "'" & slStatus & "'") Then
        mTitleRecHistBoth = -1
        Exit Function
    End If

End Function

Public Function mTitleCashRpts()
 mTitleCashRpts = False                     'assume formula ok to send
    If RptSel!rbcSelCInclude(0).Value Then          'cash
        If Not gSetFormula("Name", "'Cash'") Then
            mTitleCashRpts = True
            Exit Function
        End If

    ElseIf RptSel!rbcSelCInclude(1).Value Then          'trade
        If Not gSetFormula("Name", "'Trade'") Then
            mTitleCashRpts = True
            Exit Function
        End If

    ElseIf RptSel!rbcSelCInclude(2).Value Then      'merchandising
        If Not gSetFormula("Name", "'Merchandise Usage'") Then
            mTitleCashRpts = True
            Exit Function
        End If
    ElseIf RptSel!rbcSelCInclude(3).Value Then          'promotions
        If Not gSetFormula("Name", "'Promotions Usage'") Then
            mTitleCashRpts = True
            Exit Function
        End If
    End If
End Function

'
'
'       Setup Report header with dates to show "All Dates", "Thru xx/xx/xx", "from xx/xx/xx" or "xx/xx/xx-xx/xx/xx"
Public Function mDateHdr(slFromDate As String, slToDate As String, slFormulaName As String) As Integer
Dim ilError As Integer
Dim slDateFrom As String
Dim slDateTo As String

        ilError = 0
        slDateFrom = slFromDate              'Dates entered, acct hist may be blank
        slDateTo = slToDate
        If slDateFrom <> "" And slDateTo <> "" Then         'Start & end dates entered
            slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
            slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
            If Not gSetFormula(slFormulaName, "'" & slDateFrom & "-" & slDateTo & "'") Then
                ilError = -1
            End If
        ElseIf slDateFrom = "" And slDateTo = "" Then           'no dates entered
            If Not gSetFormula(slFormulaName, "'All Dates'") Then
                ilError = -1
            End If
        ElseIf slDateFrom = "" Then                             'only end date entred
            slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
            If Not gSetFormula(slFormulaName, "'Thru " & slDateTo & "'") Then
                ilError = -1
            End If
        Else                                                    'only start date entred
            slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'makesure year included
            If Not gSetFormula(slFormulaName, "'From " & slDateFrom & "'") Then
                ilError = -1
            End If
        End If
        mDateHdr = ilError
End Function

Function mShowTransComments(ShowTransIn As Control) As Integer
    mShowTransComments = False
    'send formula to show transaction comments
    If ShowTransIn.Value = vbChecked Then    '9-16-03
        If Not gSetFormula("ShowTransComments", "'Y'") Then 'Show transaction commnets
            mShowTransComments = True
            Exit Function
        End If
    Else
        If Not gSetFormula("ShowTransComments", "'N'") Then 'dont show trans comments
            mShowTransComments = True
            Exit Function
        End If
    End If
End Function
'
'           mInpDateFormula - setup the start/end dates entered for formulas sent to crystal for
'                       report headings
'
'           <input> DateFrom - control used to input start date
'                   DateTo - control used to input end date
'                   slFormula Name -formula passed to crystal
'           <return> =-1 if error in sending formula to crystal report

Public Function mInpDateFormula(DateFrom As Control, DateTo As Control, slFormulaName As String) As Integer
Dim slDateFrom As String
Dim slDateTo As String

    mInpDateFormula = 0                     'no error
    slDateFrom = DateFrom.Text                 'Dates entered, acct hist may be blank
    slDateTo = DateTo.Text
    If slDateFrom <> "" And slDateTo <> "" Then         'Start & end dates entered
        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
        slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
        If Not gSetFormula(slFormulaName, "'" & slDateFrom & "-" & slDateTo & "'") Then
            mInpDateFormula = -1
            Exit Function
        End If
    ElseIf slDateFrom = "" And slDateTo = "" Then           'no dates entered
        If Not gSetFormula(slFormulaName, "'All Dates'") Then
            mInpDateFormula = -1
            Exit Function
        End If
    ElseIf slDateFrom = "" Then                             'only end date entred
        slDateTo = Format$(gDateValue(slDateTo), "m/d/yy")    'makesure year included
        If Not gSetFormula(slFormulaName, "'Thru " & slDateTo & "'") Then
            mInpDateFormula = -1
            Exit Function
        End If
    Else                                                    'only start date entred
        slDateFrom = Format$(gDateValue(slDateFrom), "m/d/yy")    'makesure year included
        If Not gSetFormula(slFormulaName, "'From " & slDateFrom & "'") Then
            mInpDateFormula = -1
            Exit Function
        End If
    End If
End Function
'
'       mSetupSpacingForForm - send report the required # of spaces to insert
'       before and after logo to fit address in windowed envelope
'
'       <return> non zero if error in formula
'       6-28-05
'
Public Function mSetupSpacingForm() As Integer
Dim ilBlanksBeforeLogo As Integer
Dim ilBlanksAfterLogo As Integer

        mSetupSpacingForm = 0
        '9-5-03 Send blanks to show in header to align to fit in windowed envelope
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
            mSetupSpacingForm = -1
            Exit Function
        End If
        If Not gSetFormula("BlanksAfterLogo", ilBlanksAfterLogo) Then
            mSetupSpacingForm = -1
            Exit Function
        End If

End Function
'
'
'           obtain dates requested for Post Log reports
'           Send the formula to Crystal reports
'       <return> 0 if ok, 1 if error in sending formula
Private Function mPLDatesFormula() As Integer
Dim slDateFrom As String
Dim slDateTo As String
Dim llDate1 As Long

    mPLDatesFormula = 0
'    slDateFrom = RptSel!edcSelCFrom.Text   'Start date
'    slDateTo = RptSel!edcSelCTo.Text   'End date
    slDateFrom = RptSel!CSI_CalFrom.Text   'Start date
    slDateTo = RptSel!CSI_CalTo.Text   'End date

    llDate1 = gDateValue(slDateFrom)
    slDateFrom = Format$(llDate1, "m/d/yy")
    llDate1 = gDateValue(slDateTo)
    slDateTo = Format$(llDate1, "m/d/yy")
    If Not gSetFormula("RptDates", "'" & slDateFrom & " - " & slDateTo & "'") Then
        mPLDatesFormula = -1
        Exit Function
    End If
    Exit Function
End Function

Private Function mSelAirNTRHardCost(slSelection As String) As Integer
'added 6/03/08 to account for checkboxes instead of radio buttons.
'input, slselection
'output, slselection and integer showing if successful
Dim blMultipleChecks As Boolean
Dim ilFirstChoice As Integer
Dim clCheckBoxObject As CheckBox
Dim ilNumberOfBoxesSelected As Integer
Dim slAirTime As String
Dim slNTR As String
Dim slHardCost As String

blMultipleChecks = False
ilNumberOfBoxesSelected = 0
'find if multiple chosen
For Each clCheckBoxObject In RptSel!ckcSelC6Add
    ilNumberOfBoxesSelected = ilNumberOfBoxesSelected + clCheckBoxObject.Value
    If ilNumberOfBoxesSelected = 2 Then
        blMultipleChecks = True
        Exit For
    End If
Next clCheckBoxObject
If ilNumberOfBoxesSelected = 0 Then 'not really necessary. command button only enabled when at least one chosen.
    mSelAirNTRHardCost = Incorrect
    Exit Function
End If
'find first chosen
ilFirstChoice = HardCost
If RptSel!ckcSelC6Add(Airtime) Then
    ilFirstChoice = Airtime
Else
    If RptSel!ckcSelC6Add(NTR) Then
        ilFirstChoice = NTR
    End If
End If
slAirTime = "( {RVF_Receivables.rvfmnfItem} = 0)"
slNTR = "( {RVF_Receivables.rvfmnfItem} > 0 and Trim({MNF_Multi_Names.mnfCodeStn}) <> 'Y')"
slHardCost = "( {RVF_Receivables.rvfmnfItem} > 0 and Trim({MNF_Multi_Names.mnfCodeStn}) = 'Y')"
 Select Case ilFirstChoice
    Case Airtime
        If blMultipleChecks Then
            slSelection = slSelection & " AND ( " & slAirTime   'airtime +
            If RptSel!ckcSelC6Add(NTR) Then
                slSelection = slSelection & " OR  " & slNTR    'airtime + NTR
                If RptSel!ckcSelC6Add(HardCost) Then
                    slSelection = slSelection & " OR  " & slHardCost & ")" 'airtime +NTR + HardCost
                Else
                    slSelection = slSelection & ")" 'airtime + NTR
                End If
            Else
                slSelection = slSelection & " OR  " & slHardCost & ")" 'airtime + HardCost
            End If
        Else
            slSelection = slSelection & " AND " & slAirTime     'airtime only
        End If
    Case NTR
        If blMultipleChecks Then
            slSelection = slSelection & " AND ( " & slNTR & " OR " & slHardCost & ")"   'NTR + HardCost
        Else
            slSelection = slSelection & " AND " & slNTR 'NTR only
        End If
    Case Else
             slSelection = slSelection & " AND " & slHardCost  'HardCost only
End Select
mSelAirNTRHardCost = Correct
End Function

Private Function mTitleCTMPForCbc() As Integer
' replaced mTitleCashTrMercProm for check boxes instead of radio buttons  6-03-08  Dan M

Dim slvalue As String
mTitleCTMPForCbc = 0                     'assume formula ok to send

' find which checkboxes checked. Combine with radiobuttons cInclude if airtime selected
If RptSel!rbcSelCInclude(0).Value Then  'cash
    slvalue = "'Cash "
ElseIf RptSel!rbcSelCInclude(1).Value Then
    slvalue = "'Trade "
ElseIf RptSel!rbcSelCInclude(2).Value Then      'merchandising
            slvalue = "'Merchandising "
Else          'promotions
                slvalue = slvalue & "'Promotions "
End If
If (RptSel!ckcSelC6Add(Airtime).Value = 1) And (RptSel!ckcSelC6Add(NTR).Value = 1) And (RptSel!ckcSelC6Add(HardCost).Value = 1) Then
    slvalue = slvalue & "Air Time, NTR, & Hard Cost"
ElseIf RptSel!ckcSelC6Add(Airtime).Value = 1 Then
        slvalue = slvalue & "Air Time"
    If RptSel!ckcSelC6Add(NTR).Value = 1 Then
        slvalue = slvalue & " & NTR"
    ElseIf RptSel!ckcSelC6Add(HardCost).Value = 1 Then
        slvalue = slvalue & " & Hard Cost"
    End If
ElseIf RptSel!ckcSelC6Add(NTR).Value = 1 Then
        slvalue = slvalue & "NTR"
        If RptSel!ckcSelC6Add(HardCost).Value = 1 Then
            slvalue = slvalue & " & Hard Cost"
        End If
Else
        slvalue = slvalue & "Hard Cost"
End If

slvalue = slvalue & "'"
If Not gSetFormula("Name", slvalue) Then
   mTitleCTMPForCbc = -1
End If

'Date: 2020-02-26 added suppress zero balance option
If RptSel!ckcSuppressZB.Value = 1 Then
    'TTP 11016 - Aging by Salespeople Report - Application Error #91 when "Salespeople Splits" is checked on
    'If Not gSetFormula("ZeroBalance", "'*Zero Balance transactions excluded'") Then
    If Not gSetFormula("ZeroBalance", "'*Zero Balance transactions excluded'", False) Then
        mTitleCTMPForCbc = -1
    End If
End If

End Function
'
'               mTitlePoliticals - setup formula to pass to crystal reports indicating
'               if Politicals, non-politicals included
Public Function mTitlePoliticals(ilPolitical As Control, ilNonPolitical As Control) As Integer
Dim slStr As String

    mTitlePoliticals = 0
    If ilPolitical = vbChecked And ilNonPolitical = vbChecked Then
        slStr = "Polit & Non-Polit"
    ElseIf ilPolitical = vbChecked Then
        slStr = "Polit"
    Else
        slStr = "Non-Polit"
    End If
    If Not gSetFormula("PolitHdr", "'" & slStr & "'") Then
       mTitlePoliticals = -1
    End If
Exit Function

End Function

Public Function mExtendedAgeing(slBaseDate) As Integer
Dim slDate As String
Dim illoop As Integer
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim ilExtendedLoop As Integer
Dim ilTempExtended As Integer

        mExtendedAgeing = 0     'valid date conversions
        'Set 6 previous end date values from base date
        'Extended month columns are 1 month (current), P2 = 1 month, P3 = 1month,
        'P4 = 1 month, P5 = 2 months, P6 = 6 months, P7 = 1 yr and prior
        If tgSpf.sRRP = "C" Then    'Calendar
            slDate = slBaseDate
            slDate = gObtainStartCal(slDate)
            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month

            For illoop = 2 To 7 Step 1
                If illoop = 5 Then
                    ilExtendedLoop = 1
                ElseIf illoop = 6 Then
                    ilExtendedLoop = 2
                ElseIf illoop = 7 Then
                    ilExtendedLoop = 6
                Else
                    ilExtendedLoop = 1
                End If
                For ilTempExtended = 1 To ilExtendedLoop   'loop as many times as required for the ageing column
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slDate = gObtainStartCal(slDate)
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                Next ilTempExtended
                If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mExtendedAgeing = -1
                    Exit Function
                End If
            Next illoop
        ElseIf tgSpf.sRRP = "F" Then 'Corporate
            slDate = slBaseDate
            slDate = gObtainStartCorp(slDate, True)
            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month

            For illoop = 2 To 7 Step 1
                If illoop = 5 Then
                    ilExtendedLoop = 1
                ElseIf illoop = 6 Then
                    ilExtendedLoop = 2
                ElseIf illoop = 7 Then
                    ilExtendedLoop = 6
                Else
                    ilExtendedLoop = 1
                End If
                For ilTempExtended = 1 To ilExtendedLoop   'loop as many times as required for the ageing column
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slDate = gObtainStartCorp(slDate, True)
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                Next ilTempExtended
                If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mExtendedAgeing = -1
                    Exit Function
                End If
            Next illoop

        Else    'Standard
            slDate = slBaseDate
            slDate = gObtainStartStd(slDate)
            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month

            For illoop = 2 To 7 Step 1
                If illoop = 5 Then
                    ilExtendedLoop = 1
                ElseIf illoop = 6 Then
                    ilExtendedLoop = 2
                ElseIf illoop = 7 Then
                    ilExtendedLoop = 6
                Else
                    ilExtendedLoop = 1
                End If
                For ilTempExtended = 1 To ilExtendedLoop   'loop as many times as required for the ageing column
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slDate = gObtainStartStd(slDate)
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")   'Start date next month
                Next ilTempExtended
                If Not gSetFormula("P" & Trim$(str$(illoop)), "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mExtendedAgeing = -1
                    Exit Function
                End If
            Next illoop
        End If
End Function
'
'               determine the sort listindex from list boxes for User Activity log
'               <input> cbcList - list box control
'                       ilIncludeNone - use none as a choice
'               return - cbcList - list index selected
Public Sub mUserActivitySortSelect(cbcList As Control, ilIncludeNone As Integer, slCode As String)
Dim ilListIndex As Integer

        ilListIndex = cbcList.ListIndex
        If Not ilIncludeNone Then
            ilListIndex = ilListIndex + 1
        End If
        If ilListIndex = 0 Then
            slCode = ""
        ElseIf ilListIndex = 1 Then
            slCode = "A"            'activity
        ElseIf ilListIndex = 2 Then
            slCode = "D"            'date
        ElseIf ilListIndex = 3 Then
            slCode = "S"            'system type
        ElseIf ilListIndex = 4 Then 'time
            slCode = "T"
        ElseIf ilListIndex = 5 Then     'user
            slCode = "U"
        End If
        
        Exit Sub
   
End Sub
'
'             mSendAsOfDateFormula - send formula "AsOfDate" to crystal report'
'           <input> slDate  - date string
'           return 0 = formula ok
'                  -1  Error
Public Function mSendAsOfDateFormula(slDate As String) As Integer
Dim slYear As String
Dim slMonth As String
Dim slDay As String

            mSendAsOfDateFormula = 0
            
            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("As Of Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                mSendAsOfDateFormula = -1
                Exit Function
            End If
            
            Exit Function
End Function
