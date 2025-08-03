Attribute VB_Name = "RPTVFYBO"


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyBO.Bas  Sales Breakout
'
Option Explicit
Option Compare Text


'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportSalesBreakout         *
'*      Spot Business Booked
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenSalesBreakout() As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
Dim slDateFrom As String
Dim slDateTo As String
Dim slDate As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
Dim slTime As String
Dim slSelection As String
Dim slInclude As String
Dim slExclude As String
Dim llDate As Long
Dim ilSort As Integer
Dim slUseVG As String * 1
Dim slSortCode As String * 1
Dim ilPerStartDate(0 To 1) As Integer
Dim ilSaveMonth As Integer
Dim ilDay As Integer
Dim ilYear As Integer
Dim slMonthInYear As String * 36
Dim slStr As String


        gCmcGenSalesBreakout = 0
        
        slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
                
        slExclude = ""
        slInclude = ""
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(1), slInclude, slExclude, "Orders"
        'gIncludeExcludeCkc RptSelBO!ckcAllTypes(2), slInclude, slExclude, "Net"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(3), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(4), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(5), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(6), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(7), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(8), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(9), slInclude, slExclude, "Promo"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(10), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(11), slInclude, slExclude, "AirTime"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(12), slInclude, slExclude, "Rep"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(13), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(14), slInclude, slExclude, "H/C"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(15), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(16), slInclude, slExclude, "Non-Polit"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(17), slInclude, slExclude, "Missed"
        gIncludeExcludeCkc RptSelBO!ckcAllTypes(18), slInclude, slExclude, "Cancel"
        gIncludeExcludeCkc RptSelBO!ckcAdj(0), slInclude, slExclude, "Rep Adj"
        gIncludeExcludeCkc RptSelBO!ckcAdj(1), slInclude, slExclude, "A/T Adj"
        'only show inclusions
        If Len(slInclude) > 0 Then
            slInclude = "Include: " & slInclude
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        End If
        If Len(slExclude) <= 0 Then
            slExclude = "Exclude: None"
        Else
            slExclude = "Exclude: " & slExclude
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        End If
                
        If RptSelBO!rbcPerType(0).Value = True Then            'weekly, Unused for now
            slDate = RptSelBO!edcStart.Text        'date entered, backup to Monday
            llDate = gDateValue(slDate)
            'backup to Monday
            ilDay = gWeekDayLong(llDate)
            Do While ilDay <> 0
                llDate = llDate - 1
                ilDay = gWeekDayLong(llDate)
            Loop
            slDate = Format$(llDate, "m/d/yy")

            gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
            If Not gSetFormula("UserEffecDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If

        Else                                                'use month
            If RptSelBO!rbcPerType(2).Value = True Then     'corporate
                slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year
                gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)
                If Not gSetFormula("StartingMonth", ilSaveMonth) Then
                    gCmcGenSalesBreakout = -1
                    Exit Function
                End If
                
                ilYear = gGetYearofCorpMonth(igMonthOrQtr, igYear)
                If Not gSetFormula("StartingYear", ilYear) Then
                    gCmcGenSalesBreakout = -1
                    Exit Function
                End If
                'adjusted actual starting month & year of the requested corporate start date to gather
                igYear = ilYear
                igMonthOrQtr = ilSaveMonth
            Else
                'igMonthOrQtr has starting Month
                If Not gSetFormula("StartingMonth", igMonthOrQtr) Then
                    gCmcGenSalesBreakout = -1
                    Exit Function
                End If
                 'igYear has starting year
                If Not gSetFormula("StartingYear", igYear) Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
            End If
            
        End If

        ilSort = Val(RptSelBO!edcPeriods.Text)
        If Not gSetFormula("NumberPeriods", ilSort) Then
            gCmcGenSalesBreakout = -1
            Exit Function
        End If
        
        If RptSelBO!rbcTotalsBy(0).Value = True Then        'detail
            If Not gSetFormula("TotalsBy", "'D'") Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        ElseIf RptSelBO!rbcTotalsBy(1).Value = True Then        'advt totals
            If Not gSetFormula("TotalsBy", "'A'") Then
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("TotalsBy", "'S'") Then      'Summary
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        End If
        
        If RptSelBO!rbcPerType(0).Value = True Then
            If Not gSetFormula("PeriodType", "'W'") Then      'Weekly
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        ElseIf RptSelBO!rbcPerType(1).Value = True Then
            If Not gSetFormula("PeriodType", "'S'") Then      'Standard
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        ElseIf RptSelBO!rbcPerType(2).Value = True Then
            If Not gSetFormula("PeriodType", "'O'") Then      'Corporate
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("PeriodType", "'C'") Then      'Calendar
                gCmcGenSalesBreakout = -1
                Exit Function
            End If

        End If
        
        If RptSelBO!rbcGrossNet(0).Value = True Then        'gross
            If Not gSetFormula("SpotOrRev", "'G'") Then      '
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        ElseIf RptSelBO!rbcGrossNet(1).Value = True Then        'Net
            If Not gSetFormula("SpotOrRev", "'N'") Then      '
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        Else                                                        'spot count (not implemented)
            If Not gSetFormula("SpotOrRev", "'S'") Then      '
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        End If
        
        
        If RptSelBO!rbcVehicleTotals(0).Value = True Then        'Combine vehicle totals
            If Not gSetFormula("CombineOrSeparateVehicle", "'C'") Then      '
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("CombineOrSeparateVehicle", "'S'") Then             'separate vehicle totals
                gCmcGenSalesBreakout = -1
                Exit Function
            End If
        End If

        '12-2-16 Disclaimer with NTR
        slStr = ""
        If (RptSelBO!ckcAllTypes(13).Value = vbChecked) Or (RptSelBO!ckcAllTypes(14).Value = vbChecked) Or (RptSelBO!ckcAdj(0).Value = vbChecked) Or (RptSelBO!ckcAdj(1).Value = vbChecked) Then
            slStr = " Inclusion of NTR does not test vehicle type, whether rep or non-rep; Inclusion of NTR adjustments test vehicle type against rep or non-rep selectivity"
        End If
        If Not gSetFormula("Disclaimer", "'" & slStr & "'") Then
            gCmcGenSalesBreakout = -1
            Exit Function
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
        If Not gSetSelection(slSelection) Then
            gCmcGenSalesBreakout = 0
            Exit Function
        End If
            
        gCmcGenSalesBreakout = 1         'ok
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportSalesBreakout         *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Validity checking of input &   *
'                               Open the Crystal report                *
'
'*******************************************************
Function gGenReportSalesBreakout() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer

        ilRet = mVerifyMonthYrPeriods()       'validity check year, month, # periods or start date & #periods Input
        If ilRet = 0 Then
            If Not gOpenPrtJob("SalesBO.Rpt") Then
                gGenReportSalesBreakout = False
                Exit Function
            End If
        Else
            gGenReportSalesBreakout = False
            Exit Function
        End If

    gGenReportSalesBreakout = True
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
    RptSelBO!frcOutput.Enabled = igOutput
    RptSelBO!frcCopies.Enabled = igCopies
    'RptSelBO!frcWhen.Enabled = igWhen
    RptSelBO!frcFile.Enabled = igFile
    RptSelBO!frcOption.Enabled = igOption
    'RptSelBO!frcRptType.Enabled = igReportType
    Beep
End Sub

'        verify input parameters for Year, Month and # periods
'        <input> None
'         return - 0: Valid input, -1 : error in input
Public Function mVerifyMonthYrPeriods() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slMonth                                                                               *
'******************************************************************************************

Dim slStr As String
Dim ilSaveMonth As Integer
Dim ilRet As Integer
Dim slMonthInYear As String * 36
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilYear As Integer

        If RptSelBO!rbcPerType(0).Value = True Then     'weekly (not implemented for now)
            RptSelBO!edcStart.Text = RptSelBO!calStart.Text
            'check for valid start
            slDateFrom = RptSelBO!edcStart.Text
            If Not gValidDate(slDateFrom) Then
                mReset
                RptSelBO!edcStart.SetFocus
                Exit Function
            End If
            llDate = gDateValue(slDateFrom)
            slDateFrom = Format$(llDate, "m/d/yy")
            ilHiLimit = 14
            slStr = RptSelBO!edcPeriods.Text            '#weeks
        Else                                                'monthly
            slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
            mVerifyMonthYrPeriods = 0
            'verify user input dates
            slStr = RptSelBO!edcStart.Text
            igYear = gVerifyYear(slStr)
            If igYear = 0 Then
                mReset
                RptSelBO!edcStart.SetFocus                 'invalid year
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If
            slStr = RptSelBO!edcMonth.Text                 'month in text form (jan..dec), or just a month # could have been entered
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                ilSaveMonth = Val(slStr)
            Else
                slStr = str$(ilSaveMonth)
            End If

            ilRet = gVerifyInt(slStr, 1, 12)        'if month number came back 0, its invalid
            If ilRet = -1 Then
                mReset
                RptSelBO!edcMonth.SetFocus                 'invalid month #
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If

            
            If RptSelBO!rbcPerType(2).Value = True Then          'corporate
                 'convert the month name to the correct relative month # of the corp calendar
                'i.e. if 10 entered and corp calendar starts with oct, the result will be july (10th month of corp cal)
                ilYear = gGetCorpCalIndex(igYear)
                If ilYear <= 0 Then                  'year not defined
                    MsgBox "Corporate Year not Defined"
                    mVerifyMonthYrPeriods = -1          'invalid corporate calendar
                    Exit Function
                End If
                slStr = RptSelBO!edcMonth.Text                 'month in text form (jan..dec), or just a month # could have been entered
                gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
                If ilSaveMonth <> 0 Then                           'input is text month name,
                    slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                    igMonthOrQtr = gGetCorpMonthNoFromMonthName(slMonthInYear, slStr)         'getmonth # relative to start of corp cal
                Else
                    igMonthOrQtr = Val(slStr)
                End If
            Else                                        'std or calendar
                igMonthOrQtr = Val(slStr)
            End If
           
            slStr = RptSelBO!edcPeriods.Text               '#periods
            ilHiLimit = 12

        End If

        ilRet = gVerifyInt(slStr, 1, ilHiLimit)             '#periods  validity check
        If ilRet = -1 Then
            mReset
            RptSelBO!edcPeriods.SetFocus                 'invalid #periods
            mVerifyMonthYrPeriods = -1
            Exit Function
        End If

        
    Exit Function
End Function

