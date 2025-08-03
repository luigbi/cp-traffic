Attribute VB_Name = "RPTVFYSPOTBB"


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfySpotBB.Bas  Spot Business Booked
'
Option Explicit
Option Compare Text

'If adding or changing order of sort/selection list boxes, change these constants and also
'see rptcrspotbb for any further tests.
Const SORT_ADVT = 1
Const SORT_AGY = 2
Const SORT_BUSCAT = 3
Const SORT_PRODPROT = 4
Const SORT_SLSP = 5
Const SORT_VEHICLE = 6
Const SORT_VG = 7

'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportSpotBB                *
'*      Spot Business Booked
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenSpotBB() As Integer
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


        gCmcGenSpotBB = 0
        
        slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
                
        slExclude = ""
        slInclude = ""
        
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(1), slInclude, slExclude, "Orders"
        'gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(2), slInclude, slExclude, "Net"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(3), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(4), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(5), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(6), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(7), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(8), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(9), slInclude, slExclude, "Promo"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(21), slInclude, slExclude, "Cash"           '8-7-19
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(10), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(11), slInclude, slExclude, "AirTime"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(12), slInclude, slExclude, "Rep"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(13), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(14), slInclude, slExclude, "H/C"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(15), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(16), slInclude, slExclude, "Non-Polit"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(17), slInclude, slExclude, "Missed"
        gIncludeExcludeCkc RptSelSpotBB!ckcAllTypes(18), slInclude, slExclude, "Cancel"
        gIncludeExcludeCkc RptSelSpotBB!ckcAdj(0), slInclude, slExclude, "Rep Adj"
        gIncludeExcludeCkc RptSelSpotBB!ckcAdj(1), slInclude, slExclude, "A/T Adj"
        
        If (RptSelSpotBB!cbcSort1.ListIndex) + 1 = SORT_SLSP Or (RptSelSpotBB!cbcSort2.ListIndex) = SORT_SLSP Or (RptSelSpotBB!cbcSort3.ListIndex) = SORT_SLSP Then
            gIncludeExcludeCkc RptSelSpotBB!ckcShowSlspSplit, slInclude, slExclude, "SlspSplit"
        End If
    
        'only show inclusions
        If Len(slInclude) > 0 Then
            slInclude = "Include: " & slInclude
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        If Len(slExclude) <= 0 Then
            slExclude = "Exclude: None"
        Else
            slExclude = "Exclude: " & slExclude
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        ilSort = (RptSelSpotBB!cbcSort1.ListIndex) + 1      '0 will indicate none for other 2 sorts
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort1", "'" & slSortCode & "'") Then
            gCmcGenSpotBB = -1
            Exit Function
        End If

        ilSort = RptSelSpotBB!cbcSort2.ListIndex
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort2", "'" & slSortCode & "'") Then
            gCmcGenSpotBB = -1
            Exit Function
        End If
        
        ilSort = RptSelSpotBB!cbcSort3.ListIndex
        slSortCode = mConvertIndexToCode(ilSort)
        If Not gSetFormula("UserSort3", "'" & slSortCode & "'") Then
            gCmcGenSpotBB = -1
            Exit Function
        End If
        
        If (RptSelSpotBB!cbcSortVG.ListIndex > 0) Then          'any vehicle groups selected?
            If Not gSetFormula("UseVehicleGroup", "'Y'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("UseVehicleGroup", "'N'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        If RptSelSpotBB!ckcSkip1.Value = vbChecked Then
            If Not gSetFormula("SkipSort1", "'Y'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort1", "'N'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If

        If RptSelSpotBB!ckcSkip2.Value = vbChecked Then
            If Not gSetFormula("SkipSort2", "'Y'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort2", "'N'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        If RptSelSpotBB!ckcSkip3.Value = vbChecked Then
            If Not gSetFormula("SkipSort3", "'Y'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSort3", "'N'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        If RptSelSpotBB!ckcSkipVG.Value = vbChecked Then
            If Not gSetFormula("SkipSortVG", "'Y'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("SkipSortVG", "'N'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        If RptSelSpotBB!rbcPerType(0).Value = True Then            'weekly, use date
            slDate = RptSelSpotBB!edcStart.Text        'date entered, backup to Monday
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
                gCmcGenSpotBB = -1
                Exit Function
            End If

        Else                                                'use month
            If RptSelSpotBB!rbcPerType(2).Value = True Then     'corporate
                slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year
                gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)
                If Not gSetFormula("StartingMonth", ilSaveMonth) Then
                    gCmcGenSpotBB = -1
                    Exit Function
                End If
                
                ilYear = gGetYearofCorpMonth(igMonthOrQtr, igYear)
                If Not gSetFormula("StartingYear", ilYear) Then
                    gCmcGenSpotBB = -1
                    Exit Function
                End If
                'adjusted actual starting month & year of the requested corporate start date to gather
                igYear = ilYear
                igMonthOrQtr = ilSaveMonth
            Else
                'igMonthOrQtr has starting Month
                If Not gSetFormula("StartingMonth", igMonthOrQtr) Then
                    gCmcGenSpotBB = -1
                    Exit Function
                End If
                 'igYear has starting year
                If Not gSetFormula("StartingYear", igYear) Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
            End If
            
        End If

        ilSort = Val(RptSelSpotBB!edcPeriods.Text)
        If Not gSetFormula("NumberPeriods", ilSort) Then
            gCmcGenSpotBB = -1
            Exit Function
        End If
        
        If RptSelSpotBB!rbcTotalsBy(0).Value = True Then        'detail
            If Not gSetFormula("TotalsBy", "'D'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        ElseIf RptSelSpotBB!rbcTotalsBy(1).Value = True Then        'advt totals
            If Not gSetFormula("TotalsBy", "'A'") Then
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("TotalsBy", "'S'") Then      'Summary
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        If RptSelSpotBB!rbcPerType(0).Value = True Then
            If Not gSetFormula("PeriodType", "'W'") Then      'Weekly
                gCmcGenSpotBB = -1
                Exit Function
            End If
        ElseIf RptSelSpotBB!rbcPerType(1).Value = True Then
            If Not gSetFormula("PeriodType", "'S'") Then      'Standard
                gCmcGenSpotBB = -1
                Exit Function
            End If
        ElseIf RptSelSpotBB!rbcPerType(2).Value = True Then
            If Not gSetFormula("PeriodType", "'O'") Then      'Corporate
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("PeriodType", "'C'") Then      'Calendar
                gCmcGenSpotBB = -1
                Exit Function
            End If

        End If
        
        If RptSelSpotBB!rbcGrossNet(0).Value = True Then        'gross
            If Not gSetFormula("SpotOrRev", "'G'") Then      '
                gCmcGenSpotBB = -1
                Exit Function
            End If
        ElseIf RptSelSpotBB!rbcGrossNet(1).Value = True Then        'Net
            If Not gSetFormula("SpotOrRev", "'N'") Then      '
                gCmcGenSpotBB = -1
                Exit Function
            End If
        Else                                                        'spot count
            If Not gSetFormula("SpotOrRev", "'S'") Then      '
                gCmcGenSpotBB = -1
                Exit Function
            End If
        End If
        
        'use sales source as major only applies to the Spot Revenue Register report
        If RptSelSpotBB!lbcRptType.ListIndex = SPOTREV_REGISTER Then
            If RptSelSpotBB!ckcUseSS.Value = vbChecked Then
                If Not gSetFormula("UseSSMajor", "'Y'") Then      '
                    gCmcGenSpotBB = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("UseSSMajor", "'n'") Then      '
                    gCmcGenSpotBB = -1
                    Exit Function
                End If
            End If
        End If
        
        '12-2-16 Disclaimer with NTR
        slStr = ""
        If (RptSelSpotBB!ckcAllTypes(13).Value = vbChecked) Or (RptSelSpotBB!ckcAllTypes(14).Value = vbChecked) Or (RptSelSpotBB!ckcAdj(0).Value = vbChecked) Or (RptSelSpotBB!ckcAdj(1).Value = vbChecked) Then
            slStr = " Inclusion of NTR does not test vehicle type, whether rep or non-rep; Inclusion of NTR adjustments test vehicle type against rep or non-rep selectivity"
        End If
        If Not gSetFormula("Disclaimer", "'" & slStr & "'") Then
            gCmcGenSpotBB = -1
            Exit Function
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
        If Not gSetSelection(slSelection) Then
            gCmcGenSpotBB = 0
            Exit Function
        End If
            
        gCmcGenSpotBB = 1         'ok
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportSpotBB                *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Validity checking of input &   *
'                               Open the Crystal report                *
'
'*******************************************************
Function gGenReportSpotBB() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer

        ilRet = mVerifyMonthYrPeriods()       'validity check year, month, # periods or start date & #periods Input
        If ilRet = 0 Then
            If RptSelSpotBB!lbcRptType.ListIndex = REV_ON_BOOKS Then
                If Not gOpenPrtJob("RevOnBooks.Rpt") Then
                    gGenReportSpotBB = False
                    Exit Function
                End If
            Else
                If Not gOpenPrtJob("SpotRevReg.Rpt") Then
                    gGenReportSpotBB = False
                    Exit Function
                End If
            End If
        Else
            gGenReportSpotBB = False
            Exit Function
        End If

    gGenReportSpotBB = True
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
    RptSelSpotBB!frcOutput.Enabled = igOutput
    RptSelSpotBB!frcCopies.Enabled = igCopies
    'RptSelSpotBB!frcWhen.Enabled = igWhen
    RptSelSpotBB!frcFile.Enabled = igFile
    RptSelSpotBB!frcOption.Enabled = igOption
    'RptSelSpotBB!frcRptType.Enabled = igReportType
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

        If RptSelSpotBB!rbcPerType(0).Value = True Then     'weekly
            RptSelSpotBB!edcStart.Text = RptSelSpotBB!calStart.Text
            'check for valid start
            slDateFrom = RptSelSpotBB!edcStart.Text
            If Not gValidDate(slDateFrom) Then
                mReset
                RptSelSpotBB!edcStart.SetFocus
                Exit Function
            End If
            llDate = gDateValue(slDateFrom)
            slDateFrom = Format$(llDate, "m/d/yy")
            ilHiLimit = 14
            slStr = RptSelSpotBB!edcPeriods.Text            '#weeks
        Else                                                'monthly
            slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
            mVerifyMonthYrPeriods = 0
            'verify user input dates
            slStr = RptSelSpotBB!edcStart.Text
            igYear = gVerifyYear(slStr)
            If igYear = 0 Then
                mReset
                RptSelSpotBB!edcStart.SetFocus                 'invalid year
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If
            slStr = RptSelSpotBB!edcMonth.Text                 'month in text form (jan..dec), or just a month # could have been entered
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                ilSaveMonth = Val(slStr)
            Else
                slStr = str$(ilSaveMonth)
            End If

            ilRet = gVerifyInt(slStr, 1, 12)        'if month number came back 0, its invalid
            If ilRet = -1 Then
                mReset
                RptSelSpotBB!edcMonth.SetFocus                 'invalid month #
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If

            
            If RptSelSpotBB!rbcPerType(2).Value = True Then          'corporate
                 'convert the month name to the correct relative month # of the corp calendar
                'i.e. if 10 entered and corp calendar starts with oct, the result will be july (10th month of corp cal)
                ilYear = gGetCorpCalIndex(igYear)
                If ilYear <= 0 Then                  'year not defined
                    MsgBox "Corporate Year not Defined"
                    mVerifyMonthYrPeriods = -1          'invalid corporate calendar
                    Exit Function
                End If
                slStr = RptSelSpotBB!edcMonth.Text                 'month in text form (jan..dec), or just a month # could have been entered
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
           
            slStr = RptSelSpotBB!edcPeriods.Text               '#periods
            ilHiLimit = 12

        End If

        ilRet = gVerifyInt(slStr, 1, ilHiLimit)             '#periods  validity check
        If ilRet = -1 Then
            mReset
            RptSelSpotBB!edcPeriods.SetFocus                 'invalid #periods
            mVerifyMonthYrPeriods = -1
            Exit Function
        End If

        
    Exit Function
End Function
'
'                   mConvertIndexToCode - convert the index number of sort selection
'                   to a alpha code to send to crystal
'                   <input> index to selection
'                   <return>  1 char code indicating the sort selected
Private Function mConvertIndexToCode(ilIndex As Integer) As String
Dim slChar As String * 1
        
        slChar = "N"            'assume NONE selected
        If ilIndex = SORT_ADVT Then
            slChar = "A"
        ElseIf ilIndex = SORT_AGY Then
            slChar = "G"
        ElseIf ilIndex = SORT_BUSCAT Then
            slChar = "B"
        ElseIf ilIndex = SORT_PRODPROT Then
            slChar = "P"
        ElseIf ilIndex = SORT_SLSP Then
            slChar = "S"
        ElseIf ilIndex = SORT_VEHICLE Then
            slChar = "V"
        End If
        mConvertIndexToCode = slChar
        Exit Function
End Function
