Attribute VB_Name = "RPTVFY30"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfy30.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSel30.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

'
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmcGen30                       *
'*                                                     *
'*      Created:6/12/13
'*                                                     *
'*      Comments: Formula setups for Crystal           *
'*      CPP/CPM by 30" Unit
'*                                                     *
'*******************************************************
Function gCmcGen30(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
    Dim slSelection As String
    Dim slNameCode As String
    Dim slInclude As String
    Dim slExclude As String
    Dim slMonthYear As String
    Dim ilSaveMonth As Integer
    Dim ilYear As Integer
    Dim slMonthInYear As String * 36
    Dim ilTemp As Integer
    Dim slSpotLenRatio As String
    Dim slLen(0 To 9) As String
    Dim slIndex(0 To 9) As String

    Dim tlSpotLenRatio As SPOTLENRATIO
        
   slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
    
    gCmcGen30 = 0

    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSel30!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSel30!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSel30!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSel30!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSel30!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSel30!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSel30!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSel30!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSel30!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSel30!ckcCType(10), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSel30!ckcCType(2), slInclude, slExclude, "Polit"
    gIncludeExcludeCkc RptSel30!ckcCType(11), slInclude, slExclude, "Non-Polit"
    
    gIncludeExcludeCkc RptSel30!ckcSpotType(0), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSel30!ckcSpotType(1), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSel30!ckcSpotType(2), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSel30!ckcSpotType(3), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSel30!ckcSpotType(4), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSel30!ckcSpotType(6), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSel30!ckcSpotType(7), slInclude, slExclude, "Spinoff"
    gIncludeExcludeCkc RptSel30!ckcSpotType(5), slInclude, slExclude, "MG"
    'ckcSpotType(8-11) unused
    
    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'Included: " & slInclude & "'") Then
            gCmcGen30 = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'Excluded: " & slExclude & "'") Then
            gCmcGen30 = -1
            Exit Function
        End If
    End If
    
    
    'RptHeader
    'CorpStd
    'Periods
    'demo   (use in one of the header for demo and book:  Schedule line book for prmary demo; sched line book for M18-34)
  
    If RptSel30!rbcByCPPCPM(0).Value Then        'CPP
         If Not gSetFormula("CPPCPM", "'P'") Then
             gCmcGen30 = -1
             Exit Function
         End If
    Else                                     'CPM
         If Not gSetFormula("CPPCPM", "'M'") Then
             gCmcGen30 = -1
             Exit Function
         End If
    End If
   
    If RptSel30!rbcBook(0).Value Then        'sch line book
        If Not gSetFormula("BookToUse", "'L'") Then    '
            gCmcGen30 = -1
            Exit Function
        End If
    ElseIf RptSel30!rbcBook(1).Value Then        'vehicle default
        If Not gSetFormula("BookToUse", "'V'") Then    'Select
            gCmcGen30 = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("BookToUse", "'C'") Then    'Closest to airing
            gCmcGen30 = -1
            Exit Function
        End If
    End If
    
    If RptSel30!rbcDemo(0).Value Then        'primary
         slStr = " for Primary Demo"
    Else                                     'Select
         'get the demo selected
         ilTemp = RptSel30!lbcSelection(1).ListIndex
         slStr = "for " + RptSel30!lbcSelection(1).List(ilTemp)
    End If
    If Not gSetFormula("DemoToUse", "'" & slStr & "'") Then
        gCmcGen30 = -1
        Exit Function
    End If
        
    If RptSel30!rbcMonthType(0).Value = True Then     'corp
        slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year
        gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)
     
        ilYear = gGetYearofCorpMonth(igMonthOrQtr, igYear)
        If Not gSetFormula("MonthHeader", "'" & str$(ilSaveMonth) & " " & str$(ilYear) & "'") Then
            gCmcGen30 = -1
            Exit Function
        End If
        'adjusted actual starting month & year of the requested corporate start date to gather
        igYear = ilYear
        igMonthOrQtr = ilSaveMonth
    Else
        'igMonthOrQtr has starting Month
         'igYear has starting year
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year
        If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(igYear) & "'") Then
            gCmcGen30 = -1
            Exit Function
        End If
    End If
    'pass starting month for requested report for columm headings
    If Not gSetFormula("StartingMonth", igMonthOrQtr) Then
        gCmcGen30 = -1
        Exit Function
    End If

    'currently only standard implemented
    If RptSel30!rbcMonthType(0).Value Then          'calendar
        If Not gSetFormula("MonthType", "'C'") Then    '
            gCmcGen30 = -1
            Exit Function
        End If
    ElseIf RptSel30!rbcMonthType(1).Value Then      'corporate, user defined
        If Not gSetFormula("MonthType", "'U'") Then    '
            gCmcGen30 = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("MonthType", "'S'") Then    'std
            gCmcGen30 = -1
            Exit Function
        End If
    End If
    
    ilTemp = Val(RptSel30!edcNoMonths.Text)
    If Not gSetFormula("NoPeriods", ilTemp) Then
        gCmcGen30 = -1
        Exit Function
    End If
    
    For ilTemp = 0 To 9
        slLen(ilTemp) = Trim$(RptSel30!edcLen(ilTemp))
        slIndex(ilTemp) = Trim$(RptSel30!edcIndex(ilTemp))
    Next ilTemp
    gBuildSpotLenAndIndexTable slLen(), slIndex(), tlSpotLenRatio

    
    slSpotLenRatio = ""
    For ilTemp = 0 To 9
        If tlSpotLenRatio.iLen(ilTemp) = 0 Then         'done
            Exit For
        Else
            slStr = gIntToStrDec(tlSpotLenRatio.iRatio(ilTemp) / 10, 1)
            If Trim$(slSpotLenRatio) <> "" Then
                slSpotLenRatio = slSpotLenRatio & ","
            End If
            slSpotLenRatio = slSpotLenRatio & str$(tlSpotLenRatio.iLen(ilTemp)) & " @" & Trim$(slStr)
        End If
    Next ilTemp

   If Not gSetFormula("SpotLenRatio", "'" & slSpotLenRatio & "'") Then    'std
        gCmcGen30 = -1
        Exit Function
    End If
    
    '1-28-19 implement gross net option
    slStr = "G"
    If RptSel30!rbcGrossNet(1).Value Then
        slStr = "N"
    End If
    If Not gSetFormula("GrossNet", "'" & Trim$(slStr) & "'") Then
        gCmcGen30 = -1
        Exit Function
    End If

    
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGen30 = -1
        Exit Function
    End If
    gCmcGen30 = 1         'ok
    Exit Function
End Function
Function gGenReport30() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer
Dim slPeriodType As String * 1
Dim slYear As String
Dim slPeriods As String             '# periods
Dim slMonthDate As String           'starting month or date
Dim slLoLimit As String
Dim slHiLimit As String
Dim ilValid As Boolean
Dim ilInvalidField As Integer

        gGenReport30 = True
        If RptSel30!rbcMonthType(0).Value Then          'calendar
            slPeriodType = "C"
        ElseIf RptSel30!rbcMonthType(1).Value Then         'corp
            slPeriodType = "U"
        Else
            slPeriodType = "S"
        End If
        
        slMonthDate = Trim$(RptSel30!edcStartMonth.Text)
        slPeriods = Trim$(RptSel30!edcNoMonths.Text)
        slYear = Trim$(RptSel30!edcYear.Text)
        
        ilValid = gVerifyMonthYrPeriods(slPeriodType, slYear, slMonthDate, slPeriods, "1", "12", ilInvalidField)       'validity check year, month, # periods or start date & #periods Input
        If ilValid Then
            If RptSel30!rbcSortBy(0).Value = True Then
                If Not gOpenPrtJob("CP30UnitAdv.Rpt") Then
                    gGenReport30 = False
                    Exit Function
                End If
            Else
                If RptSel30!ckcDetail.Value = vbChecked Then
                    If Not gOpenPrtJob("CP30UnitVehDet.Rpt") Then
                        gGenReport30 = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("CP30UnitVeh.Rpt") Then
                        gGenReport30 = False
                        Exit Function
                    End If
                End If
            End If
        Else
            mReset
            If ilInvalidField = 1 Then
                RptSel30!edcYear.SetFocus
            ElseIf ilInvalidField = 2 Then
                RptSel30!edcStartMonth.SetFocus
            ElseIf ilInvalidField = 3 Then
                RptSel30!edcNoMonths.SetFocus
            End If
            RptSel30!cmcGen.Enabled = False
            gGenReport30 = False
            Exit Function
        End If

    Exit Function
End Function
Sub mReset()
    igGenRpt = False
    RptSel30!frcOutput.Enabled = igOutput
    RptSel30!frcCopies.Enabled = igCopies
    'RptSelIA!frcWhen.Enabled = igWhen
    RptSel30!frcFile.Enabled = igFile
    RptSel30!frcOption.Enabled = igOption
    'RptSelIA!frcRptType.Enabled = igReportType
    Beep
End Sub
