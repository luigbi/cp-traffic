Attribute VB_Name = "RPTVFYAVGCMP"


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyParPay.Bas  Participant Payables
'
Option Explicit
Option Compare Text


'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportParPayables           *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcAverageCompare() As Integer
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
Dim slForPeriod As String
Dim slShowRCPrice As String

    gCmcAverageCompare = 0
    
    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
            
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(1), slInclude, slExclude, "Orders"
    'gIncludeExcludeCkc RPTSELAVGCMP!ckcAllTypes(2), slInclude, slExclude, "Net"     Unused
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(10), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(11), slInclude, slExclude, "AirTime"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(12), slInclude, slExclude, "Rep"
    'gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(13), slInclude, slExclude, "NTR"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(14), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(15), slInclude, slExclude, "Polit"
    gIncludeExcludeCkc RptSelAvgCmp!ckcAllTypes(16), slInclude, slExclude, "Non-Polit"
    slInclude = slInclude & ", " & IIF(RptSelAvgCmp!rbcUseLines(0) = True, "Pkg Lines", "Air Lines")
    slExclude = slExclude & ", " & IIF(RptSelAvgCmp!rbcUseLines(0) = False, "Pkg Lines", "Air Lines")
    
    'only show inclusions
    If Len(slInclude) > 0 Then
        slInclude = "Include: " & slInclude
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcAverageCompare = -1
            Exit Function
        End If
    End If
    If Len(slExclude) <= 0 Then
        slExclude = "Exclude: None"
    Else
        slExclude = "Exclude: " & slExclude
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcAverageCompare = -1
            Exit Function
        End If
    End If
            
    If Not gSetFormula("StartingYear", igYear) Then
        gCmcAverageCompare = -1
        Exit Function
    End If
    
    'Average Rate or Spot Price Comparison
    If Not gSetFormula("AvgRateSpot", IIF(RptSelAvgCmp!rbcAvgRatePrice(0).Value = True, "'Rate'", "'Price'")) Then
        gCmcAverageCompare = -1
        Exit Function
    End If
    
    'Spot Price Comparison: 30/60 or Combined
    If Not gSetFormula("AvgBy", IIF(RptSelAvgCmp!rbcAvgBy(0).Value = True, "'3060'", "'ALL'")) Then
        gCmcAverageCompare = -1
        Exit Function
    End If
    
    slForPeriod = "for " & str$(CInt(RptSelAvgCmp!edcStartYear.Text) - CInt(Trim(RptSelAvgCmp!edcYears.Text)) + 1) & " - " & RptSelAvgCmp!edcStartYear.Text
    If Not gSetFormula("ForPeriod", "'" & slForPeriod & "'") Then
        gCmcAverageCompare = -1
        Exit Function
    End If
    
    slShowRCPrice = IIF(RptSelAvgCmp!ckcShowUnitPrice.Value = 0, "N", "Y")
    If Not gSetFormula("ShowRCPrice", "'" & slShowRCPrice & "'") Then
        gCmcAverageCompare = -1
        Exit Function
    End If
    
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    'slSelection = "{GRF_Generic_Report.grfPer1Genl} = 1 and "
    slSelection = " {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
    If Not gSetSelection(slSelection) Then
        gCmcAverageCompare = 0
        Exit Function
    End If
        
    gCmcAverageCompare = 1         'ok

End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportParPayables         *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Validity checking of input &   *
'                               Open the Crystal report                *
'
'*******************************************************
Function gGenAverageCompare() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer

        ilRet = mVerifyMonthYrPeriods()       'validity check year, month, # periods or start date & #periods Input
        If ilRet = 0 Then
            If Not gOpenPrtJob("AvgRateSpotCmp.Rpt") Then
                gGenAverageCompare = False
                Exit Function
            End If
        Else
            gGenAverageCompare = False
            Exit Function
        End If

    gGenAverageCompare = True
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
    RptSelAvgCmp!frcOutput.Enabled = igOutput
    RptSelAvgCmp!frcCopies.Enabled = igCopies
    'RPTSELAVGCMP!frcWhen.Enabled = igWhen
    RptSelAvgCmp!frcFile.Enabled = igFile
    RptSelAvgCmp!frcOption.Enabled = igOption
    'RPTSELAVGCMP!frcRptType.Enabled = igReportType
    Beep
End Sub

'        verify input parameters for Year, Month and # periods
'        <input> None
'         return - 0: Valid input, -1 : error in input
Function mVerifyMonthYrPeriods() As Integer
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

    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
    mVerifyMonthYrPeriods = 0
    'verify user input dates
    slStr = RptSelAvgCmp!edcStartYear.Text
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelAvgCmp!edcStartYear.SetFocus          'invalid year
        mVerifyMonthYrPeriods = -1
        Exit Function
    End If
    
    slStr = RptSelAvgCmp!edcYears.Text          'month in text form (jan..dec), or just a month # could have been entered
    ilRet = gVerifyInt(slStr, 1, 5)            'if month number came back 0, its invalid
    If ilRet = -1 Then
        mReset
        RptSelAvgCmp!edcYears.SetFocus          'invalid month #
        mVerifyMonthYrPeriods = -1
        Exit Function
    End If
    
    igMonthOrQtr = Val(slStr)
End Function


