Attribute VB_Name = "RPTVFYPARPAY"


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
Function gCmcGenParPayables() As Integer
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


        gCmcGenParPayables = 0
        
        slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
                
        slExclude = ""
        slInclude = ""
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(1), slInclude, slExclude, "Orders"
        'gIncludeExcludeCkc RptSelParPay!ckcAllTypes(2), slInclude, slExclude, "Net"     Unused
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(3), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(4), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(5), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(6), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(7), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(8), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(9), slInclude, slExclude, "Promo"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(10), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(11), slInclude, slExclude, "AirTime"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(12), slInclude, slExclude, "Rep"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(13), slInclude, slExclude, "NTR"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(14), slInclude, slExclude, "H/C"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(15), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelParPay!ckcAllTypes(16), slInclude, slExclude, "Non-Polit"
      
        'only show inclusions
        If Len(slInclude) > 0 Then
            slInclude = "Include: " & slInclude
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenParPayables = -1
                Exit Function
            End If
        End If
        If Len(slExclude) <= 0 Then
            slExclude = "Exclude: None"
        Else
            slExclude = "Exclude: " & slExclude
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenParPayables = -1
                Exit Function
            End If
        End If
                

        'igMonthOrQtr has starting Month
        If Not gSetFormula("StartingMonth", igMonthOrQtr) Then
            gCmcGenParPayables = -1
            Exit Function
        End If
         'igYear has starting year
        If Not gSetFormula("StartingYear", igYear) Then
            gCmcGenParPayables = -1
            Exit Function
        End If

        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
        If Not gSetFormula("LastBilled", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
            gCmcGenParPayables = -1
            Exit Function
        End If

        If RptSelParPay!rbcTotalsBy(0).Value = True Then        'detail
            If Not gSetFormula("TotalsBy", "'D'") Then
                gCmcGenParPayables = -1
                Exit Function
            End If
        ElseIf RptSelParPay!rbcTotalsBy(1).Value = True Then   'vehicle
            If Not gSetFormula("TotalsBy", "'V'") Then
                gCmcGenParPayables = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("TotalsBy", "'P'") Then      'partner
                gCmcGenParPayables = -1
                Exit Function
            End If
        End If

        If RptSelParPay!rbcBillOrCollect(0).Value = True Then        'Billing Rendered
            If Not gSetFormula("WhichReport", "'B'") Then
                gCmcGenParPayables = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("WhichReport", "'C'") Then      'Cash
                gCmcGenParPayables = -1
                Exit Function
            End If
        End If
        
        If RptSelParPay!ckcAllVehicles.Value = vbChecked Then        'all vehicles selected?
            If Not gSetFormula("Legend", "''") Then         'show nothing in legend
                gCmcGenParPayables = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("Legend", "'Note: Selective vehicles are printed'") Then      'Cash
                gCmcGenParPayables = -1
                Exit Function
            End If
        End If
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfPer1Genl} = 1 and "
        slSelection = slSelection & " {GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    
        If Not gSetSelection(slSelection) Then
            gCmcGenParPayables = 0
            Exit Function
        End If
            
        gCmcGenParPayables = 1         'ok
    Exit Function
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
Function gGenReportParPayables() As Integer
Dim slStr As String
Dim ilValue As Integer
Dim slDateFrom As String
Dim llDate As Long
Dim ilHiLimit As Integer
Dim ilRet As Integer

        ilRet = mVerifyMonthYrPeriods()       'validity check year, month, # periods or start date & #periods Input
        If ilRet = 0 Then
            If Not gOpenPrtJob("ParPayable.Rpt") Then
                gGenReportParPayables = False
                Exit Function
            End If
        Else
            gGenReportParPayables = False
            Exit Function
        End If

    gGenReportParPayables = True
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
    RptSelParPay!frcOutput.Enabled = igOutput
    RptSelParPay!frcCopies.Enabled = igCopies
    'RptSelParPay!frcWhen.Enabled = igWhen
    RptSelParPay!frcFile.Enabled = igFile
    RptSelParPay!frcOption.Enabled = igOption
    'RptSelParPay!frcRptType.Enabled = igReportType
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
            slStr = RptSelParPay!edcStart.Text
            igYear = gVerifyYear(slStr)
            If igYear = 0 Then
                mReset
                RptSelParPay!edcStart.SetFocus                 'invalid year
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If
            slStr = RptSelParPay!edcMonth.Text                 'month in text form (jan..dec), or just a month # could have been entered
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                ilSaveMonth = Val(slStr)
            Else
                slStr = str$(ilSaveMonth)
            End If

            ilRet = gVerifyInt(slStr, 1, 12)        'if month number came back 0, its invalid
            If ilRet = -1 Then
                mReset
                RptSelParPay!edcMonth.SetFocus                 'invalid month #
                mVerifyMonthYrPeriods = -1
                Exit Function
            End If

            igMonthOrQtr = Val(slStr)
          

    Exit Function
End Function

