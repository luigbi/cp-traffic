Attribute VB_Name = "RptVfyRK"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyrk.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyRK.Bas
'
' Release: 5.6
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gCmcGenPriceRanking() As Integer
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
    Dim ilLoop As Integer
    Dim ilSaveMonth As Integer

        gCmcGenPriceRanking = 0
    
        If RptSelRk.rbcPeriodType(0).Value Then         'month (vs week)
        
            slStr = RptSelRk!edcSelCFrom.Text             'month in text form (jan..dec)
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                ilSaveMonth = Val(slStr)
                If ilSaveMonth = 0 Then                     'invalid converion
                    mReset
                    RptSelRk!edcSelCFrom.SetFocus
                    Exit Function
                End If
            End If
            If Val(RptSelRk!edcYear.Text) < 1970 Or Val(RptSelRk!edcYear.Text) > 2069 Then
                mReset
                RptSelRk!edcYear.SetFocus
                Exit Function
            End If
            'Format the base date Month & year spans to send to Crystal
            slDate = Trim$(Str$(ilSaveMonth)) & "/15/" & Trim$(RptSelRk!edcYear.Text)
            slDate = gObtainStartStd(slDate)
    
            If Not gValidDate(slDate) Then
                mReset
                RptSelRk!edcSelCFrom.SetFocus
                Exit Function
            End If
            slStr = RptSelRk!edcSelCFrom1.Text
            ilRet = gVerifyInt(slStr, 1, 12)                    '12 months maximum, 1 yr
            If ilRet = -1 Then
                mReset
                RptSelRk!edcSelCFrom1.SetFocus                 'invalid
                Exit Function
            End If
            igYear = Val(RptSelRk!edcYear.Text)
            igMonthOrQtr = ilSaveMonth                       'put # periods in global variable
        Else                                                'weekly
            slDate = RptSelRk!CSI_CalWeek.Text                  'verify the date entered
            If Not gValidDate(RptSelRk!CSI_CalWeek.Text) Then
                mReset
                RptSelRk!CSI_CalWeek.SetFocus
                Exit Function
            End If
    
            slStr = RptSelRk!edcSelCFrom1.Text                  'edit # periods
            ilRet = gVerifyInt(slStr, 1, 53)                    '53 weeks maximum, 1 yr
            If ilRet = -1 Then
                mReset
                RptSelRk!edcSelCFrom1.SetFocus                 'invalid
                Exit Function
            End If

            
         End If
    
        slExclude = ""
        slInclude = ""
        gIncludeExcludeCkc RptSelRk!ckcCType(0), slInclude, slExclude, "Holds"
        gIncludeExcludeCkc RptSelRk!ckcCType(1), slInclude, slExclude, "Orders"
        'gIncludeExcludeCkc RptSelRk!ckcCType(2), slInclude, slExclude, "Feed"
        gIncludeExcludeCkc RptSelRk!ckcCType(3), slInclude, slExclude, "Std"
        gIncludeExcludeCkc RptSelRk!ckcCType(4), slInclude, slExclude, "Reserve"
        gIncludeExcludeCkc RptSelRk!ckcCType(5), slInclude, slExclude, "Remnant"
        gIncludeExcludeCkc RptSelRk!ckcCType(6), slInclude, slExclude, "DR"
        gIncludeExcludeCkc RptSelRk!ckcCType(7), slInclude, slExclude, "PI"
        gIncludeExcludeCkc RptSelRk!ckcCType(8), slInclude, slExclude, "PSA"
        gIncludeExcludeCkc RptSelRk!ckcCType(9), slInclude, slExclude, "Promo"
        gIncludeExcludeCkc RptSelRk!ckcCType(10), slInclude, slExclude, "Trade"
        gIncludeExcludeCkc RptSelRk!ckcSpots(11), slInclude, slExclude, "Polit"
        gIncludeExcludeCkc RptSelRk!ckcSpots(12), slInclude, slExclude, "Non-Polit"
        gIncludeExcludeCkc RptSelRk!ckcUnder30, slInclude, slExclude, "Under30"

        gIncludeExcludeCkc RptSelRk!ckcSpots(0), slInclude, slExclude, "Missed"
        gIncludeExcludeCkc RptSelRk!ckcSpots(1), slInclude, slExclude, "Charge"
        gIncludeExcludeCkc RptSelRk!ckcSpots(2), slInclude, slExclude, "0.00"
        gIncludeExcludeCkc RptSelRk!ckcSpots(3), slInclude, slExclude, "ADU"
        gIncludeExcludeCkc RptSelRk!ckcSpots(4), slInclude, slExclude, "Bonus"
        gIncludeExcludeCkc RptSelRk!ckcSpots(5), slInclude, slExclude, "+Fill"
        gIncludeExcludeCkc RptSelRk!ckcSpots(6), slInclude, slExclude, "-Fill"
        gIncludeExcludeCkc RptSelRk!ckcSpots(7), slInclude, slExclude, "N/C"
        gIncludeExcludeCkc RptSelRk!ckcSpots(8), slInclude, slExclude, "Recap"
        gIncludeExcludeCkc RptSelRk!ckcSpots(9), slInclude, slExclude, "Spinoff"
        gIncludeExcludeCkc RptSelRk!ckcSpots(10), slInclude, slExclude, "MG"        '10-28-10
    
        If Len(slInclude) > 0 Then
            If Not gSetFormula("Included", "'" & slInclude & "'") Then
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        End If
        If Len(slExclude) > 0 Then
            If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        End If


        For ilLoop = 0 To RptSelRk!lbcSelection(1).ListCount - 1 Step 1
            If RptSelRk!lbcSelection(1).Selected(ilLoop) Then
                slNameCode = tgRateCardCode(ilLoop).sKey
                Exit For
            End If
        Next ilLoop
        ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
        If Not gSetFormula("RCHeader", "'" & slNameCode & "'") Then
            gCmcGenPriceRanking = -1
            Exit Function
        End If

        
        If RptSelRk!ckcNewPage.Value = vbChecked Then       'new page each vehicle
            If Not gSetFormula("NewPage", "'Y'") Then    '
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then    '
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        End If
        
        'Show DP within vehicle totals, or entire vehicle together
        If RptSelRk!rbcTotalsBy(0).Value = True Then       'totals by DP or vehicle
            If Not gSetFormula("DPorVehicleTotals", "'D'") Then    '
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("DPorVehicleTotals", "'V'") Then    '
                gCmcGenPriceRanking = -1
                Exit Function
            End If
        End If

        'Column to use to sort Top down
        ilLoop = RptSelRk!cbcSort.ListIndex
        If Not gSetFormula("TopDownColumn", ilLoop) Then    '
            gCmcGenPriceRanking = -1
            Exit Function
        End If
        
'        If RptSelRk!ckcUnder30.Value = vbChecked Then       'show information based on fractions because avails and/or spots under 30" includd
'            If Not gSetFormula("ShowFractions", "'Y'") Then    '
'                gCmcGenPriceRanking = -1
'                Exit Function
'            End If
'        Else
'            If Not gSetFormula("ShowFractions", "'N'") Then    '
'                gCmcGenPriceRanking = -1
'                Exit Function
'            End If
'        End If
'
        
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenPriceRanking = -1
            Exit Function
        End If
        gCmcGenPriceRanking = 1         'ok
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
    RptSelRk!frcOutput.Enabled = igOutput
    RptSelRk!frcCopies.Enabled = igCopies
    RptSelRk!frcFile.Enabled = igFile
    RptSelRk!frcOption.Enabled = igOption
    Beep
End Sub
