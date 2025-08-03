Attribute VB_Name = "RPTVFYCM"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyos.bas on Fri 3/12/10 @ 11:00 A
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCM.Bas - Competitive Categories by Daypart within vehicle or by Vehicle
'
' Release: 1.0
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
Function gCmcGenCM() As Integer
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
    Dim ilTopHowMany As Integer
    Dim llDate As Long
    Dim ilDay As Integer
    
    gCmcGenCM = 0
    slDate = RptSelCM!edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelCM!edcSelCFrom.SetFocus
        Exit Function
    End If

    slStr = RptSelCM!edcSelCFrom1.Text                  'edit qtr

    ilRet = gVerifyInt(slStr, 1, 13)                    '13 weeks maximum
    If ilRet = -1 Then
        mReset
        RptSelCM!edcSelCFrom1.SetFocus                 'invalid
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable

    slStr = RptSelCM!edcHowMany.Text
    ilRet = gVerifyInt(slStr, 0, 100)
    If ilRet = -1 Then
        mReset
        RptSelCM!edcTopMany.SetFocus                 'invalid
        Exit Function
    End If
    ilTopHowMany = Val(slStr)
    If ilTopHowMany = 0 Or ilTopHowMany >= 99 Then      'crystal needs to know max to print, set at high number
        ilTopHowMany = 99
    End If
    
'    slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
'    ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
'    If Not gSetFormula("RCHeader", "'" & slNameCode & "'") Then
'        gCmcGenCM = -1
'        Exit Function
'    End If
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelCM!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelCM!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSelCM!ckcCType(2), slInclude, slExclude, "Feed"
    gIncludeExcludeCkc RptSelCM!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelCM!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelCM!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelCM!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelCM!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelCM!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelCM!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelCM!ckcCType(10), slInclude, slExclude, "Trade"

    gIncludeExcludeCkc RptSelCM!ckcSpots(0), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelCM!ckcSpots(1), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelCM!ckcSpots(2), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelCM!ckcSpots(3), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelCM!ckcSpots(4), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelCM!ckcSpots(5), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelCM!ckcSpots(6), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelCM!ckcSpots(7), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelCM!ckcSpots(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelCM!ckcSpots(9), slInclude, slExclude, "Spinoff"
    gIncludeExcludeCkc RptSelCM!ckcSpots(10), slInclude, slExclude, "MG"

    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenCM = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenCM = -1
            Exit Function
        End If
    End If
 
    If RptSelCM!rbcTotals(0).Value Then        'Counts by DP within vehicle, or vehicleonly
         If Not gSetFormula("TotalsBy", "'D'") Then    'detail, daypart within vehicle
             gCmcGenCM = -1
             Exit Function
         End If
    Else
         If Not gSetFormula("TotalsBy", "'V'") Then    'summary by vehicle
             gCmcGenCM = -1
             Exit Function
         End If
    End If
   
    If RptSelCM!ckcNewPage.Value = vbChecked Then        'new page each new vehicle
        If Not gSetFormula("NewPage", "'Y'") Then    '
            gCmcGenCM = -1
            Exit Function
        End If
   Else
        If Not gSetFormula("NewPage", "'N'") Then    'do not skip to new page each vehicle
            gCmcGenCM = -1
            Exit Function
        End If
   End If
   
   If Not gSetFormula("MaxWeeks", igMonthOrQtr) Then    'max number weeks to print
        gCmcGenCM = -1
        Exit Function
    End If
   If Not gSetFormula("TopHowMany", ilTopHowMany) Then    'top how many competitives to print
        gCmcGenCM = -1
        Exit Function
    End If
    
    llDate = gDateValue(slDate)
    'backup to Monday
    ilDay = gWeekDayLong(llDate)
    Do While ilDay <> 0
        llDate = llDate - 1
        ilDay = gWeekDayLong(llDate)
    Loop

    slDate = Format$(llDate, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If Not gSetFormula("P1", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        gCmcGenCM = -1
        Exit Function
    End If
        
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenCM = -1
        Exit Function
    End If
    gCmcGenCM = 1         'ok
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
    RptSelCM!frcOutput.Enabled = igOutput
    RptSelCM!frcCopies.Enabled = igCopies
    'RptSelCM!frcWhen.Enabled = igWhen
    RptSelCM!frcFile.Enabled = igFile
    RptSelCM!frcOption.Enabled = igOption
    'RptSelCM!frcRptType.Enabled = igReportType
    Beep
End Sub
