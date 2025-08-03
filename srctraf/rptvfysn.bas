Attribute VB_Name = "RptVfySN"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfysn.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelSN.Bas
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
Function gCmcGenSN() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                                                                            *
'******************************************************************************************

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
    Dim slInclude As String
    Dim slExclude As String
    gCmcGenSN = 0
    slDate = RptSelSN!edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelSN!edcSelCFrom.SetFocus
        Exit Function
    End If

    slStr = RptSelSN!edcSelCFrom1.Text                  'edit # weeks

    ilRet = gVerifyInt(slStr, 1, 52)                    '52 weeks maximum
    If ilRet = -1 Then
        mReset
        RptSelSN!edcSelCFrom1.SetFocus                 'invalid
        Exit Function
    End If


    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelSN!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelSN!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSelSN!ckcCType(2), slInclude, slExclude, "Feed"
    gIncludeExcludeCkc RptSelSN!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelSN!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelSN!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelSN!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelSN!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelSN!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelSN!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelSN!ckcCType(10), slInclude, slExclude, "Trade"

    gIncludeExcludeCkc RptSelSN!ckcSpots(0), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelSN!ckcSpots(1), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelSN!ckcSpots(2), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelSN!ckcSpots(3), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelSN!ckcSpots(4), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelSN!ckcSpots(5), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelSN!ckcSpots(6), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelSN!ckcSpots(7), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelSN!ckcSpots(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelSN!ckcSpots(9), slInclude, slExclude, "Spinoff"
    gIncludeExcludeCkc RptSelSN!ckcSpots(10), slInclude, slExclude, "MG"     '10-29-10
    
     If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenSN = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenSN = -1
            Exit Function
        End If
    End If

   If RptSelSN!rbcSort(0).Value Then        'show vehicle within week vs week within vehicle
        If Not gSetFormula("VehicleSortOrWeekSort", "'W'") Then     'skip to new page each week
            gCmcGenSN = -1
            Exit Function
        End If
   Else
        If Not gSetFormula("VehicleSortOrWeekSort", "'V'") Then     'skip to new page each vehicle
            gCmcGenSN = -1
            Exit Function
        End If
   End If

    If RptSelSN!ckcNewPage.Value Then        'skip to new page each vehicle or week
        If Not gSetFormula("NewPage", "'Y'") Then
            gCmcGenSN = -1
            Exit Function
        End If
   Else
        If Not gSetFormula("NewPage", "'N'") Then
            gCmcGenSN = -1
            Exit Function
        End If
   End If


    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenSN = -1
        Exit Function
    End If
    gCmcGenSN = 1         'ok
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
    RptSelSN!frcOutput.Enabled = igOutput
    RptSelSN!frcCopies.Enabled = igCopies
    'RptSelSN!frcWhen.Enabled = igWhen
    RptSelSN!frcFile.Enabled = igFile
    RptSelSN!frcOption.Enabled = igOption
    'RptSelSN!frcRptType.Enabled = igReportType
    Beep
End Sub
