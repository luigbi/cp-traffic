Attribute VB_Name = "RPTVFYQB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyqb.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
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
Function gCmcGenQB(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    Dim llDate As Long
    Dim slSelectedAvails As String
    Dim illoop As Integer
    Dim llStartMonday As Long
    Dim llEndMonday As Long
    Dim slEffDate As String
    
    gCmcGenQB = 0
'    slDate = RptSelQB!edcSelCFrom.Text
    slDate = RptSelQB!CSI_CalFrom.Text          '12-11-19 change to use csi calendar control
    If Not gValidDate(slDate) Then
        mReset
        RptSelQB!CSI_CalFrom.SetFocus
        Exit Function
    End If

    slStr = RptSelQB!edcSelCFrom1.Text                  'edit qtr
    
    If RptSelQB!rbcVersion(1).Value = True Then
        'TTP 10729 - Quarterly Booked Spots report: add digital lines - This version of the report always runs for one broadcast month, which is 4 or 5 weeks.
        'Get the # of weeks in the selected Std Month
        slStr = DateDiff("W", gObtainStartStd(RptSelQB!CSI_CalFrom.Text), gObtainEndStd(RptSelQB!CSI_CalFrom.Text)) + 1
    End If
    
    'ilRet = gVerifyInt(slStr, 1, 4)
    ilRet = gVerifyInt(slStr, 1, 53)                '4-27-12 chg from quarters to weeks
    If ilRet = -1 Then
        mReset
        If RptSelQB!edcSelCFrom1.Visible = True Then
            RptSelQB!edcSelCFrom1.SetFocus                 'invalid qtr
        End If
        gCmcGenQB = -1
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable

    'TTP 10729 - Quarterly Booked Spots report
    If RptSelQB!rbcVersion(1).Value = True Then
        slDate = gObtainStartStd(RptSelQB!CSI_CalFrom.Text) 'The start date will always be the start of a standard broadcast month. If a start date is selected that is not the start of the broadcast month, it will get automatically set to the start of the standard broadcast month that the selected date is in.
    End If
    
    llDate = gDateValue(slDate)
    slDate = Format$(llDate, "m/d/yy")               'insure year is appended to month/day
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If Not gSetFormula("EffDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        gCmcGenQB = -1
        Exit Function
    End If
    
    '4-27-12 determine last monday to print
    llStartMonday = gDateValue(slDate)
    'backup start date to Monday
    illoop = gWeekDayLong(llStartMonday)
    Do While illoop <> 0
        llStartMonday = llStartMonday - 1
        illoop = gWeekDayLong(llStartMonday)
    Loop
    
    'Calculate last monday to print (based on using std start date or date entered)to send to crystal
    'llStartMonday = monday date
    slDate = Format(llStartMonday, "m/d/yy")
    If RptSelQB!rbcStart(0).Value Then          'use std qtr
        'get standard bdcst year from the start date entered
        slDate = gObtainYearStartDate(0, slDate)
        llDate = gDateValue(slDate)
        Do While (llStartMonday < llDate) Or (llStartMonday > llDate + 13 * 7 - 1)
            llDate = llDate + 13 * 7
        Loop
    End If

    '6-6-12
    'slDAte = start of std year (string), llDate = start of std year (long), llStartMonday = start date of avails
    'determine how many weeks in the past that doesnt apply until gathering actual data
    'will need to adjust the # of weeks to print.  Need to print actual weeks of data starting with the user start date entered
    igMonthOrQtr = ((llStartMonday - llDate) / 7) + igMonthOrQtr
    llEndMonday = llDate + ((igMonthOrQtr - 1) * 7)
    slDate = Format$(llEndMonday, "m/d/yy")
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If Not gSetFormula("MonEndDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        gCmcGenQB = -1
        Exit Function
    End If
    
    slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
    If Not gSetFormula("RCHeader", "'" & slNameCode & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    If RptSelQB!ckcSelC3(0).Value = vbChecked Then              'hide rates
        slStr = "H"
    Else
        slStr = "S"
    End If
    If Not gSetFormula("HideRate", "'" & slStr & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    If RptSelQB!ckcSelC4(0).Value = vbChecked Then              'hide selling vehicle
        slStr = "H"
    Else
        slStr = "S"
    End If
    If Not gSetFormula("HideSellVeh", "'" & slStr & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    If RptSelQB!rbcSelC2(0).Value Then              'hide
        slStr = "H"
    ElseIf RptSelQB!rbcSelC2(1).Value Then          'show separately
        slStr = "S"
    Else
        slStr = "E"                                 'exclude
    End If
    If Not gSetFormula("Reserved", "'" & slStr & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelQB!ckcSelC1(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(1), slInclude, slExclude, "Orders"

    gIncludeExcludeCkc RptSelQB!ckcSelC1(2), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(3), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(4), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(5), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(6), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(7), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(8), slInclude, slExclude, "Promo"

    gIncludeExcludeCkc RptSelQB!ckcSelC1(9), slInclude, slExclude, "Trade"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(10), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(11), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelQB!ckcSelC1(12), slInclude, slExclude, "Fill"

    If tgSpf.sSystemType = "R" Then     'only show exclusion if Radio system (vs Network/syndicator)
        If Not RptSelQB!ckcCntrFeed(0).Value = vbChecked Then
            gIncludeExcludeCkc RptSelQB!ckcCntrFeed(0), slInclude, slExclude, "Contract Spots"
        End If
        If Not RptSelQB!ckcCntrFeed(1).Value = vbChecked Then
            gIncludeExcludeCkc RptSelQB!ckcCntrFeed(1), slInclude, slExclude, "Feed Spots"
        End If
    End If

    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenQB = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenQB = -1
            Exit Function
        End If
    End If

    slSelectedAvails = ""
    If RptSelQB!ckcAllAvails.Value = vbUnchecked Then
        For illoop = 0 To RptSelQB!lbcSelection(3).ListCount - 1
            If RptSelQB!lbcSelection(3).Selected(illoop) Then
                slNameCode = tgNamedAvail(illoop).sKey         'sales source code
                ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                If slSelectedAvails = "" Then
                    slSelectedAvails = "Missed Named Avails : " & Trim$(slNameCode)
                Else
                    slSelectedAvails = slSelectedAvails & ", " & Trim$(slNameCode)
                End If
            End If
        Next illoop
    End If

    If Not gSetFormula("SelectedAvails", "'" & slSelectedAvails & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    
    '6-14-19 feature to show by 30" units or spot counts
    If RptSelQB!rbc30sOrUnits(0).Value Then              '30" unit counts
        slStr = "3"
    ElseIf RptSelQB!rbc30sOrUnits(1).Value Then          'spot count
        slStr = "U"
    End If
    If Not gSetFormula("30sOrUnits", "'" & slStr & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If
    '6-14-19 feature to calc gross or net rates
    If RptSelQB!rbcGrossNet(0).Value Then
        slStr = "G"
    ElseIf RptSelQB!rbcGrossNet(1).Value Then
        slStr = "N"
    End If
    If Not gSetFormula("GrossNet", "'" & slStr & "'") Then
        gCmcGenQB = -1
        Exit Function
    End If



    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{AVR_Quarterly_Avail.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({AVR_Quarterly_Avail.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenQB = -1
        Exit Function
    End If
    gCmcGenQB = 1         'ok
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
    RptSelQB!frcOutput.Enabled = igOutput
    RptSelQB!frcCopies.Enabled = igCopies
    'RptSelQB!frcWhen.Enabled = igWhen
    RptSelQB!frcFile.Enabled = igFile
    RptSelQB!frcOption.Enabled = igOption
    'RptSelQB!frcRptType.Enabled = igReportType
    Beep
End Sub
