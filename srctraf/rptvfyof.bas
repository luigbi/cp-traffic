Attribute VB_Name = "RPTVFYOF"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyof.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyOF.Bas  Order Fullfilment Report
'
' Release: 4.7  10/9/2000
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag() As String
'Public tgRptAgencyCode() As SORTCODE
'Public sgRptAgencyCodeTag As String
'Public tgRptSalespersonCode() As SORTCODE
'Public sgRptSalespersonCodeTag As String
'Public tgRptAdvertiserCode() As SORTCODE
'Public sgRptAdvertiserCodeTag As String
'Public tgRptNameCode() As SORTCODE
'Public sgRptNameCodeTag As String
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
''Global spot types for Spots by Advt & spots by Date & Time
''bit selectivity for charged and different types of no charge spots
''bits defined right to left (0 to 9)
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
'
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
Function gCmcGenOF(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim slDateFrom As String
    Dim slDateTo As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTime As String
    Dim slSelection As String
    Dim slInclude As String
    Dim slExclude As String
    Dim llDate As Long

    gCmcGenOF = 0
    'check for valid start & end dates
'    slDateFrom = RptSelOF!edcSelCFrom.Text
    slDateFrom = RptSelOF!CSI_CalFrom.Text
    If Not gValidDate(slDateFrom) Then
        mReset
        RptSelOF!CSI_CalFrom.SetFocus
        Exit Function
    End If
    llDate = gDateValue(slDateFrom)
    slDateFrom = Format$(llDate, "m/d/yy")
    slDateTo = RptSelOF!CSI_CalTo.Text
    If slDateTo = "" Then
        slDateTo = slDateFrom
    End If
    'If slDateTo <> "" Then
        If Not gValidDate(slDateTo) Then
            mReset
            RptSelOF!CSI_CalTo.SetFocus
            Exit Function
        End If
    'End If
    llDate = gDateValue(slDateTo)
    slDateTo = Format$(llDate, "m/d/yy")


    If Not gSetFormula("RptDates", "'" & slDateFrom & " - " & slDateTo & "'") Then
        gCmcGenOF = -1
        Exit Function
    End If

    If RptSelOF!CkcDiscrepOnly.Value = vbChecked Then              'discreps only
        If Not gSetFormula("DiscrepsOnly", "'D'") Then
            gCmcGenOF = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("DiscrepsOnly", "'A'") Then           'all spots
            gCmcGenOF = -1
            Exit Function
        End If
    End If
    'the following inclusions/exclusions have been hidden in screen selectivity
    'Maintained if more selectivity is required since this code was copied from
    'Daily Spot report modules
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelOF!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelOF!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSelOF!ckcCType(2), slInclude, slExclude, "Net"
    gIncludeExcludeCkc RptSelOF!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelOF!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelOF!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelOF!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelOF!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelOF!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelOF!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelOF!ckcCType(10), slInclude, slExclude, "Trade"

    gIncludeExcludeCkc RptSelOF!ckcSpots(1), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelOF!ckcSpots(2), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelOF!ckcSpots(3), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelOF!ckcSpots(4), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelOF!ckcSpots(5), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelOF!ckcSpots(6), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelOF!ckcSpots(7), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelOF!ckcSpots(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelOF!ckcSpots(9), slInclude, slExclude, "Spinoff"

   ' gIncludeExcludeCkc RptSelOF!ckcRank(0), slInclude, slExclude, "Fixed Time"
   ' gIncludeExcludeCkc RptSelOF!ckcRank(1), slInclude, slExclude, "Sponsor"
   ' gIncludeExcludeCkc RptSelOF!ckcRank(2), slInclude, slExclude, "DP"
   ' gIncludeExcludeCkc RptSelOF!ckcRank(3), slInclude, slExclude, "ROS"
    'If Len(slInclude) > 0 Then
    '    If Not gSetFormula("Included", "'" & slInclude & "'") Then
    '        gCmcGenOF = -1
    '        Exit Function
    '    End If
    'End If
    'If Len(slExclude) <= 0 Then
    '    slExclude = "None"
    'End If
    'If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
    '    gCmcGenOF = -1
    '    Exit Function
    'End If
    'If RptSelOF!ckcDays(7).Value = vbChecked Then              'skip new page
    '    If Not gSetFormula("NewPage", "'Y'") Then
    '        gCmcGenOF = -1
    '        Exit Function
    '    End If
    'Else
    '    If Not gSetFormula("NewPage", "'N'") Then
    '        gCmcGenOF = -1
    '        Exit Function
    '    End If
    'End If
    gCurrDateTime slDateFrom, slTime, slMonth, slDay, slYear
    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenOF = -1
        Exit Function
    End If
    gCmcGenOF = 1         'ok
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
    RptSelOF!frcOutput.Enabled = igOutput
    RptSelOF!frcCopies.Enabled = igCopies
    'RptSelOF!frcWhen.Enabled = igWhen
    RptSelOF!frcFile.Enabled = igFile
    RptSelOF!frcOption.Enabled = igOption
    'RptSelOF!frcRptType.Enabled = igReportType
    Beep
End Sub
