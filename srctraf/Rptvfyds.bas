Attribute VB_Name = "RPTVFYDS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyds.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyDS.Bas  Daily Spot Report
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
Function gCmcGenDS(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTime As String
    Dim slSelection As String
    Dim slInclude As String
    Dim slExclude As String
    gCmcGenDS = 0
    'check for valid start & end dates
'    slDate = RptSelDS!edcSelCFrom.Text
    slDate = RptSelDS!csi_CalFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelDS!csi_CalFrom.SetFocus
        Exit Function
    End If
    slDate = RptSelDS!csi_CalTo.Text
    If slDate <> "" Then
        If Not gValidDate(slDate) Then
            mReset
            RptSelDS!csi_CalTo.SetFocus
            Exit Function
        End If
    End If
    'check for valid start and end times
    slTime = RptSelDS!edcSTime.Text
    If Not gValidTime(slTime) Then
        mReset
        RptSelDS!edcSTime.SetFocus
        Exit Function
    End If

    slTime = RptSelDS!edcETime.Text
    If Not gValidTime(slTime) Then
        mReset
        RptSelDS!edcETime.SetFocus
        Exit Function
    End If
    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelDS!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelDS!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSelDS!ckcCType(2), slInclude, slExclude, "Net"
    gIncludeExcludeCkc RptSelDS!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelDS!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelDS!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelDS!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelDS!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelDS!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelDS!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelDS!ckcCType(10), slInclude, slExclude, "Trade"

    gIncludeExcludeCkc RptSelDS!ckcSpots(1), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelDS!ckcSpots(2), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelDS!ckcSpots(3), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelDS!ckcSpots(4), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelDS!ckcSpots(5), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelDS!ckcSpots(6), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelDS!ckcSpots(7), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelDS!ckcSpots(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelDS!ckcSpots(9), slInclude, slExclude, "Spinoff"
    gIncludeExcludeCkc RptSelDS!ckcSpots(0), slInclude, slExclude, "MG"

    gIncludeExcludeCkc RptSelDS!ckcRank(0), slInclude, slExclude, "Fixed Time"
    gIncludeExcludeCkc RptSelDS!ckcRank(1), slInclude, slExclude, "Sponsor"
    gIncludeExcludeCkc RptSelDS!ckcRank(2), slInclude, slExclude, "DP"
    gIncludeExcludeCkc RptSelDS!ckcRank(3), slInclude, slExclude, "ROS"
    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenDS = -1
            Exit Function
        End If
    End If
    If Len(slExclude) <= 0 Then
        slExclude = "None"
    End If
    If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
        gCmcGenDS = -1
        Exit Function
    End If
    If RptSelDS!ckcDays(7).Value = vbChecked Then              'skip new page
        If Not gSetFormula("NewPage", "'Y'") Then
            gCmcGenDS = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("NewPage", "'N'") Then
            gCmcGenDS = -1
            Exit Function
        End If
    End If
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenDS = -1
        Exit Function
    End If
    gCmcGenDS = 1         'ok
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
    RptSelDS!frcOutput.Enabled = igOutput
    RptSelDS!frcCopies.Enabled = igCopies
    'RptSelDS!frcWhen.Enabled = igWhen
    RptSelDS!frcFile.Enabled = igFile
    RptSelDS!frcOption.Enabled = igOption
    'RptSelDS!frcRptType.Enabled = igReportType
    Beep
End Sub
