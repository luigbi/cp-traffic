Attribute VB_Name = "RPTVFYAS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyas.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAS.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgRptSelAdvertiserCode() As SORTCODE
'Public sgRptSelAdvertiserCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
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
Function gCmcGenAs(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    gCmcGenAs = 0
'    slDate = RptSelAS!edcSelCFrom.Text              'start date
    slDate = RptSelAS!CSI_CalFrom.Text              'start date   9-4-19 use csi calendar control vs edit box
    If gValidDate(slDate) Then                      'end date
'        slDate = RptSelAS!edcSelCFrom1.Text
        slDate = RptSelAS!CSI_CalTo.Text
        If gValidDate(slDate) Then
            slDate = RptSelAS!CSI_CalTo.Text
        Else
            mReset
            RptSelAS!CSI_CalTo.SetFocus
            Exit Function
        End If
    Else
        mReset
        RptSelAS!CSI_CalFrom.SetFocus
        Exit Function
    End If
    slExclude = ""
    slInclude = ""
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(0), slInclude, slExclude, "Holds"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(1), slInclude, slExclude, "Orders"

    'gIncludeExcludeCkc RptSelAS!ckcSelC1(2), slInclude, slExclude, "Std"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(3), slInclude, slExclude, "Reserve"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(4), slInclude, slExclude, "Remnant"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(5), slInclude, slExclude, "DR"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(6), slInclude, slExclude, "PI"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(7), slInclude, slExclude, "PSA"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(8), slInclude, slExclude, "Promo"

    'gIncludeExcludeCkc RptSelAS!ckcSelC1(9), slInclude, slExclude, "Trade"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(10), slInclude, slExclude, "Missed"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(11), slInclude, slExclude, "N/C"
    'gIncludeExcludeCkc RptSelAS!ckcSelC1(12), slInclude, slExclude, "Fill"

    If igmnfRated <> 0 Then
        gIncludeExcludeCkc RptSelAS!ckcSelC3(0), slInclude, slExclude, "Rated"
        gIncludeExcludeCkc RptSelAS!ckcSelC3(1), slInclude, slExclude, "Non-Rated"
        gIncludeExcludeCkc RptSelAS!ckcSelC3(2), slInclude, slExclude, "Suburban"
    End If

    'local / national
    gIncludeExcludeCkc RptSelAS!ckcLocNatl(0), slInclude, slExclude, "Local Contracts"
    gIncludeExcludeCkc RptSelAS!ckcLocNatl(1), slInclude, slExclude, "Natl Contracts"

   'only show inclusion/exclusion of feed spots if not included
    If tgSpf.sSystemType = "R" Then         'radio system are only ones that can have feed spots
        gIncludeExcludeCkc RptSelAS!ckcCntrFeed(1), slInclude, slExclude, "Feed spots"
    End If
    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenAs = -1
        End If
    Else
        If Not gSetFormula("Included", "'" & " " & "'") Then
            gCmcGenAs = -1
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenAs = -1
        End If
    Else
        If Not gSetFormula("Excluded", "'" & " " & "'") Then
            gCmcGenAs = -1
        End If
    End If


    If RptSelAS!ckcLocNatl(0).Value = vbUnchecked Or RptSelAS!ckcLocNatl(1).Value = vbUnchecked Or RptSelAS!ckcLocNatl(2).Value = vbUnchecked Then
        If RptSelAS!ckcLocNatl(0).Value = vbChecked Then    'include local
            slSelection = slSelection & " and ({MNF_Multi_Names.mnfGroupNo} = 1"
            If Not RptSelAS!ckcLocNatl(1).Value = vbChecked Then
                slSelection = slSelection & ")"
                'slOrigin = "(Local Contracts Only)"
            End If
        End If
        If RptSelAS!ckcLocNatl(1).Value = vbChecked Then         'include natl
            If RptSelAS!ckcLocNatl(0).Value = vbChecked Then    'include local, append the natl test
                slSelection = slSelection & " or {MNF_Multi_Names.mnfGroupNo} = 3)"
            Else
                slSelection = slSelection & " and ({MNF_Multi_Names.mnfGroupNo} = 3)"
                'slOrigin = "(National Contracts Only)"
            End If
        End If
        'Place code for Regional to include if later implemented from sales source
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

    If Not gSetSelection(slSelection) Then
        gCmcGenAs = -1
        Exit Function
    End If
    gCmcGenAs = 1         'ok
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
    RptSelAS!frcOutput.Enabled = igOutput
    RptSelAS!frcCopies.Enabled = igCopies
    'rptselas!frcWhen.Enabled = igWhen
    RptSelAS!frcFile.Enabled = igFile
    RptSelAS!frcOption.Enabled = igOption
    'rptselas!frcRptType.Enabled = igReportType
    Beep
End Sub
