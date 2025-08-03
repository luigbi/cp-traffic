Attribute VB_Name = "RPTVFYAD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyad.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyAD.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Function gGenReportAD(ilListIndex As Integer) As Integer

    If ilListIndex = DELIVERY_AUDIENCE Then
        If Not gOpenPrtJob("AudDel.Rpt") Then
            gGenReportAD = False
            Exit Function
        End If
    ElseIf ilListIndex = DELIVERY_POSTBUY Then
        If Not gOpenPrtJob("PostBuy.Rpt") Then
            gGenReportAD = False
            Exit Function
        End If
    End If
    gGenReportAD = True
End Function

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
Function gCmcGenAd() As Integer
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
    Dim slStr As String
    Dim slTime As String
    Dim slSelection As String
    ReDim ilDate(0 To 1) As Integer
    Dim llDate As Long
    Dim ilListIndex As Integer
    Dim slInclude As String
    Dim slExclude As String

    gCmcGenAd = 0
    ilListIndex = RptSelAD!lbcRptType.ListIndex

    If (RptSelAD!ckcAll.Value = vbChecked) Then
'        slDate = RptSelAD!edcSelCFrom.Text
        slDate = RptSelAD!CSI_CalFrom.Text          '9-3-19 use csi cal control vs edit box
        If Not gValidDate(slDate) Then
            mReset
            RptSelAD!CSI_CalFrom.SetFocus
            Exit Function
        End If
        slDate = RptSelAD!CSI_CalTo.Text
        If Not gValidDate(slDate) Then
            mReset
            RptSelAD!CSI_CalTo.SetFocus
            Exit Function
        End If
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenAd = -1
        Exit Function
    End If
    'Active Dates entered
'    slDate = RptSelAD!edcSelCFrom.Text
    slDate = RptSelAD!CSI_CalFrom.Text          '9-3-19 use csi cal control vs edit box
    gPackDate slDate, ilDate(0), ilDate(1)
    gUnpackDateLong ilDate(0), ilDate(1), llDate
    If slDate = "" Then
        slStr = ""
    Else
        slStr = Format$(llDate, "m/d/yy")
    End If
'    slDate = RptSelAD!edcSelCTo.Text
    slDate = RptSelAD!CSI_CalTo.Text
    gPackDate slDate, ilDate(0), ilDate(1)
    gUnpackDateLong ilDate(0), ilDate(1), llDate
    If slStr = "" And slDate = "" Then                          'selective contracts, no dates enterd
        slStr = "All dates for selective contracts"
    Else
        slStr = slStr & " - " & Format$(llDate, "m/d/yy")
    End If
    If Not gSetFormula("ActiveDates", "'" & slStr & "'") Then
        gCmcGenAd = -1
        Exit Function
    End If

    If ilListIndex = DELIVERY_AUDIENCE Then         'Audience delivery report
        If RptSelAD!rbcBook(0).Value Then       'closest book
            If Not gSetFormula("Book", "'Use Closest book to airing'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        ElseIf RptSelAD!rbcBook(1).Value Then
            If Not gSetFormula("Book", "'Use vehicle default book'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("Book", "'Use schedule line book'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If
        'send formulas for report header
        If RptSelAD!rbcCPPCPM(0).Value Then     'CPP
            If Not gSetFormula("CPPCPM", "'P'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("CPPCPM", "'M'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If

        If RptSelAD!rbcSortby(0).Value Then     'sort by ADvt?
            If Not gSetFormula("Sortby", "'V'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        ElseIf RptSelAD!rbcSortby(1).Value Then 'sort by Slsp
            If RptSelAD!rbcSubsort(1).Value Then        'subsort ascending
                If Not gSetFormula("Sortby", "'U'") Then
                    gCmcGenAd = -1
                    Exit Function
                End If
            ElseIf RptSelAD!rbcSubsort(2).Value Then    'subsort descending
                If Not gSetFormula("Sortby", "'O'") Then
                    gCmcGenAd = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("Sortby", "'S'") Then
                    gCmcGenAd = -1
                    Exit Function
                End If
            End If
        ElseIf RptSelAD!rbcSortby(2).Value Then 'sort by over/under (ascending)
            If Not gSetFormula("Sortby", "'D'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else                                    'sort by over/under (descending)
            If Not gSetFormula("Sortby", "'A'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If
    Else                                         'Post Buy
        If RptSelAD!ckcShowTime.Value = vbChecked Then        'show time column
            If Not gSetFormula("ShowTimeColumn", "'Y'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowTimeColumn", "'N'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If
        If RptSelAD!ckcMGAud.Value = vbChecked Then        'show Pd+MG line
            If Not gSetFormula("ShowMGDiff", "'Y'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowMGDiff", "'N'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If
        If RptSelAD!ckcBonusAud.Value = vbChecked Then        'show Bonus line
            If Not gSetFormula("ShowBonusDiff", "'Y'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("ShowBonusDiff", "'N'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If

        If RptSelAD!ckcNewPage.Value = vbChecked Then        'new page per adv
            If Not gSetFormula("NewPage", "'Y'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If

        If RptSelAD!rbcSortby(0).Value Then 'sort Advertiser
            If Not gSetFormula("TotalsBy", "'A'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else                                    'sort by contract
            If Not gSetFormula("TotalsBy", "'C'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If

        If RptSelAD!rbcShowGrimps(0).Value Then 'show gross impressions by units or thousands
            If Not gSetFormula("ShowImpBy", "'U'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        Else                                    'show grimps by thousands - how site aud is set
            If Not gSetFormula("ShowImpBy", "'T'") Then
                gCmcGenAd = -1
                Exit Function
            End If
        End If
        slInclude = ""
        slExclude = ""
        gIncludeExcludeCkc RptSelAD!ckcSelC5(0), slInclude, slExclude, "Charge"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(1), slInclude, slExclude, "0.00"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(2), slInclude, slExclude, "ADU"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(3), slInclude, slExclude, "Bonus"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(4), slInclude, slExclude, "+Fill"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(5), slInclude, slExclude, "-Fill"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(6), slInclude, slExclude, "N/C"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(7), slInclude, slExclude, "MG"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(8), slInclude, slExclude, "Recap"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(9), slInclude, slExclude, "Spinoff"
        gIncludeExcludeCkc RptSelAD!ckcSelC5(10), slInclude, slExclude, "Missed"
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenAd = -1
            Exit Function
            End If
        If Len(slExclude) <= 0 Then
            slExclude = "None"
        End If
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenAd = -1
            Exit Function
        End If

    End If
    gCmcGenAd = 1         'ok
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
    RptSelAD!frcOutput.Enabled = igOutput
    RptSelAD!frcCopies.Enabled = igCopies
    'RptSelAD!frcWhen.Enabled = igWhen
    RptSelAD!frcFile.Enabled = igFile
    RptSelAD!frcOption.Enabled = igOption
    'RptSelAD!frcRptType.Enabled = igReportType
    Beep
End Sub
