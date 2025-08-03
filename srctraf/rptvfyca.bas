Attribute VB_Name = "RptVfyCA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyca.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyCA.Bas
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
Function gCmcGenCA() As Integer
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
    Dim ilListIndex As Integer
    Dim illoop As Integer
    Dim blGross As Boolean      'Date: 1/10/2020 gross or net computation
    
    ilListIndex = RptSelCA!lbcRptType.ListIndex
    gCmcGenCA = 0

    'For sports version - Start Date and end date fields
    'for non-sports version, start date and # weeks
'    slDate = RptSelCA!edcSelCFrom.Text
    slDate = RptSelCA!CSI_CalFrom.Text          '9-4-19 use csi calendar control instead of edit box
    If Not gValidDate(slDate) Then
        mReset
        RptSelCA!CSI_CalFrom.SetFocus
        Exit Function
    End If

    If ilListIndex = AVAILSCOMBO_SPORTS Then
        slStr = RptSelCA!CSI_CalTo.Text          '9-4-19
        If Not gValidDate(slStr) Then
            mReset
            RptSelCA!CSI_CalTo.SetFocus
            Exit Function
        End If
    Else
        slStr = RptSelCA!edcSelCFrom1.Text                  'edit qtr
        ilRet = gVerifyInt(slStr, 1, 14)                    '14 weeks maximum
        If ilRet = -1 Then
            mReset
            RptSelCA!edcSelCFrom1.SetFocus                 'invalid
            Exit Function
        End If
        igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
    End If

    slExclude = ""
    slInclude = ""
    gIncludeExcludeCkc RptSelCA!ckcCType(0), slInclude, slExclude, "Holds"
    gIncludeExcludeCkc RptSelCA!ckcCType(1), slInclude, slExclude, "Orders"
    gIncludeExcludeCkc RptSelCA!ckcCType(2), slInclude, slExclude, "Feed"
    gIncludeExcludeCkc RptSelCA!ckcCType(3), slInclude, slExclude, "Std"
    gIncludeExcludeCkc RptSelCA!ckcCType(4), slInclude, slExclude, "Reserve"
    gIncludeExcludeCkc RptSelCA!ckcCType(5), slInclude, slExclude, "Remnant"
    gIncludeExcludeCkc RptSelCA!ckcCType(6), slInclude, slExclude, "DR"
    gIncludeExcludeCkc RptSelCA!ckcCType(7), slInclude, slExclude, "PI"
    gIncludeExcludeCkc RptSelCA!ckcCType(8), slInclude, slExclude, "PSA"
    gIncludeExcludeCkc RptSelCA!ckcCType(9), slInclude, slExclude, "Promo"
    gIncludeExcludeCkc RptSelCA!ckcCType(10), slInclude, slExclude, "Trade"

    gIncludeExcludeCkc RptSelCA!ckcSpots(0), slInclude, slExclude, "Missed"
    gIncludeExcludeCkc RptSelCA!ckcSpots(1), slInclude, slExclude, "Charge"
    gIncludeExcludeCkc RptSelCA!ckcSpots(2), slInclude, slExclude, "0.00"
    gIncludeExcludeCkc RptSelCA!ckcSpots(3), slInclude, slExclude, "ADU"
    gIncludeExcludeCkc RptSelCA!ckcSpots(4), slInclude, slExclude, "Bonus"
    gIncludeExcludeCkc RptSelCA!ckcSpots(5), slInclude, slExclude, "+Fill"
    gIncludeExcludeCkc RptSelCA!ckcSpots(6), slInclude, slExclude, "-Fill"
    gIncludeExcludeCkc RptSelCA!ckcSpots(7), slInclude, slExclude, "N/C"
    gIncludeExcludeCkc RptSelCA!ckcSpots(8), slInclude, slExclude, "Recap"
    gIncludeExcludeCkc RptSelCA!ckcSpots(9), slInclude, slExclude, "Spinoff"
    gIncludeExcludeCkc RptSelCA!ckcSpots(10), slInclude, slExclude, "MG"        '10-28-10

    If Len(slInclude) > 0 Then
        If Not gSetFormula("Included", "'" & slInclude & "'") Then
            gCmcGenCA = -1
            Exit Function
        End If
    End If
    If Len(slExclude) > 0 Then
        If Not gSetFormula("Excluded", "'" & slExclude & "'") Then
            gCmcGenCA = -1
            Exit Function
        End If
    End If
    
    'Date: added Net computation for schedules and rate card display
    blGross = True      'Gross computation
    If RptSelCA!rbcGrossNet(1).Value = True Then blGross = False  'Net computation

   'Sports version:  show avail Names
   'Non-sports version:  if by day (detail), by week (summary)
   If ilListIndex = AVAILSCOMBO_SPORTS Then
        If RptSelCA!ckcShowNamedAvails = vbChecked Then        'show named avails (detail version)
             If Not gSetFormula("Totals", "'D'") Then    'detail
                 gCmcGenCA = -1
                 Exit Function
             End If
        Else
             If Not gSetFormula("Totals", "'S'") Then    'summary
                 gCmcGenCA = -1
                 Exit Function
             End If
        End If
        'include multimedia
        If RptSelCA!ckcMultimedia = vbChecked Then        'show named avails (detail version)
             If Not gSetFormula("IncludeMM", "'I'") Then    'include Multimedia
                 gCmcGenCA = -1
                 Exit Function
             End If
        Else
             If Not gSetFormula("IncludeMM", "'E'") Then    'Exclude multimedia
                 gCmcGenCA = -1
                 Exit Function
             End If
        End If
    Else

        For illoop = 0 To RptSelCA!lbcSelection(1).ListCount - 1 Step 1
            If RptSelCA!lbcSelection(1).Selected(illoop) Then
                slNameCode = tgRateCardCode(illoop).sKey
                Exit For
            End If
        Next illoop
        ilRet = gParseItem(slNameCode, 2, "\", slNameCode)
        If Not gSetFormula("RCHeader", "'" & slNameCode & "'") Then
            gCmcGenCA = -1
            Exit Function
        End If

        'Major sort field
        If RptSelCA!rbcMajor(2).Value Then          'major is by day/week
            If RptSelCA!rbcTotals(0).Value Then     'totals by day
                If Not gSetFormula("MajorSortType", "'D'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            Else
                If Not gSetFormula("MajorSortType", "'W'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            End If
        ElseIf RptSelCA!rbcMajor(0).Value Then      'major is by vehicle
            If Not gSetFormula("MajorSortType", "'V'") Then    '
                    gCmcGenCA = -1
                    Exit Function
                End If
        Else                                        'major is by dayart (rate card)
            If Not gSetFormula("MajorSortType", "'R'") Then    '
                gCmcGenCA = -1
                Exit Function
            End If
        End If

         'Intermdiate sort field
        If RptSelCA!rbcInterm(2).Value Then          'intermediate is by day/week
            If RptSelCA!rbcTotals(0).Value Then     'totals by day
                If Not gSetFormula("IntermediateSortType", "'D'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            Else                                    'intermdiate is by week
                If Not gSetFormula("IntermediateSortType", "'W'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            End If
        ElseIf RptSelCA!rbcInterm(0).Value Then      'intermediate is by vehicle
            If Not gSetFormula("IntermediateSortType", "'V'") Then    '
                    gCmcGenCA = -1
                    Exit Function
                End If
        Else                                        'intermediate is by dayart (rate card)
            If Not gSetFormula("IntermediateSortType", "'R'") Then    '
                gCmcGenCA = -1
                Exit Function
            End If
        End If

        'Minor sort field
        If RptSelCA!rbcMinor(2).Value Then          'minor is by day/week
            If RptSelCA!rbcTotals(0).Value Then     'totals by day
                If Not gSetFormula("MinorSortType", "'D'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            Else                                    'minor is by week
                If Not gSetFormula("MinorSortType", "'W'") Then    '
                     gCmcGenCA = -1
                     Exit Function
                 End If
            End If
        ElseIf RptSelCA!rbcMinor(0).Value Then      'minor is by vehicle
            If Not gSetFormula("MinorSortType", "'V'") Then    '
                    gCmcGenCA = -1
                    Exit Function
                End If
        Else                                        'minor is by dayart (rate card)
            If Not gSetFormula("MinorSortType", "'R'") Then    '
                gCmcGenCA = -1
                Exit Function
            End If
        End If
        If RptSelCA!ckcNewPage.Value = vbChecked Then       'new page each vehicle
            If Not gSetFormula("NewPage", "'Y'") Then    '
                gCmcGenCA = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then    '
                gCmcGenCA = -1
                Exit Function
            End If
        End If

        'Date: 1/17/2020 set value for GrossNet formula in the report
        If Not gSetFormula("GrossNet", IIF(blGross, "'G'", "'N'")) Then
            gCmcGenCA = -1
            Exit Function
        End If
        
        'TTP 10407 - Avails Combo: Equalize by 30 and equalize by 60 column header and calculation is wrong
        If ilListIndex = AVAILSCOMBO_NONSPORTS Then
            If tgSpf.sAvailEqualize = "3" Then
                If Not gSetFormula("AvailsEqualize", """3""") Then
                    gCmcGenCA = -1
                    Exit Function
                End If
            ElseIf tgSpf.sAvailEqualize = "6" Then
                If Not gSetFormula("AvailsEqualize", """6""") Then
                    gCmcGenCA = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("AvailsEqualize", """""") Then
                    gCmcGenCA = -1
                    Exit Function
                End If
            End If
        End If
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    slSelection = slSelection & " And {GRF_Generic_Report.grfBktType} = ''"   'get game info vs multimedia info
    If Not gSetSelection(slSelection) Then
        gCmcGenCA = -1
        Exit Function
    End If
    gCmcGenCA = 1         'ok
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
    RptSelCA!frcOutput.Enabled = igOutput
    RptSelCA!frcCopies.Enabled = igCopies
    RptSelCA!frcFile.Enabled = igFile
    RptSelCA!frcOption.Enabled = igOption
    Beep
End Sub
