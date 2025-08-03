Attribute VB_Name = "rptvfyNT"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfynt.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  mVerifyIntNT                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptVfyNT.Bas
'
' Release: 5.1
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

'
'*******************************************************
'*                                                     *
'*      Procedure Name:gCmcGen                         *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*
'*          Return : 0 =  either error in input, stay in
'*                   -1 = error in Crystal, return to
'*                        calling program
''*                       failure of gSetformula or another
'*                    1 = Crystal successfully completed
'*                    2 = successful Bridge
'*******************************************************
Function gCmcGenNT(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                       ilTemp                                                  *
'******************************************************************************************

    Dim slSelection As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilMonth As Integer
    Dim slInclude As String
    Dim slExclude As String
    Dim ilSaveMonth As Integer
    Dim ilMajor As Integer
    Dim ilMinor As Integer
    Dim slMonthInYear As String * 36
    Dim ilStartQtr As Integer
    Dim ilYear As Integer

        gCmcGenNT = 0
        slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"

        If ilListIndex = CNT_NTRRECAP Then

'            If RptSelNT!edcDate1.Text = "" Then
            If RptSelNT!CSI_CalFrom.Text = "" Then
                slStartDate = "1/1/1970"
            Else
                slStartDate = RptSelNT!CSI_CalFrom.Text
            End If
            If gValidDate(slStartDate) Then         'the start date is valid, continue to test end date
'                If RptSelNT!edcDate2.Text = "" Then
                If RptSelNT!csi_CalTo.Text = "" Then
                    slEndDate = "12/31/2069"
                Else
                    slEndDate = RptSelNT!csi_CalTo.Text
                End If
                If Not gValidDate(slEndDate) Then
                    mReset
                    RptSelNT!csi_CalTo.SetFocus
                    Exit Function
                End If
            Else
                mReset
                RptSelNT!CSI_CalFrom.SetFocus
                Exit Function
            End If
            'make sure end date is greater than end date
            If gDateValue(slEndDate) < gDateValue(slStartDate) Then
                mReset
                RptSelNT!csi_CalTo.SetFocus
                Exit Function
            End If

            If RptSelNT!ckcShowDescr.Value = vbChecked Then     'show NTR description
                If Not gSetFormula("ShowDescr", "'" & "Y" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("ShowDescr", "'" & "N" & "'") Then
                        gCmcGenNT = -1
                        Exit Function
                End If
            End If
            If RptSelNT!ckcSkipPage.Value = vbChecked Then     'skip to new page each new group
                If Not gSetFormula("SkipPage", "'" & "Y" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SkipPage", "'" & "N" & "'") Then
                        gCmcGenNT = -1
                        Exit Function
                End If
            End If
            slInclude = ""
            slExclude = ""
            If RptSelNT!rbcTotalsBy(2).Value = True Then        'include both billed & unbilled
                slInclude = "Billed, Unbilled"
            Else
                gIncludeExcludeRbc RptSelNT!rbcTotalsBy(0), slInclude, slExclude, "Billed Only"
                gIncludeExcludeRbc RptSelNT!rbcTotalsBy(1), slInclude, slExclude, "UnBilled Only"
            End If
            'If RptSelNT!ckcInclHardCost.Value = vbUnchecked Then
                gIncludeExcludeCkc RptSelNT!ckcInclHardCost, slInclude, slExclude, "Hard Cost"
            'End If
            If Not gSetFormula("Include", "'" & slInclude & "'") Then
                gCmcGenNT = -1
                Exit Function
            End If
            If Not gSetFormula("Exclude", "'" & slExclude & "'") Then
                gCmcGenNT = -1
                Exit Function
            End If
            
            '1-14-10 Use Acq Cost Only
            If RptSelNT!ckcUseAcqCost.Value = vbChecked Then     'Use Acq cost only
                If Not gSetFormula("UseAcqCost", "'" & "Y" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("UseAcqCost", "'" & "N" & "'") Then
                        gCmcGenNT = -1
                        Exit Function
                End If
            End If

        Else            'NTR/Multimedia Billed & Booked

            '7-3-08 swap controls using start qtr & year to year & start month
            slStr = RptSelNT!edcDate1.Text                 'starting year
            igYear = mVerifyYear(slStr)
            If igYear = 0 Then
                mReset
                RptSelNT!edcDate1.SetFocus                 'invalid year
                gCmcGenNT = False
                Exit Function
            End If

            slStr = RptSelNT!edcDate2.Text            'starting period
            igMonthOrQtr = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                mReset
                RptSelNT!edcDate2.SetFocus
                Exit Function
            End If

            slStr = RptSelNT!edcPeriods.Text            '#periods
            ilMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                mReset
                RptSelNT!edcPeriods.SetFocus
                Exit Function
            End If
            If Not gSetFormula("NumberPeriods", ilMonth) Then
                gCmcGenNT = -1
                Exit Function
            End If

            'following formulas sent to Crystal to monthly headings
'            slStr = ""
'            ilValue = Val(RptSelNT!edcDate1.Text)
'            If ilValue = 1 Then
'                slStr = "1st Quarter "
'            ElseIf ilValue = 2 Then
'                slStr = "2nd Quarter "
'            ElseIf ilValue = 3 Then
'                slStr = "3rd Quarter "
'            Else
'                slStr = "4th Quarter "
'            End If

'           quarter headings have to be calculated based on the month entered, see below
'            slStr = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)  'pass starting month and year for header
'                                                                        '7-3-08 changed from start qtr to start month
'            slStr = slStr & " " & RptSelNT!edcDate1.Text           'year

            ilStartQtr = gGetQtrForColumns(igMonthOrQtr)   'column headings for quarter totals

            If RptSelNT!rbcBillBy(0).Value Then             'corp
                slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year

                ilYear = gGetYearofCorpMonth(igMonthOrQtr, igYear)
                'changed from starting qtr to allow for starting month and to show in report header
                If Not gSetFormula("MonthHeader", "'" & slMonth & " " & str$(ilYear) & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
                gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)
                'pass starting month for requested report for column headings
                If Not gSetFormula("StartingMonth", ilSaveMonth) Then
                    gCmcGenNT = -1
                    Exit Function
                End If

                If Not gSetFormula("CorpStd", "'C'") Then
                     gCmcGenNT = -1
                     Exit Function
                 End If
            Else                    '8-13-19 cal or std , format month & year (Jan 2007) for report heading
                slStr = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)  'pass starting month and year for header
                slStr = slStr & " " & RptSelNT!edcDate1.Text           'year
                If Not gSetFormula("MonthHeader", "'" & slStr & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If

                'pass starting month for requested report for columm headings
                If Not gSetFormula("StartingMonth", igMonthOrQtr) Then
                    gCmcGenNT = -1
                    Exit Function
                End If

                If RptSelNT!rbcBillBy(1) Then
                    If Not gSetFormula("CorpStd", "'S'") Then       'std
                        gCmcGenNT = -1
                        Exit Function
                    End If
                Else
                    If Not gSetFormula("CorpStd", "'A'") Then       'calendar
                        gCmcGenNT = -1
                        Exit Function
                    End If
                End If
            End If


            If Not gSetFormula("QtrHeader", "'" & Trim$(str(ilStartQtr)) & "'") Then
                gCmcGenNT = -1
                Exit Function
            End If

                 '3-27-02 option for totals by contract or advt
            If RptSelNT!rbcTotalsBy(0).Value Then           'totals by contract
                If Not gSetFormula("TotalsBy", "'" & "C" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            ElseIf RptSelNT!rbcTotalsBy(1).Value Then          'totals by Advt
                If Not gSetFormula("TotalsBy", "'" & "A" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("TotalsBy", "'" & "S" & "'") Then        'summary
                    gCmcGenNT = -1
                    Exit Function
                End If
            End If

            If RptSelNT!rbcGrossNet(0).Value = True Then     'Gross or Net for header
                If Not gSetFormula("GrossNet", "'" & "G" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("GrossNet", "'" & "N" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            End If

            If RptSelNT!ckcSkipPage.Value = vbChecked Then     'skip to new page each new group
                If Not gSetFormula("SkipPage", "'" & "Y" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SkipPage", "'" & "N" & "'") Then
                        gCmcGenNT = -1
                        Exit Function
                End If
            End If

            'Send which major/minor sorts to use
            ilMajor = RptSelNT!cbcSet1.ListIndex
            If Not gSetFormula("MajorSort", ilMajor) Then
                gCmcGenNT = -1
                Exit Function
            End If
            ilMinor = RptSelNT!cbcSet2.ListIndex
            If ilMinor = 0 Or ilMajor = ilMinor - 1 Then  'doesnt make sense for the minor and major to be save sorts
                ilMinor = 9                 'let Crystal know there no minor sort
            Else
                ilMinor = ilMinor - 1       '0 from list box is NONE
            End If
            If Not gSetFormula("MinorSort", ilMinor) Then
                gCmcGenNT = -1
                Exit Function
            End If
            gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
            If Trim$(slStr) = "" Then       'no last invoiced date yet, startup
                slStr = "1/1/1975"
            End If
            gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
            If Not gSetFormula("LastBilled", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                gCmcGenNT = -1
                Exit Function
            End If
            
            If (ilMajor = 4 And RptSelNT!ckcMajorSplit.Value = vbChecked) Or (ilMinor = 4 And RptSelNT!ckcMinorSplit.Value = vbChecked) Then          '9-24-19
                If Not gSetFormula("SlspSplit", "'" & "Y" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            Else
                If Not gSetFormula("SlspSplit", "'" & "N" & "'") Then
                    gCmcGenNT = -1
                    Exit Function
                End If
            End If
        End If


        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenNT = -1
            Exit Function
        End If

    gCmcGenNT = 1
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportNT(ilListIndex As Integer) As Integer

    If ilListIndex = CNT_NTRRECAP Then
        If Not gOpenPrtJob("NTRRecap.Rpt") Then
            gGenReportNT = False
            Exit Function
        End If
    ElseIf ilListIndex = CNT_NTRBB Then
        If Not gOpenPrtJob("NTRBB.Rpt") Then
            gGenReportNT = False
            Exit Function
        End If
    ElseIf ilListIndex = CNT_MULTIMBB Then          'multimedia B & B
        If Not gOpenPrtJob("MultiMBB.Rpt") Then
            gGenReportNT = False
            Exit Function
        End If
    End If

    gGenReportNT = True
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
    RptSelNT!frcOutput.Enabled = igOutput
    RptSelNT!frcCopies.Enabled = igCopies
    'RptSelNT!frcWhen.Enabled = igWhen
    RptSelNT!frcFile.Enabled = igFile
    RptSelNT!frcOption.Enabled = igOption
    'RptSelNT!frcRptType.Enabled = igReportType
    Beep
End Sub

'
'               mVerifyIntNT - verify input.  Value must be between two arguments provided
'               <input>     slStr - user input
'                           ilLowInt - lowest value allowed
'                           ilHiInt - highest value allowed
'               <output>    Return - converted integer
'                                    -1 if invalid
Function mVerifyIntNT(slStr As String, ilLowInt As Integer, ilHiInt As Integer) As Integer 'VBC NR
Dim ilInput As Integer 'VBC NR
    mVerifyIntNT = 0 'VBC NR
    ilInput = Val(slStr) 'VBC NR
    If (ilInput < ilLowInt) Or (ilInput > ilHiInt) Then 'VBC NR
        mVerifyIntNT = -1 'VBC NR
    Else 'VBC NR
        mVerifyIntNT = ilInput 'VBC NR
    End If 'VBC NR
End Function 'VBC NR
'
'               mVerifyYear - Verify that the year entered is valid
'                             If not 4 digit year, add 1900 or 2000
'                             Valid year must be > than 1950 and < 2050
'                             Input - string containing input
'                             Output - Integer containing year else 0
'
Private Function mVerifyYear(slStr As String) As Integer
Dim ilInput As Integer
    mVerifyYear = 0
    If IsNumeric(slStr) Then
        ilInput = Val(slStr)
        If ilInput < 100 Then           'only 2 digit year input ie.  96, 95,
            If ilInput < 50 Then        'adjust for year 1900 or 2000
                ilInput = 2000 + ilInput
            Else
                ilInput = 1900 + ilInput
            End If
        End If
        If (ilInput < 1950) Or (ilInput > 2050) Then
            mVerifyYear = 0
        Else
            mVerifyYear = ilInput
        End If
    End If
End Function
