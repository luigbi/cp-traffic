Attribute VB_Name = "rptvfyrv"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RPTVFYRV.BAS on Wed 6/17/09 @ 12:56 P
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
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelAgencyCode() As SORTCODE
'Public sgRptSelAgencyCodeTag As String
'Public tgRptSelSalespersonCode() As SORTCODE
'Public sgRptSelSalespersonCodeTag As String
'Public tgRptSelAdvertiserCode() As SORTCODE
'Public sgRptSelAdvertiserCodeTag As String
'Public tgRptSelNameCode() As SORTCODE
'Public sgRptSelNameCodeTag As String
'Public tgRptSelBudgetCode() As SORTCODE
'Public sgRptSelBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelDemoCode() As SORTCODE
'Public sgRptSelDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public tgBookName() As SORTCODE
'Public sgBookNameTag As String

'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public igRCSelectedIndex As Integer         'selected r/c index
'Public igBSelectedIndex As Integer          'selected budget Plan index
'Public igBFCSelectedIndex As Integer        'selected budget forecast index
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer

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
Function gCmcGenRV(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    gCmcGenRV = 0
           'dan M 6-30-08 NTR/Hard Cost option
    slStr = ""
    If RptSelRV!ckcSelCInclude(0).Value = 1 Or RptSelRV!ckcSelCInclude(1).Value = 1 Then
        slStr = slStr & "'With"
        If RptSelRV!ckcSelCInclude(0).Value = 1 Then
            slStr = slStr & " NTR"
            If RptSelRV!ckcSelCInclude(1).Value = 1 Then
                slStr = slStr & " and Hard Cost"
            End If
        Else
            slStr = slStr & " Hard Cost"
        End If
    End If
    If slStr <> "" Then
        slStr = slStr & "'"
    End If
    If Not gSetFormula("reporttitle", slStr) Then
        gCmcGenRV = False
        Exit Function
    End If

    slDate = RptSelRV!CSI_CalFrom.Text  'Date: 10/28/2019 using CSI calendar for date entry --> edcSelCFrom.Text
    If Not gValidDate(slDate) Then
        mReset
        RptSelRV!CSI_CalFrom.SetFocus   'Date: 10/28/2019 using CSI calendar for date entry --> edcSelCFrom.SetFocus
        Exit Function
    End If
    slStr = RptSelRV!edcSelCTo.Text                 'entered year
    igYear = gVerifyYear(slStr)
    If igYear = 0 Then
        mReset
        RptSelRV!edcSelCTo.SetFocus                 'invalid year
        gCmcGenRV = -1
        Exit Function
    End If
    slStr = RptSelRV!edcSelCTo1.Text                  'edit qtr
    ilRet = gVerifyInt(slStr, 1, 4)
    If ilRet = -1 Then
        mReset
        RptSelRV!edcSelCTo1.SetFocus                 'invalid qtr
        gCmcGenRV = -1
        Exit Function
    End If
    igMonthOrQtr = Val(slStr)                       'put qtr entered in global variable
    If Not mWeekQtrHdr(slDate) Then           'pass year/month as formula to crystal report
        gCmcGenRV = -1
        Exit Function
    End If
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenRV = -1
        Exit Function
    End If
    gCmcGenRV = 1         'ok
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
    RptSelRV!frcOutput.Enabled = igOutput
    RptSelRV!frcCopies.Enabled = igCopies
    'RptSelRV!frcWhen.Enabled = igWhen
    RptSelRV!frcFile.Enabled = igFile
    RptSelRV!frcOption.Enabled = igOption
    'RptSelRV!frcRptType.Enabled = igReportType
    Beep
End Sub


' *********************************************************************
'
'            mWeekQtrHdr()
'           Setup Week and Quarter Header and send to Crystal reports
'           as a formula named WeekQtrHeader
'           Format output as Quarter X Year XXXX
'           <output> slDate - First date of quarter
'           Created:  7/3/96
'
'***********************************************************************
Function mWeekQtrHdr(slDate As String) As Integer
Dim ilMonth As Integer
Dim ilYear As Integer
Dim slStr As String

        mWeekQtrHdr = True
        ilYear = RptSelRV!edcSelCTo.Text                'starting year
        ilMonth = RptSelRV!edcSelCTo1.Text              'month
        slDate = Trim$(str$(((ilMonth - 1) * 3 + 1))) & "/15/" & Trim$(str$(ilYear))
        slDate = gObtainStartStd(slDate)
        If ilMonth = 1 Then
            slStr = "1st"
        ElseIf ilMonth = 2 Then
            slStr = "2nd"
        ElseIf ilMonth = 3 Then
            slStr = "3rd"
        Else
            slStr = "4th"
        End If
        slStr = slStr & " Quarter" & str$(ilYear)      'add Year
         If Not gSetFormula("WeekQtrHeader", "'" & slStr & "'") Then
             mWeekQtrHdr = False
             Exit Function
         End If
End Function
