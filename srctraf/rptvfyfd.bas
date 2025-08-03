Attribute VB_Name = "RptVfyFD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyfd.bas on Wed 6/17/09 @ 12:56 P
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
Function gCmcGenFD() As Integer
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
    Dim llStart As Long
    Dim llEnd As Long
    Dim slDatesRequested As String
    Dim ilListIndex As Integer

    gCmcGenFD = 0
    ilListIndex = rptSelFD!lbcRptType.ListIndex
    If ilListIndex = PREFEED_DUMP Then
        slDate = rptSelFD!edcFromDate.Text
        If gValidDate(slDate) Then
            llStart = gDateValue(slDate)
            slDatesRequested = " as of " & Format$(llStart, "m/d/yy")
            If Not gSetFormula("DatesRequested", "'" & slDatesRequested & "'") Then
                gCmcGenFD = -1
                Exit Function
            End If

        Else
            mReset
            rptSelFD!edcFromDate.SetFocus
            Exit Function
        End If
    Else
        If ilListIndex = 0 Then             'Feed Recap
            'check validity of time input
            slTime = rptSelFD!edcFromTime.Text
            If gValidTime(slTime) Then
                llStart = gTimeToLong(slTime, False)
                slTime = rptSelFD!edcToTime.Text
                If Not gValidTime(slTime) Then
                    mReset
                    rptSelFD!edcToTime.SetFocus
                    Exit Function
                Else
                    llEnd = gTimeToLong(slTime, False)
                End If
                If llStart > llEnd Then
                    mReset
                    rptSelFD!edcToTime.SetFocus
                    Exit Function
                End If
            Else
                mReset
                rptSelFD!edcFromTime.SetFocus
                Exit Function
            End If
        End If
    
        'Feed Recap or Feed Pledges
        slDate = rptSelFD!edcFromDate.Text
        If gValidDate(slDate) Then
            llStart = gDateValue(slDate)
            slDatesRequested = "for " & Format$(llStart, "m/d/yy")
            slDate = rptSelFD!edcToDate.Text
            If Not gValidDate(slDate) Then
                mReset
                rptSelFD!edcToDate.SetFocus
                Exit Function
            Else
                llEnd = gDateValue(slDate)
                slDatesRequested = slDatesRequested & " - " & Format$(llEnd, "m/d/yy")
            End If
            If llStart > llEnd Then         'end date cant be earlier than start
                mReset
                rptSelFD!edcToDate.SetFocus
                Exit Function
            End If
        Else
            mReset
            rptSelFD!edcFromDate.SetFocus
            Exit Function
        End If
    
        If Not gSetFormula("DatesRequested", "'" & slDatesRequested & "'") Then
            gCmcGenFD = -1
            Exit Function
        End If
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenFD = -1
        Exit Function
    End If
    gCmcGenFD = 1         'ok
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
    rptSelFD!frcOutput.Enabled = igOutput
    rptSelFD!frcCopies.Enabled = igCopies
    'rptSelFD!frcWhen.Enabled = igWhen
    rptSelFD!frcFile.Enabled = igFile
    rptSelFD!frcOption.Enabled = igOption
    'rptSelFD!frcRptType.Enabled = igReportType
    Beep
End Sub


' *********************************************************************

