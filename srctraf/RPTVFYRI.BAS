Attribute VB_Name = "RptvfyRI"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RPTVFYRI.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRI.Bas
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
'Public tgMktCode() As SORTCODE
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
Function gcmcgenRI(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    Dim slMonthHdr As String * 36
    Dim ilSaveMonth As Integer
    gcmcgenRI = 0

    slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
    slStr = RptSelRI!edcSelCFrom.Text             'month in text form (jan..dec)
    gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
    If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
        ilSaveMonth = Val(slStr)
        ilRet = gVerifyInt(slStr, 1, 12)
        If ilRet = -1 Then
            mReset
            RptSelRI!edcSelCFrom.SetFocus                 'invalid # periods
            Exit Function
        End If
    End If

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))

    If Not gSetSelection(slSelection) Then
        gcmcgenRI = -1
        Exit Function
    End If

    If ilListIndex = INV_REMOTE_INVWS Then          'REP Invoice worksheet
        'send summary or detail
        If RptSelRI!rbcSelC4(0).Value Then      'detail
            If Not gSetFormula("SummaryDetail", "'D'") Then
            gcmcgenRI = -1
            Exit Function
            End If
        Else
            If Not gSetFormula("SummaryDetail", "'S'") Then
            gcmcgenRI = -1
            Exit Function
            End If
        End If

       'send Market or advt
        If RptSelRI!rbcWSSort(0).Value Then      'market
            If Not gSetFormula("SortBy", "'M'") Then
            gcmcgenRI = -1
            Exit Function
            End If
        Else
            If Not gSetFormula("SortBy", "'A'") Then    'advertiser
            gcmcgenRI = -1
            Exit Function
            End If
        End If
    ElseIf ilListIndex = INV_BARTERPAYMENT Then         '11-28-06
        'send summary or detail
        If RptSelRI!rbcSelC4(0).Value Then      'detail
            If Not gSetFormula("TotalLevel", "'D'") Then
            gcmcgenRI = -1
            Exit Function
            End If
        Else
            If Not gSetFormula("TotalLevel", "'S'") Then
            gcmcgenRI = -1
            Exit Function
            End If
        End If
        If RptSelRI!ckcInclRecd.Value = vbChecked Then       'update barter flag
            If Not gSetFormula("BarterUpdated", "'Y'") Then
                gcmcgenRI = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("BarterUpdated", "'N'") Then
                gcmcgenRI = -1
                Exit Function
            End If
        End If
    End If

    gcmcgenRI = 1         'ok
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
    RptSelRI!frcOutput.Enabled = igOutput
    RptSelRI!frcCopies.Enabled = igCopies
    'RptSelRI!frcWhen.Enabled = igWhen
    RptSelRI!frcFile.Enabled = igFile
    RptSelRI!frcOption.Enabled = igOption
    'RptSelRI!frcRptType.Enabled = igReportType
    Beep
End Sub
