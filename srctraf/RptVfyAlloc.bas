Attribute VB_Name = "RPTVFYALLOC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfytx.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAlloc.Bas
'
' Release: 4.7
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenReportAlloc                              *
'*                                                             *
'*             Created:6/16/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                             *
'*         Comments: Formula setups for Crystal                *
'*                                                             *
'*          Return : 0 =  either error in input, stay in       *
'*                   -1 = error in Crystal, return to          *
'*                        calling program                      *
''*                       failure of gSetformula or another    *
'*                    1 = Crystal successfully completed       *
'*                    2 = successful Bridge                    *
'***************************************************************
Function gCmcGenAlloc() As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slTime As String
    Dim slStr As String
    Dim slStartStd As String
    Dim slEndStd As String
    Dim slWhichSort As String * 1
    
    gCmcGenAlloc = 0
    
    'Obtain the standard months start and end dates
    slStr = Trim$(str(igMonthOrQtr)) & "/15/" & Trim$(str(igYear))     'form mm/dd/yy
    slStartStd = gObtainStartStd(slStr)               'obtain std start date for month
    slEndStd = gObtainEndStd(slStr)                 'obtain std end date for month
            
    If Not gSetFormula("StdMonthRequested", "'" & slStartStd & " - " & slEndStd & "'") Then
        gCmcGenAlloc = -1
    End If
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    If RptSelALLOC!rbcSortBy(0).Value Then      'sort by Cash/Trade, Market Rank, Station
        slWhichSort = "R"
    ElseIf RptSelALLOC!rbcSortBy(1).Value Then  'sort by Cash/Trade, Market Name, Station
        slWhichSort = "N"
    ElseIf RptSelALLOC!rbcSortBy(2).Value Then  'sort by Cash/Trade, Station
        slWhichSort = "S"
    ElseIf RptSelALLOC!rbcSortBy(3).Value Then  'sory by Vehicle Group, Cash/Trade, Market Rank, Station
        slWhichSort = "R"
    ElseIf RptSelALLOC!rbcSortBy(4).Value Then  'sort by Vehicle Group, Cash/Trade, Market Name, Station
        slWhichSort = "N"
    ElseIf RptSelALLOC!rbcSortBy(5).Value Then  'sort by Vehicle Group, Cash/Trade, Station
        slWhichSort = "S"
    End If
    If Not gSetFormula("WhichSort", "'" & slWhichSort & "'") Then
        gCmcGenAlloc = -1
    End If
    
    If RptSelALLOC!rbcOrderAir(0).Value Then            'use ordered
        If Not gSetFormula("UseOrderOrAir", "'O'") Then
            gCmcGenAlloc = -1
        End If
    Else
        If Not gSetFormula("UseOrderOrAir", "'A'") Then
            gCmcGenAlloc = -1
        End If
    End If
          
    slSelection = ""
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    If Not gSetSelection(slSelection) Then
        gCmcGenAlloc = -1
        Exit Function
    End If

    gCmcGenAlloc = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportAlloc                     *
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
Function gGenReportAlloc() As Integer
Dim slMonthHdsr As String
Dim slStr As String
Dim ilSaveMonth As Integer
Dim ilRet As Integer
Dim slMonthHdr As String * 36
        gGenReportAlloc = 0
    
        slMonthHdr = "JanFebMarAprMayJunJulAugSepOctNovDec"
        slStr = RptSelALLOC!edcMonth.Text             'month in text form (jan..dec)
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
            ilRet = gVerifyInt(slStr, 1, 12)
            If ilRet = -1 Then
                mReset
                RptSelALLOC!edcMonth.SetFocus
                Exit Function
            Else
                igMonthOrQtr = ilSaveMonth
            End If
        Else
            igMonthOrQtr = Val(ilSaveMonth)           'place month in global variable
        End If

        slStr = RptSelALLOC!edcYear.Text
        ilRet = gVerifyInt(slStr, 1970, 2069)        'if month number came back 0, its invalid
        If ilRet = -1 Then
            mReset
            RptSelALLOC!edcYear.SetFocus
            Exit Function
        Else
            igYear = Val(slStr)                     'place year in global variable
        End If

        If Not gOpenPrtJob("RevAlloc.Rpt") Then
            gGenReportAlloc = False
            Exit Function
        End If
    

    gGenReportAlloc = True
End Function
Sub mReset()
    igGenRpt = False
    RptSelALLOC!frcOutput.Enabled = igOutput
    RptSelALLOC!frcCopies.Enabled = igCopies
    'RptSelAp!frcWhen.Enabled = igWhen
    RptSelALLOC!frcFile.Enabled = igFile
    RptSelALLOC!frcOption.Enabled = igOption
    'RptSelAp!frcRptType.Enabled = igReportType
    Beep
End Sub
