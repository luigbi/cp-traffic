Attribute VB_Name = "RptVfyID"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptvfyid.bas on Wed 6/17/09 @ 12:56 P
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
Function gcmcgenID(ilGenShiftKey As Integer, slLogUserCode As String) As Integer
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
    gcmcgenID = 0

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))

    If Not gSetSelection(slSelection) Then
        gcmcgenID = -1
        Exit Function
    End If

    If RptSelID!ckcDiscrepOnly.Value = vbChecked Then
        If Not gSetFormula("DiscrepOnly", "'" & "Discrepancy Contracts for " & "'") Then
            gcmcgenID = -1
            Exit Function
        End If
    Else
        If Not gSetFormula("DiscrepOnly", "'" & "All Contracts for " & "'") Then
            gcmcgenID = -1
            Exit Function
        End If
    End If

    gcmcgenID = 1         'ok
    Exit Function
End Function
'*********************************************************************
'
'               gCurrDateTime
'               <Output> slDate - current date (xx/xx/xx)
'                        slTime - current time (xx:xx:xxa/p)
'                        Some routines may not use these return values
'                        slMonth - xx  (1-12)
'                        slDay - XX  (1-31)
'                        slYear - xxxx (19xx-20xx)
'               obtain system current date and time and return it
'               in string format
'
'               Created:  7/3/96
'*********************************************************************

'***************************************************************
'
'           gUnpackCurDateTime - from igNowDate & igNowTime
'           convert to string
'           <output> slCurrDate xx/xx/xxxx
'                    slCurrTime xx:xx:xxa/p
'                    slMonth as string
'                    slDay as string
'                    slYear as string (xxxx)
'***************************************************************
'Sub gUnpackCurrDateTime(slCurrDate As String, slCurrTime As String, slMonth As String, slDay As String, slYear As String)
'    gUnpackDate igNowDate(0), igNowDate(1), slCurrDate
'    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slCurrTime
'    gObtainYearMonthDayStr slCurrDate, True, slYear, slMonth, slDay
'End Sub
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
