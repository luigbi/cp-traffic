Attribute VB_Name = "RptVfySS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyss.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelSS.Bas
'
' Release: 4.5
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
Public tgRptSelDBAgencyCode() As SORTCODE
Public tgRptSelDBSalespersonCode() As SORTCODE
Public tgRptSelDBAdvertiserCode() As SORTCODE
Public tgRptSelDBNameCode() As SORTCODE
Public tgRptSelDBBudgetCode() As SORTCODE
Public tgRptSelDBDemoCode() As SORTCODE
'
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenReportSS                              *
'*                                                             *
'*             Created:6/16/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                            *
'*         Comments: Formula setups for Crystal                *
'*                                                             *
'*          Return : 0 =  either error in input, stay in       *
'*                   -1 = error in Crystal, return to          *
'*                        calling program                      *
''*                       failure of gSetformula or another    *
'*                    1 = Crystal successfully completed       *
'*                    2 = successful Bridge                    *
'***************************************************************
Function gCmcGenSS(ilListIndex As Integer, ilGenShiftKey As Integer) As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    gCmcGenSS = 0
    gUnpackDate igStartOfWk(0), igStartOfWk(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If Not gSetFormula("StartOfWeek", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        gCmcGenSS = -1
        Exit Function
    End If
    If Not gSetFormula("MissedText", "'" & sgMissedText & "'") Then
        gCmcGenSS = -1
        Exit Function
    End If

    slSelection = ""
    slSelection = "{SWF_Spot_Week_Dump.swfurfCode} = " & igUserCode & " and {SWF_Spot_Week_Dump.swfvefCode} =" & igVehCode
    If Not gSetSelection(slSelection) Then
        gCmcGenSS = -1
        Exit Function
    End If

    gCmcGenSS = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportSS                      *
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
Function gGenReportSS() As Integer
    If Not igUsingCrystal Then
        gGenReportSS = True
        Exit Function
    End If

    If Not gOpenPrtJob("Snapshot.Rpt") Then
        gGenReportSS = False
        Exit Function
    End If
    gGenReportSS = True
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
    RptSelSS!frcOutput.Enabled = igOutput
    RptSelSS!frcCopies.Enabled = igCopies
    'RptSelDB!frcWhen.Enabled = igWhen
    RptSelSS!frcFile.Enabled = igFile
    RptSelSS!frcOption.Enabled = igOption
    'RptSelDB!frcRptType.Enabled = igReportType
    Beep
End Sub
