Attribute VB_Name = "RPTVFYOA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptvfyOA.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelOA.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

'
'***************************************************************
'*                                                             *
'*      Procedure Name:gGenOrderAudit                          *
'*                                                             *
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
Function gCmcGenOA() As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slTime As String

    gCmcGenOA = 0
    slSelection = ""
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    gUnpackDate igNowDate(0), igNowDate(1), slDate
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime    '10-20-01
    slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
    slSelection = slSelection & " and {CBF_Contract_BR.cbfRdfDPSort} = 0"
        If Not gSetSelection(slSelection) Then
                gCmcGenOA = -1
                Exit Function
        End If

    gCmcGenOA = 1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportOA                    *
'*                                                     *
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportOA() As Integer
    If Not igUsingCrystal Then
        gGenReportOA = True
        Exit Function
    End If

    If Not gOpenPrtJob("OrderAudit.Rpt") Then
            gGenReportOA = False
            Exit Function
    End If
    gGenReportOA = True
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
'Sub mReset()
'    igGenRpt = False
'    RptSelOA!frcOutput.Enabled = igOutput
'    RptSelOA!frcCopies.Enabled = igCopies
'    'RptSelOA!frcWhen.Enabled = igWhen
'    RptSelOA!frcFile.Enabled = igFile
'    RptSelOA!frcOption.Enabled = igOption
'    'RptSelOA!frcRptType.Enabled = igReportType
'    Beep
'End Sub
