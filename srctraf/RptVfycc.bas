Attribute VB_Name = "RptvfyCC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptVfycc.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelCC.Bas
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
'*      Procedure Name:gCmcGenCC                       *
'*                                                     *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gcmcgenCC(ilListIndex As Integer) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'

    Dim slTime As String
    Dim slDate As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slSelection As String


    gcmcgenCC = 0

    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))

    If Not gSetSelection(slSelection) Then
        gcmcgenCC = -1
        Exit Function
    End If

    gcmcgenCC = 1         'ok
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
    RptSelCC!frcOutput.Enabled = igOutput
    RptSelCC!frcCopies.Enabled = igCopies
    'RptSelCC!frcWhen.Enabled = igWhen
    RptSelCC!frcFile.Enabled = igFile
    RptSelCC!frcOption.Enabled = igOption
    'RptSelCC!frcRptType.Enabled = igReportType
    Beep
End Sub
