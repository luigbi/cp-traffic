Attribute VB_Name = "RptvfyRg"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRg.Bas
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
'*      Procedure Name:gCmcGenRg                       *
'*                                                     *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Formula setups for Crystal        *
'*                                                     *
'*******************************************************
Function gcmcgenRg() As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'

    Dim slTime As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slSelection As String
    Dim slDate As String


    gcmcgenRg = 0

'   8-23-19 use csi calendar control vs edit box
'    If RptSelRg!edcStartDate.Text <> "" Then
'        slStartDate = RptSelRg!edcStartDate.Text
    If RptSelRg!CSI_calStart.Text <> "" Then
        slStartDate = RptSelRg!CSI_calStart.Text

        If Not gValidDate(slStartDate) Then
            mReset
            RptSelRg!CSI_calStart.SetFocus
            Exit Function
        End If
    End If
    llStartDate = gDateValue(slStartDate)
'    If RptSelRg!edcEndDate.Text <> "" Then
'        If StrComp(RptSelRg!edcEndDate.Text, "TFN", 1) <> 0 Then
'            slEndDate = RptSelRg!edcEndDate.Text
    If RptSelRg!CSI_CalEnd.Text <> "" Then
        If StrComp(RptSelRg!CSI_CalEnd.Text, "TFN", 1) <> 0 Then
            slEndDate = RptSelRg!CSI_CalEnd.Text

            If Not gValidDate(slEndDate) Then
                mReset
                RptSelRg!CSI_CalEnd.SetFocus
                Exit Function
            End If
        End If
    End If
    llEndDate = gDateValue(slEndDate)
    If llStartDate > llEndDate Then
        mReset
        RptSelRg!CSI_calStart.SetFocus
        Exit Function
    End If
    
    slStartDate = Format$(llStartDate, "m/d/yy")
    slEndDate = Format$(llEndDate, "m/d/yy")
    If Not gSetFormula("DatesRequested", "'" & slStartDate & "-" & slEndDate & "'") Then
        gcmcgenRg = 0
        Exit Function
    End If
    
    gCurrDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

    If Not gSetSelection(slSelection) Then
        gcmcgenRg = 0
        Exit Function
    End If

    gcmcgenRg = 1         'ok
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
    RptSelRg!frcOutput.Enabled = igOutput
    RptSelRg!frcCopies.Enabled = igCopies
    'RptSelRg!frcWhen.Enabled = igWhen
    RptSelRg!frcFile.Enabled = igFile
    RptSelRg!frcOption.Enabled = igOption
    'RptSelRg!frcRptType.Enabled = igReportType
    Beep
End Sub
