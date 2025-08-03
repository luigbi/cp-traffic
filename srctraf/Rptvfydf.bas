Attribute VB_Name = "RPTVFYDF"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfydf.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: rptseldf.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public lgNowTime As Long
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelPjAgencyCode() As SORTCODE
'Public sgRptSelPjAgencyCodeTag As String
'Public tgRptSelPjSalespersonCode() As SORTCODE
'Public sgRptSelPjSalespersonCodeTag As String
'Public tgRptSelPjAdvertiserCode() As SORTCODE
'Public sgRptSelPjAdvertiserCodeTag As String
'Public tgRptSelPjNameCode() As SORTCODE
'Public sgRptSelPjNameCodeTag As String
'Public tgRptSelPjBudgetCode() As SORTCODE
'Public sgRptSelPjBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelPjDemoCode() As SORTCODE
'Public sgRptSelPjDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public Const DF_SETNAMES = 0
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
'Library calendar file- used to obtain post log date status
'Not used
'Dim lmStartDates() As Long          'array of 13 bdcst or corp start dates
'Dim lmEndDates() As Long            'array of 13 bdcst or corp end dates
'
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenGenDF                         *
'*                                                     *
'*             Created:7/24/97       By:W. Bjerke      *
'*            Modified:              By:               *
'*                                                     *
'*         Comments: Duplicate of gGenGenDF due to       *
'*                   expanding code base.              *
'*                   Used for projections              *
'*******************************************************
Function gGenGenDF(ilListIndex As Integer, ilGenShiftKey As Integer, slLogUserCode As String) As Integer
    Dim ilLoop As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slTime As String
'    ReDim lmStartDates(1 To 13) As Long
'    ReDim lmEndDates(1 To 13) As Long
gGenGenDF = 0
'11/04/20 - TTP # 10014 - Cleanup AsOfT (pt2)
'gCurrDateTime slStr, slTime, slMonth, slDay, slYear
'If Not gSetFormula("AsOfT", Trim$(str$(CLng(gTimeToCurrency(slTime, False))))) Then
'    gGenGenDF = -1
'    Exit Function
'End If
slSelection = ""
If Not RptSelDF!ckcAll.Value = vbChecked Then
    For ilLoop = 0 To RptSelDF!lbcSelection(0).ListCount - 1 Step 1
        If RptSelDF!lbcSelection(0).Selected(ilLoop) = True Then
            If slSelection <> "" Then
                slSelection = slSelection & " Or " & "{SNF_Set_Name.snfName} = " & "'" & Trim$(RptSelDF!lbcSelection(0).List(ilLoop)) & "'"
            Else
                slSelection = "{SNF_Set_Name.snfName} = " & "'" & Trim$(RptSelDF!lbcSelection(0).List(ilLoop)) & "'"
            End If
        End If
    Next ilLoop
End If
If Not gSetSelection(slSelection) Then
    gGenGenDF = -1
    Exit Function
End If

gGenGenDF = 1
Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportDF                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:8/14/97       By:W. Bjerke      *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportDF() As Integer
    Dim ilListIndex As Integer
    ilListIndex = RptSelDF!lbcRptType.ListIndex

    'If illistindex = PRJ_SALESPERSON Then
    '    If Not gOpenPrtJob("pjsls.Rpt") Then
    '        gGenReportDF = False
    '        Exit Function
    '    End If
    'ElseIf illistindex = PRJ_VEHICLE Or illistindex = PRJ_OFFICE Or illistindex = PRJ_CATEGORY Then
    '    If Not gOpenPrtJob("pjvehofc.Rpt") Then
    '        gGenReportDF = False
    '        Exit Function
    '    End If
    ''ElseIf illistindex = PRJ_POTENTIAL Then
    '    If Not gOpenPrtJob("pjpot.Rpt") Then
    '        gGenReportDF = False
    '        Exit Function
    '    End If
    'End If

    If ilListIndex = 0 Then
        If Not gOpenPrtJob("rptsetdf.Rpt") Then
            gGenReportDF = False
            Exit Function
        End If
    End If

    gGenReportDF = True
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
    RptSelDF!frcOutput.Enabled = igOutput
    RptSelDF!frcCopies.Enabled = igCopies
    'rptseldf!frcWhen.Enabled = igWhen
    RptSelDF!frcFile.Enabled = igFile
    RptSelDF!frcOption.Enabled = igOption
    'rptseldf!frcRptType.Enabled = igReportType
    Beep
End Sub
'
'
'           mVerifyDate - verify date entered as valid
'           <input> edcDate - control field (edit box) containing date string
'           <output> llDate - date converted as Long
'           <return> 0 = OK, 1 = invalid date entered
'
Function mVerifyDate(edcDate As Control, llValidDate As Long) As Integer
Dim slDate As String
    llValidDate = 0
    mVerifyDate = 0
    slDate = edcDate.Text
    If (slDate <> "") Then              'date isnt reqd
        If Not gValidDate(slDate) Then
            mReset
            edcDate.SetFocus
            mVerifyDate = -1
            Exit Function
        Else
            llValidDate = gDateValue(slDate)
        End If
    End If
End Function
