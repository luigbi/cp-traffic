Attribute VB_Name = "RPTVFYRS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfyrs.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRS.Bas
'
' Release: 4.5
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Function gGenReportRS(ilListIndex As Integer, hlDnf As Integer, tlDnf As DNF) As Integer
Dim slWhichRpt As String
Dim ilLoop As Integer
Dim ilRet As Integer
Dim ilDnfCode As Integer
Dim slCode As String
Dim slNameCode As String

    If ilListIndex = RS_SUMMARY Then                'regular research summary

'        'Determine if 18 buckets (new format) or 16 buckets
'        For ilLoop = 0 To RptSelRS!lbcSelection(1).ListCount - 1 Step 1
'            If RptSelRS!lbcSelection(1).Selected(ilLoop) Then
'                slNameCode = tgBookNameCode(ilLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                ilDnfCode = Val(slCode)
'                ilRet = btrGetEqual(hlDnf, tlDnf, Len(tlDnf), ilDnfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                If ilRet <> BTRV_ERR_NONE Then
'                    MsgBox "Invalid Book Name Selected - RptSelRS (cmcgen)"
'                    Exit Function
'                Else
'                    Exit For
'                End If
'           End If
'        Next ilLoop
'        'dan M 12/21/10
'        If RptSelRS!ckcSelC3(13).Value = vbChecked Or tlDnf.sForm = "8" Then
'            slWhichRpt = "Resear18.rpt"
'        Else
'            slWhichRpt = "Research.rpt"
'        End If

        slWhichRpt = "Resear18.rpt"         '5-20-14 always use the 18 demo version of the report; adjust the 16 demo records to align with 18 demo columns
        If Not gOpenPrtJob(slWhichRpt) Then
            gGenReportRS = False
            Exit Function
        End If
    ElseIf ilListIndex = RS_SPECIALSUMMARY Then         'special subscriber summary
        If Not gOpenPrtJob("SpecResearch.Rpt") Then
            gGenReportRS = False
            Exit Function
        End If
    ElseIf ilListIndex = RS_DEMORANK Then
        If Not gOpenPrtJob("DemoRank.Rpt") Then
            gGenReportRS = False
            Exit Function
        End If
    End If

    gGenReportRS = True
End Function
'********************************************************
'*                                                      *
'*      Procedure Name:gGenReportRS                     *
'       <input> Report index selected
'
'       Send formulas to report
'*                                                      *
'********************************************************
Function gCmcGenRS(ilListIndex As Integer) As Integer
'
'
'
'   ilRet (O)-  -1=Terminate
'               0 = Exit sub
'               1 = Ok
'
    Dim slTime As String
    Dim slDate As String
    Dim slSelection As String

    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilTemp As Integer
    Dim ilPrimaryIndex As Integer
    Dim slPrimaryDemo As String
    
    gCmcGenRS = 0

    If ilListIndex = RS_SUMMARY Then            'Research summary
        If (RptSelRS!ckcSelC3(0).Value = vbChecked) Then    'Show Column Headings
            If Not gSetFormula("ShowSocEco1", "'Soc'") Then
                gCmcGenRS = -1
                Exit Function
            End If
            If Not gSetFormula("ShowSocEco2", "'Eco'") Then
                gCmcGenRS = -1
                Exit Function
            End If
        Else                         'Don't Show Column Headings
            If Not gSetFormula("ShowSocEco1", "''") Then
                gCmcGenRS = -1
                Exit Function
            End If
            If Not gSetFormula("ShowSocEco2", "''") Then
                gCmcGenRS = -1
                Exit Function
            End If
        End If
        
        '8-25-16 option to skip to new page each book
        If (RptSelRS!ckcNewPage.Value = vbChecked) Then    'Show Column Headings
            If Not gSetFormula("NewPage", "'Y'") Then
                gCmcGenRS = -1
                Exit Function
            End If
        Else
            If Not gSetFormula("NewPage", "'N'") Then
                gCmcGenRS = -1
                Exit Function
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
        slSelection = "{RSR_Research_Data.rsrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({RSR_Research_Data.rsrGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenRS = -1
            Exit Function
        End If
    ElseIf ilListIndex = RS_SPECIALSUMMARY Then
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "({GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False)))) & ")"
    ElseIf ilListIndex = RS_DEMORANK Then
        'top how many
        If RptSelRS!edcTopHowMany.Text = "" Then
            ilTemp = 0
        Else
            ilTemp = Val(RptSelRS!edcTopHowMany.Text)
        End If
        If Not gSetFormula("TopHowMany", ilTemp) Then
            gCmcGenRS = -1
            Exit Function
        End If
        
        ilPrimaryIndex = RptSelRS!cbcPrimaryDemo.ListIndex
        slPrimaryDemo = Trim$(RptSelRS!cbcPrimaryDemo.List(ilPrimaryIndex))
        If Not gSetFormula("PrimaryDemo", "'" & slPrimaryDemo & "'") Then
            gCmcGenRS = -1
            Exit Function
        End If
        
        If Not gSetFormula("SiteAudData", "'" & tgSpf.sSAudData & "'") Then
            gCmcGenRS = -1
            Exit Function
        End If
        
        
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'filter for GRf on matching generated date & time
        slSelection = "{RSR_Research_Data.rsrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({RSR_Research_Data.rsrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            gCmcGenRS = -1
            Exit Function
        End If
    
    End If

    gCmcGenRS = 1         'ok
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReset                          *
'*                                                     *
'*             Created:1/31/96       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Reset controls                 *
'*                                                     *
'*******************************************************
Sub mReset()
    igGenRpt = False
    RptSelRS!frcOutput.Enabled = igOutput
    RptSelRS!frcCopies.Enabled = igCopies
    'RptSelRS!frcWhen.Enabled = igWhen
    RptSelRS!frcFile.Enabled = igFile
    RptSelRS!frcOption.Enabled = igOption
    'RptSelRS!frcRptType.Enabled = igReportType
    'RPTSelRS!edcDates.SetFocus
    Beep
End Sub
