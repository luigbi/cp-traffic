Attribute VB_Name = "ReportListSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budget.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ReportList.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text

'Used to get current setting of color, vertical and horizontal resoluation
Public Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long
    
    

Public sgReportFormExe As String
Public igReportRnfCode As Integer
Public sgReportCtrlSaveName As String
Public fgReportForm As Form
Public igReportButtonIndex As Integer '0=Restore All; 1= Restore All except Dates; 2=Use default





'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterModalForm                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center modal form within        *
'*                     Traffic Form                    *
'*                                                     *
'*******************************************************
Sub gCenterModalForm(FrmName As Form)
'
'   gCenterModalForm FrmName
'   Where:
'       FrmName (I)- Name of modal form to be centered within Traffic form
'
        gCenterStdAlone FrmName
    'Dim flLeft As Single
    'Dim flTop As Single
    'flLeft = Traffic.Left + (Traffic.Width - Traffic.ScaleWidth) / 2 + (Traffic.ScaleWidth - FrmName.Width) / 2
    'flTop = Traffic.Top + (Traffic.Height - FrmName.Height + 2 * Traffic.cmcTask(0).Height - 60) / 2 + Traffic.cmcTask(0).Height
    'FrmName.Move flLeft, flTop
End Sub

Public Sub gSaveReportCtrlsSetting()
    Dim ilRet As Integer
    ReDim slBypassCtrls(0 To 0) As String
    
'    If StrComp(sgReportFormExe, "RptNoSel", 1) = 0 Then
'        'RptNoSel.Show vbModal
'        ilRet = gCreateFormControlFile(RptNoSel, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSel", 1) = 0 Then
'        'RptSel.Show vbModal
'        ilRet = gCreateFormControlFile(RptSel, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSel30", 1) = 0 Then           '6-12-13 cpp/cpm 30"unit
'        'RptSel30.Show vbModal
'        ilRet = gCreateFormControlFile(RptSel30, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelaa", 1) = 0 Then
'        'RptSelAA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelac", 1) = 0 Then
'        'RptSelAc.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAc, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelAcqPay", 1) = 0 Then       '8-5-15
'        'RptSelAcqPay.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAcqPay, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelad", 1) = 0 Then
'        'RptSelAD.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAD, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelal", 1) = 0 Then       '4-14-04
'        'RptSelAL.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAL, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelalloc", 1) = 0 Then       '12-13-18   Revenue Allocation
'        'RptSelALLOC.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelALLOC, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelap", 1) = 0 Then
'        'RptSelAp.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAp, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelas", 1) = 0 Then
'        'RptSelAS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelav", 1) = 0 Then
'        'RptSelAv.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAv, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelBO", 1) = 0 Then       '7-22-11 Sales Breakout
'        'RptSelBO.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelBO, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelcb", 1) = 0 Then
'        'RptSelCb.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCb, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelcc", 1) = 0 Then       '1-15-04 Producer/Provider reports
'        'RptSelCC.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCC, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelcm", 1) = 0 Then       '10-5-10 Competitive Categories
'        'RptSelCM.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCM, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelcp", 1) = 0 Then
'        'RptSelCp.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCp, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelCt", 1) = 0 Then
'        'RptSelCt.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCt, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSeldb", 1) = 0 Then
'        'RptSelDB.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelDB, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSeldf", 1) = 0 Then
'        'RptSelDF.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelDF, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelds", 1) = 0 Then
'        'RptSelDS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelDS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelFd", 1) = 0 Then       '8-4-04 Feed Report
'        'rptSelFD.Show vbModal
'        ilRet = gCreateFormControlFile(rptSelFD, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelia", 1) = 0 Then
'        'RptSelIA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelIA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelid", 1) = 0 Then       '5-21-02
'        'RptSelID.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelID, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelin", 1) = 0 Then
'        ''RptSelIn.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSelir", 1) = 0 Then       '7-13-05
'        'RptSelIR.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelIR, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSeliv", 1) = 0 Then
'        'RptSelIv.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelIv, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSellg", 1) = 0 Then
'        ''RptSellg.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSelNT", 1) = 0 Then       '4-2-03
'        'RptSelNT.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelNT, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelOF", 1) = 0 Then       '7-21-06
'        'RptSelOF.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelOF, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelos", 1) = 0 Then
'        'RptSelOS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelOS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelpa", 1) = 0 Then
'        'RptSelPA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelMA", 1) = 0 Then       '6-18-13 margin allocation
'        'RptSelMA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelMA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelParPay", 1) = 0 Then    '8-25-17 Participant Payables
'        'RptSelParPay.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelParPay, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelpc", 1) = 0 Then
'        'RptSelPC.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPC, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelpj", 1) = 0 Then
'        'RptSelPJ.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPJ, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelpp", 1) = 0 Then
'        'RptSelPP.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPP, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelpr", 1) = 0 Then       '6-15-04  Proposal Research Recap
'        'RptSelPr.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPr, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelps", 1) = 0 Then
'        'RptSelPS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelPS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelqb", 1) = 0 Then
'        'RptSelQB.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelQB, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelra", 1) = 0 Then
'        'RptSelRA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRD", 1) = 0 Then           '5-13-03
'        'RptSelRD.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRD, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRevEvt", 1) = 0 Then       '10-15-14 Revenue by Event
'        'RptSelRevEvt.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRevEvt, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRG", 1) = 0 Then           '12-22-09  Regional copy assignment
'        'RptSelRg.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRg, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRk", 1) = 0 Then          '7-30-12 Spot Price Ranking
'        'RptSelRk.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRk, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRI", 1) = 0 Then
'        'RptSelRI.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRI, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelRP", 1) = 0 Then           '11-1-02 Remote Posting
'        'RptSelRP.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRP, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelrs", 1) = 0 Then
'        'RptSelRS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelrr", 1) = 0 Then           '6-20-03 Research Revenue
'        'RptSelRR.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRR, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelrv", 1) = 0 Then
'        'RptSelRV.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelRV, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelca", 1) = 0 Then      '12/18/07 combo avails
'        'RptSelCA.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelCA, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelSN", 1) = 0 Then      '04-10-08 Split Network Avails
'        'RptSelSN.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelSN, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelsp", 1) = 0 Then
'        'RptSelSP.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelSP, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelspotBB", 1) = 0 Then       'Business on the books for air time, ntr and rep
'        'RptSelSpotBB.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelSpotBB, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelSR", 1) = 0 Then           '9-19-06 split regions
'        'RptSelSR.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelSR, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelss", 1) = 0 Then
'        ''RptSelss.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSeltx", 1) = 0 Then
'        'RptSelTx.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelTx, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelus", 1) = 0 Then
'        'RptSelUS.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelUS, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(sgReportFormExe, "RptSelAvgCmp", 1) = 0 Then
'        'RptSelAvgCmp.Show vbModal
'        ilRet = gCreateFormControlFile(RptSelAvgCmp, sgReportCtrlSaveName, slBypassCtrls())
'    ElseIf StrComp(UCase(sgReportFormExe), UCase("ExptReRate"), 1) = 0 Or StrComp(UCase(sgReportFormExe), UCase("ReRate"), 1) = 0 Then
'        'The control saving is within the ReRate and is not required here
'        'ExptReRate.Show vbModal
'        'ilRet = gCreateFormControlFile(ExptReRate, sgReportCtrlSaveName, slBypassCtrls())
'    Else
'        Exit Sub
'    End If

    ilRet = gCreateFormControlFile(fgReportForm, sgReportCtrlSaveName, slBypassCtrls())
    
End Sub

Public Sub gSetReportCtrlsSetting()
    Dim ilRet As Integer
    Dim slName As String
    Dim slStr As String
    
    On Error Resume Next
    
'    If StrComp(sgReportFormExe, "RptNoSel", 1) = 0 Then
'        'RptNoSel.Show vbModal
'        ilRet = gSetFormCtrls(RptNoSel, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSel", 1) = 0 Then
'        'RptSel.Show vbModal
'        ilRet = gSetFormCtrls(RptSel, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSel30", 1) = 0 Then           '6-12-13 cpp/cpm 30"unit
'        'RptSel30.Show vbModal
'        ilRet = gSetFormCtrls(RptSel30, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelaa", 1) = 0 Then
'        'RptSelAA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelac", 1) = 0 Then
'        'RptSelAc.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAc, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelAcqPay", 1) = 0 Then       '8-5-15
'        'RptSelAcqPay.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAcqPay, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelad", 1) = 0 Then
'        'RptSelAD.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAD, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelal", 1) = 0 Then       '4-14-04
'        'RptSelAL.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAL, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelalloc", 1) = 0 Then       '12-13-18   Revenue Allocation
'        'RptSelALLOC.Show vbModal
'        ilRet = gSetFormCtrls(RptSelALLOC, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelap", 1) = 0 Then
'        'RptSelAp.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAp, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelas", 1) = 0 Then
'        'RptSelAS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelav", 1) = 0 Then
'        'RptSelAv.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAv, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelBO", 1) = 0 Then       '7-22-11 Sales Breakout
'        'RptSelBO.Show vbModal
'        ilRet = gSetFormCtrls(RptSelBO, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelcb", 1) = 0 Then
'        'RptSelCb.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCb, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelcc", 1) = 0 Then       '1-15-04 Producer/Provider reports
'        'RptSelCC.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCC, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelcm", 1) = 0 Then       '10-5-10 Competitive Categories
'        'RptSelCM.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCM, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelcp", 1) = 0 Then
'        'RptSelCp.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCp, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelCt", 1) = 0 Then
'        'RptSelCt.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCt, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSeldb", 1) = 0 Then
'        'RptSelDB.Show vbModal
'        ilRet = gSetFormCtrls(RptSelDB, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSeldf", 1) = 0 Then
'        'RptSelDF.Show vbModal
'        ilRet = gSetFormCtrls(RptSelDF, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelds", 1) = 0 Then
'        'RptSelDS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelDS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelFd", 1) = 0 Then       '8-4-04 Feed Report
'        'rptSelFD.Show vbModal
'        ilRet = gSetFormCtrls(rptSelFD, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelia", 1) = 0 Then
'        'RptSelIA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelIA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelid", 1) = 0 Then       '5-21-02
'        'RptSelID.Show vbModal
'        ilRet = gSetFormCtrls(RptSelID, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelin", 1) = 0 Then
'        ''RptSelIn.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSelir", 1) = 0 Then       '7-13-05
'        'RptSelIR.Show vbModal
'        ilRet = gSetFormCtrls(RptSelIR, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSeliv", 1) = 0 Then
'        'RptSelIv.Show vbModal
'        ilRet = gSetFormCtrls(RptSelIv, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSellg", 1) = 0 Then
'        ''RptSellg.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSelNT", 1) = 0 Then       '4-2-03
'        'RptSelNT.Show vbModal
'        ilRet = gSetFormCtrls(RptSelNT, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelOF", 1) = 0 Then       '7-21-06
'        'RptSelOF.Show vbModal
'        ilRet = gSetFormCtrls(RptSelOF, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelos", 1) = 0 Then
'        'RptSelOS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelOS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelpa", 1) = 0 Then
'        'RptSelPA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelMA", 1) = 0 Then       '6-18-13 margin allocation
'        'RptSelMA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelMA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelParPay", 1) = 0 Then    '8-25-17 Participant Payables
'        'RptSelParPay.Show vbModal
'        ilRet = gSetFormCtrls(RptSelParPay, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelpc", 1) = 0 Then
'        'RptSelPC.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPC, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelpj", 1) = 0 Then
'        'RptSelPJ.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPJ, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelpp", 1) = 0 Then
'        'RptSelPP.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPP, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelpr", 1) = 0 Then       '6-15-04  Proposal Research Recap
'        'RptSelPr.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPr, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelps", 1) = 0 Then
'        'RptSelPS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelPS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelqb", 1) = 0 Then
'        'RptSelQB.Show vbModal
'        ilRet = gSetFormCtrls(RptSelQB, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelra", 1) = 0 Then
'        'RptSelRA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRD", 1) = 0 Then           '5-13-03
'        'RptSelRD.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRD, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRevEvt", 1) = 0 Then       '10-15-14 Revenue by Event
'        'RptSelRevEvt.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRevEvt, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRG", 1) = 0 Then           '12-22-09  Regional copy assignment
'        'RptSelRg.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRg, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRk", 1) = 0 Then          '7-30-12 Spot Price Ranking
'        'RptSelRk.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRk, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRI", 1) = 0 Then
'        'RptSelRI.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRI, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelRP", 1) = 0 Then           '11-1-02 Remote Posting
'        'RptSelRP.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRP, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelrs", 1) = 0 Then
'        'RptSelRS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelrr", 1) = 0 Then           '6-20-03 Research Revenue
'        'RptSelRR.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRR, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelrv", 1) = 0 Then
'        'RptSelRV.Show vbModal
'        ilRet = gSetFormCtrls(RptSelRV, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelca", 1) = 0 Then      '12/18/07 combo avails
'        'RptSelCA.Show vbModal
'        ilRet = gSetFormCtrls(RptSelCA, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelSN", 1) = 0 Then      '04-10-08 Split Network Avails
'        'RptSelSN.Show vbModal
'        ilRet = gSetFormCtrls(RptSelSN, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelsp", 1) = 0 Then
'        'RptSelSP.Show vbModal
'        ilRet = gSetFormCtrls(RptSelSP, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelspotBB", 1) = 0 Then       'Business on the books for air time, ntr and rep
'        'RptSelSpotBB.Show vbModal
'        ilRet = gSetFormCtrls(RptSelSpotBB, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelSR", 1) = 0 Then           '9-19-06 split regions
'        'RptSelSR.Show vbModal
'        ilRet = gSetFormCtrls(RptSelSR, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelss", 1) = 0 Then
'        ''RptSelss.Show vbModal
'    ElseIf StrComp(sgReportFormExe, "RptSeltx", 1) = 0 Then
'        'RptSelTx.Show vbModal
'        ilRet = gSetFormCtrls(RptSelTx, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelus", 1) = 0 Then
'        'RptSelUS.Show vbModal
'        ilRet = gSetFormCtrls(RptSelUS, sgReportCtrlSaveName)
'    ElseIf StrComp(sgReportFormExe, "RptSelAvgCmp", 1) = 0 Then
'        'RptSelAvgCmp.Show vbModal
'        ilRet = gSetFormCtrls(RptSelAvgCmp, sgReportCtrlSaveName)
'    ElseIf StrComp(UCase(sgReportFormExe), UCase("ExptReRate"), 1) = 0 Or StrComp(UCase(sgReportFormExe), UCase("ReRate"), 1) = 0 Then
'        'The control saving is within the ReRate and is not required here
'        'ExptReRate.Show vbModal
'        'ilRet = gSetFormCtrls(ExptReRate, sgReportCtrlSaveName)
'    Else
'        Exit Sub
'    End If
    ilRet = gSetFormCtrls(fgReportForm, sgReportCtrlSaveName)
End Sub


