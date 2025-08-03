Attribute VB_Name = "RPTNOSELSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptnosel.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptNoSel.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report Generation code
Option Explicit
Option Compare Text
Public sgCurrentDollars As String
Public sgThisMonthsClosing As String
Dim tmRvf As RVF
Dim imRvfRecLen As Integer     'Record length
'*******************************************************
'*                                                     *
'*      Procedure Name:gReconcTotals                   *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute reconcile totals       *
'*                                                     *
'*******************************************************
Sub gReconcTotals(hlRvf As Integer)
    Dim slEndPrevPeriod As String
    Dim slEndCurrentPeriod As String
    Dim llEndPrevPeriod1 As Long
    Dim llEndCurrentPeriod As Long
    Dim slDate As String
    Dim ilRet As Integer
    Dim slNet As String
    Dim llTaxes As Long     '5-9-03
    Dim slTaxes As String   '5-9-03

    imRvfRecLen = Len(tmRvf)
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slEndPrevPeriod
    llEndPrevPeriod1 = gDateValue(slEndPrevPeriod) + 1
    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slEndCurrentPeriod
    llEndCurrentPeriod = gDateValue(slEndCurrentPeriod)
    sgCurrentDollars = ".00"
    sgThisMonthsClosing = ".00"
    ilRet = btrGetFirst(hlRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE)
        If tmRvf.sCashTrade = "C" Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slDate
            If gDateValue(slDate) <= llEndCurrentPeriod Then
                '5-9-03 incorporate taxes if applicable
                llTaxes = tmRvf.lTax1 + tmRvf.lTax2
                slTaxes = gLongToStrDec(llTaxes, 2)
                gPDNToStr tmRvf.sNet, 2, slNet

                sgThisMonthsClosing = gAddStr(gAddStr(sgThisMonthsClosing, slNet), slTaxes)
                If gDateValue(slDate) >= llEndPrevPeriod1 Then
                    sgCurrentDollars = gAddStr(gAddStr(sgCurrentDollars, slNet), slTaxes)
                End If
            End If
        End If
        ilRet = btrGetNext(hlRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub
