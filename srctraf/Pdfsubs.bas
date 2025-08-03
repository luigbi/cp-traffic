Attribute VB_Name = "PDFSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Pdfsubs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PDFSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Form subs and functions
Option Explicit
Option Compare Text
'*******************************************************
'*                                                     *
'*      Procedure Name:gSwitchToPrinter                *
'*                                                     *
'*             Created:5/6/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Switch Print Driver             *
'*                                                     *
'*******************************************************
Sub gSwitchToPDF(cdcSetup As control, ilType As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    '
    '  ilType(I)- 0=Switch to PDF, 1=Switch to Default Printer
    '
    cdcSetup.flags = cdlPDPrintSetup
'    SendKeys "%" & Trim$(tgUrf(0).sPrtNameAltKey)  'alt N or P to go to name field
'    If ilType = 0 Then
'        'SendKeys tgUrf(0).sPDFDrvChar
'        For ilLoop = 1 To tgUrf(0).iPDFDnArrowCnt Step 1
'            'SendKeys "{Down}"
'            SendKeys tgUrf(0).sPDFDrvChar
'        Next ilLoop
'    Else
'        'SendKeys tgUrf(0).sPrtDrvChar
'        For ilLoop = 1 To tgUrf(0).iPrtDnArrowCnt Step 1
'            'SendKeys "{Down}"
'            SendKeys tgUrf(0).sPrtDrvChar
'        Next ilLoop
'    End If
'    For ilLoop = 1 To tgUrf(0).iPrtNoEnterKeys Step 1
'        SendKeys "{Enter}" 'Enter Key
'    Next ilLoop
    cdcSetup.Action = 5    'DLG_PRINT
End Sub
