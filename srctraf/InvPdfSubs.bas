Attribute VB_Name = "InvPDFSUBS"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: InvPDFSubs.BAS
'
' Release: 7.0
'
' Description:
'   This file contains the record definitions for Invoice PDF Email Addresses
Option Explicit


'Invoice email array to create separate pdfs by agency


Type INVPDF_INFO
        iAgfCode As Integer
        iAdfCode As Integer
        ilArfInvCode As Integer
        sPayeeName As String * 40
        sPDFExportPath As String * 120
        sPDFEmailAddress As String * 160        '11-18-16 estimate 4 addresses at 40 bytes each
        iSelectiveEmail As Integer              'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
        bHasAirTime As Boolean                  'Fix RE: v81 TTP 10826 - updated test results Thu 1/18/24 2:18 PM (Issue 4)
        bHasNTR As Boolean                      'Fix RE: v81 TTP 10826 - updated test results Thu 1/18/24 2:18 PM (Issue 4)
End Type

Type InvPDF_DETAILINFO                          '7-6-17 Array of pdf invoices created
        iAgfCode As Integer
        iAdfCode As Integer
        sAgencyName As String
        sAdvtName As String
        lCntrNo As Long
        lInvNo As Long
End Type

Type InvPDFEmailer_ErrMsg
    'sMsg As String * 75
    'sMsg As String * 120
    sMsg As String * 512
End Type

Public tgInvPDF_Info() As INVPDF_INFO  '
Public tgInvPDF_DetailInfo() As InvPDF_DETAILINFO       '7-6-17
Public tgInvPDFEmailer_ErrMsg() As InvPDFEmailer_ErrMsg

Public sgSetSelectionForFinals() As String      '12-27-16
Public sgSetSelectionForAll() As String         '12-27-16

