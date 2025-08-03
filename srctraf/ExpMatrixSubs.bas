Attribute VB_Name = "ExpMatrixSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExpMatrixSubs.bas on Wed 6/17/09 @ 12
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'       Matrix Export
' Release: 5.1
'
' Description:
'   This file contains the Report Get Data for Crystal screen code
Option Explicit
Option Compare Text

'6-4-14 fields added to use with Efficio export
Type MATRIXINFO
    iVefCode As Integer            'vehicle code
    iSlfCode As Integer             'primary slsp
    iAgfCode As Integer            'agency code
    lAgfCRMId As Long               'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
    iAdfCode As Integer            'advertiser code
    lAdfCRMId As Long               'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
    sProduct As String * 35        'product string
    sOrderType As String * 1        'C=Standard; V=Reservation; T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
    sCashTrade As String * 1        'C = cash , T = trade
    sAirNTR As String * 1           'A = air time, N = NTR
    '6-6-14 change all arrays to 2 yrs to 3 years
    'lDirect(1 To 37) As Long         '1st slsp has the direct $ (gross) of contract
    'iYear(1 To 37) As Integer                'year
    'iMonth(1 To 37) As Integer               'month #
    'lGross(1 To 37) As Long         'gross $
    'lNet(1 To 37) As Long           'net $
    'lAcquisition(1 To 37) As Long   'acquisition$ 1-28-14
    'Index zero ignored in arrays below
    lDirect(0 To 37) As Long         '1st slsp has the direct $ (gross) of contract
    iYear(0 To 37) As Integer                'year
    iMonth(0 To 37) As Integer               'month #
    lGross(0 To 37) As Long         'gross $
    lNet(0 To 37) As Long           'net $
    lAcquisition(0 To 37) As Long   'acquisition$ 1-28-14
    iMnfComp1 As Integer            'prod protection 1
    iMnfComp2 As Integer            'prod protection 2
    'efficio added fields (not used by matrix export)
    lCntrNo As Long                 'order #
    lExtCntrNo As Long               'TTP 10503 - RAB export: include agency CRM ID, advertiser CRM ID, Boostr Campaign Number
    sSalesRegion As String * 1        'from sales source:  L = local, R = regional, N = National
    lInvoice As Long                'invoice #
    iNTRType As Integer              'ntrtype internal code (RAB export only)
    lOHDPacingDate As Long          '6-17-20 for RAB
    sTransactionType As String      '11-03-20 for RAB, TTP 10004
    lReceivablesDateEntered As Long '1-4-21 for CustomRevExport, TTP 9992
    sComment(0 To 37) As String     'For TTP 10666
    iLineNo As Integer              'For TTP 10743
    dCurrMoAvg(0 To 37) As Double            'TTP 10742 - RAB Cal Spots manual export: when "include digital avg comments" is checked on, show current month and next month averages in separate comment columns to assist troubleshooting
    dNextMoAvg(0 To 37) As Double            'For TTP 10742
End Type

