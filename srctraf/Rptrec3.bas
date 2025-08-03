Attribute VB_Name = "RPTREC3"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptrec3.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  IMRKEY0                                                                               *
'******************************************************************************************

' Proprietary Software, Do not copy
'
' File Name: RptRec3.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice support functions
Option Explicit
Option Compare Text
'IMR the ivr must match thought PayAddr
Type IVR
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long             'generation time
    'iGenTime(0 To 1) As Integer 'Generation time
    lSpotKeyNo As Long          'Spoy Key Number
    iType As Integer            'Record Type: (8/18/22 - for Combined and Seperate)
                                'AIR TIME and NTR Combined
                                '    0=Detail Air Time
                                '    1=Bonus spot
                                '    2=Vehicle Subtotal
                                '    3=Air Time total
                                '    4=CPM Detail  (new)
                                '    5=CPM Total   (new)
                                '    7=NTR Item Detail
                                '    8=NTR Total
                                '    9=Combination contract Total or Installment
                                '
                                'REP Invoice
                                '   2 = vehicle and market
                                '   3 = REP total

    'sTitle(1 To 4) As String * 15
    sTitle(0 To 3) As String * 15
    lChfCode As Long            'Contract Header Code
    iInvDate(0 To 1) As Integer 'Invoice Date
    lInvNo As Long              'Invoice Number
    sTerms As String * 30       'Terms
    iShowInvType As Integer     '0=None; 1=Prel; 2=Reprint
    sCashTrade As String * 1    'C=Cash; T= Trade
    iCTSplit As Integer         'Cash/Trade Split percent
    'sAddr(1 To 5) As String * 40 'Attention Account Payable
    sAddr(0 To 4) As String * 40 'Attention Account Payable
    'sPayAddr(1 To 4) As String * 40
    sPayAddr(0 To 3) As String * 40
    iLineNo As Integer          'Line Number
    sOVehName As String * 40    'Ordered Vehicle Name
    sODPName As String * 20     'Ordered Daypart Name
    iLen As Integer             'Spot Length
    sODays As String * 20       'Ordered Days
    iWkNo As Integer            'Week Number
    lONoSpots As Long           'Ordered Number of Spots
    sORate As String * 14       'Ordered Rate
    sADayDate As String * 20    'Aired Day, Date
    sATime As String * 12       'Aired Time
    sARate As String * 14       'Aired Rate
    'sACopy(1 To 4) As String * 45       'ISCI, Product
    sACopy(0 To 3) As String * 45       'ISCI, Product
    sAVehName As String * 40    'Aired Vehicle Name
    sRRemark As String * 40     'Reconc Remark
    sRAmount As String * 14     'Reconc Amount
    iPctComm As Integer         'Agy/Advt Commission
    lOTotalSpots As Long        'Total Ordered Number of Spots
    lOTotalGross As Long        'Total Ordered Gross Amount
    lATotalSpots As Long        'Total Aired Number of Spots
    lATotalGross As Long        'Total Aired Gross Amount
    lRTotalGross As Long        'Reconc Gross Amount
    'lComment(1 To 4) As Long    'Comments
    lComment(0 To 3) As Long    'Comments
    sEDIComment As String * 60
    sKey As String * 245        '5-17-06
    iMnfSort As Integer         'Mnf Sort Code
    iInvStartDate(0 To 1) As Integer 'Invoice Date
    iPrgEnfCode As Integer  'Program Event Name Code, only set for invoice form 3
    lDisclaimer As Long     'Comment code from Site Pref for Invoice Disclaimer
    lATotalNet As Long      '11-13-01 Inv total aired net
    lTax1 As Long
    lTax2 As Long
    iFormType As Integer    'Used with Form 2: 0=show with Order form; 1=Show with air form
    sSpotType As String * 1     '3-28-05 chg unused for Spot Type to determine billboard open/close spot
    lCode As Long
    lClfCxfCode As Long     'Line comment if to show on invoice otherwise zero
    sInstallCntr As String * 1  'Installment contract (Y/N or blank)
    lClfCode As Long        'Contract line reference code.  For NTR items this will be zero
    iInstallInvoiceNo     As Integer         ' Installment number of the invoice
                                             ' for the contract  (x of
                                             ' iivrTotalInstallInvs
    iTotalInstallInvs     As Integer         ' Total Installment number of
                                             ' invoices that will be generated
                                             ' for the contract
                                             ' (ivrInstallInvoiceNo of x)
    iRdfCode As Integer         'Daypart Reference code (Used to obtain daypart name or Avail Name)
    sPDFType              As String * 1      ' 0=No PDF; 1=PDF For Direct
                                             ' Advertiser; 2=PDF for Agency.
    iAdfOrAgfCode         As Integer         ' AdfCode (Direct advertiser) or
                                             ' AgfCode. This is used along with
                                             ' ivrPDFType
    sUnused As String * 2      'Unused - chged 11-8-01 from 22 to 18
End Type
Type IVRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'gen time
    'iGenTime(0 To 1) As Integer 'Generation time
    lInvNo As Long              'Invoice Number
End Type

'IMR the ivr must match thought PayAddr
Type IMR
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long             'generation time
    'iGenTime(0 To 1) As Integer 'Generation time
    lSpotKeyNo As Long          'Spoy Key Number
    iType As Integer            'Record Type: 0= Spot; 1= Bonus; 2=Subtotal; 3=Total
    'sTitle(1 To 4) As String * 15
    sTitle(0 To 3) As String * 15
    lChfCode As Long            'Contract Header Code
    iInvDate(0 To 1) As Integer 'Invoice Date
    lInvNo As Long              'Invoice Number
    sTerms As String * 30       'Terms
    iShowInvType As Integer     '0=None; 1=Prel; 2=Reprint
    sCashTrade As String * 1    'C=Cash; T= Trade
    iCTSplit As Integer         'Cash/Trade Split percent
    'sAddr(1 To 5) As String * 40 'Attention Account Payable
    sAddr(0 To 4) As String * 40 'Attention Account Payable
    'sPayAddr(1 To 4) As String * 40
    sPayAddr(0 To 3) As String * 40
    lIvrCode As Long
    sUnused As String * 20      'Unused - chged 11-8-01 from 22 to 18
End Type
Type IMRKEY0 'VBC NR
    iGenDate(0 To 1) As Integer 'Generation Date 'VBC NR
    '10-10-01
    lGenTime As Long            'gen time 'VBC NR
    'iGenTime(0 To 1) As Integer 'Generation time
End Type 'VBC NR

