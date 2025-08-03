Attribute VB_Name = "ExpBkCpySubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expbkcpy.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpBkCpy.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ExpBkCpy subs and functions
Option Explicit
Option Compare Text
Type SORTCRF
    sKey As String * 140 'Advertiser; Contract #; Vehicle; Rotation # (descending)
    lCntrNo As Long
    sCntrProd As String * 35    'Chf.sProduct
    sType As String * 1 'Chf.sType
    lCrfRecPos As Long
    iSelected As Integer
    iCombineIndex As Integer   '-1=None; >=0 first Crf to combine
    iDuplIndex As Integer       '-1=None; >=0 next Crf to combine
    iVpfIndex As Integer    'Index into tgVpfInfo
    tCrf As CRF
End Type
Type COMBINECRF
    lCntrNo As Long
    sVehName As String * 40
    lCrfRecPos As Long
    iCombineIndex As Integer   '-1=None; >=0 next Crf to combine
    iVpfIndex As Integer    'To get vehicle code
    tCrf As CRF
End Type
Type DUPLCRF
    lCntrNo As Long
    sVehName As String * 40
    lCrfRecPos As Long
    iDuplIndex As Integer       '-1=None; >=0 next Crf to combine
    iVpfIndex As Integer    'To get vehicle code
    tCrf As CRF
End Type
Type CYFTEST
    lCifCode As Long
    iVefCode As Integer
    sSource As String * 1
    sTimeZone As String * 3
    lRafCode As Long
End Type
Type SENDROTINFO
    sKey As String * 250 'Vehicle Name; Contract Product; Rotation Start Date; Rotation Start Time; Zone
    lCrfCode As Long
    iVefCode As Integer
    iStatus As Integer  '1=Send; 2=Don't send
    iRevised As Integer 'True=Revised instructions
    iSortCrfIndex As Integer
End Type
Type SENDCOPYINFO
    sKey As String * 300 'Vehicle name; Product; Cif.lCode; Start Date; End Date
    sXFKey As String * 80  'Product; Cart; ISCI; Creative Title
    tCyf As CYF
    sChfProduct As String * 35
    lRotStartDate As Long
    lRotEndDate As Long
    lPrevFdDate As Long
    iFdDateNew As Integer
    iAdfCode As Integer
    iLen As Integer
End Type
Type VPFINFO
    tVpf As VPF
    iNoVefLinks As Integer      'Number of vehicles to be combined (secondary vehicles)
    'iVefLink(0 To 9) As Integer 'Vef of vehicle to combine (secondary)
    'sVefName(0 To 9) As String  'Name of vehicle to combine (secondary)
    iFirstLkVehInfo As Integer
    iFirstSALink As Integer
End Type
Type LKVEHINFO
    iVefCode As Integer
    sVefName As String * 40
    iNextLkVehInfo As Integer
End Type
'Used only in cmcView_click
Type VEHTIMES
    iVefCode As Integer
    sVefName As String * 220
    lTotalSpotTime As Long
End Type
Public igBFCall As Integer      '1=Resend; 2=Suppress
Public igBFReturn As Integer    '0=Cancelled; 1=Update rotation
Public tgSortCrf() As SORTCRF
Public tgCombineCrf() As COMBINECRF
Public tgDuplCrf() As DUPLCRF
Public tgAddCyf() As SENDCOPYINFO
