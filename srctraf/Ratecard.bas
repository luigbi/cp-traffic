Attribute VB_Name = "RATECARDSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Ratecard.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RateCard.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
Public Const RCAVGRATEINDEX = 12 'This is the AvgRate column that is Shared between RateCard, StdPkg, and CPMPkg screens!
Public tmRateCard() As SORTCODE
Public smRateCardTag As String
Public igStdPkgModel As Integer
Public igStdPkgReturn As Integer    '0=Cancel; 1=Done
Type RIFREC
    sKey As String * 60 '0 or 1 (0=Spots defined); Vehicle Name; Days Allowed (A=Yes; B=No); Time
    tRif As RIF
    iType As Integer    'Used for Comparison only(0=Base Rate Card Record; 1= Average Price to make Budget; 2= Difference)
    iStatus As Integer
    lRecPos As Long
    lLkYear As Long  'Link to matching record but for different years
                        'This allows for multi-year rif records
End Type
Type RCPBDPGEN
    sKey As String * 60
    sSvKey As String * 60    'Same as sKey except missing high order part of sort
    iRdfCode As Integer
    sVehName As String * 50
    sDPName As String * 100
    iVefCode As Integer
    lAvgPrice As Long
    lSvAvgPrice As Long
    iAvgRating As Integer
    lAvgAud As Long
    lGrImp As Long
    lGRP As Long
    lCPP As Long
    lCPM As Long
    lPop As Long
    iVehDormant As Integer
    iDPDormant As Integer
    iPkgVeh As Integer
    sMedium As String
End Type
Public tmRifRec() As RIFREC
Public tmLkRifRec() As RIFREC  'Multiyear records
Public smRCSave() As String * 40  'Values saved (1=Vehicle; 2=Daypart; 3=Acquisition, 4=Base; 5=Report; 6=Sort)    '3=$ Index; 4=% Inv)
Public lmRCSave() As Long  'Values saved (1-4=Dollars, 5=Avg Dollar; 6=Total Dollar)
Public imRCSave() As Integer   'Values saved (1-4=Number of weeks with dollars, 5-8=Actual number of weeks,
                               '9=Dormant Vehicle(1=Yes;0=No); 10=Dormant Daypart(1=Yes; 0=No); 11=Package vehicle(1=Yes; 0=No);
                               '12-15=Week with zero rate(1=Yes; 0=No))
Public smRCShow() As String * 40  'Values shown in rate card area
Public smDPShow() As String * 40  'Values shown in rate card area (one extra for base daypart flag)
Public smBdShow() As String * 40
'Public igShowHelpMsg As Integer
'Public igDPAltered As Integer
Type USERVEH
    iCode As Integer
    sName As String * 40
End Type
Type PDGROUPS
    iYear As Integer
    iStartWkNo As Integer
    iNoWks As Integer
    iTrueNoWks As Integer
    iFltNo As Integer
    sStartDate As String
    sEndDate As String
End Type
Public tgRcfI As RCF
Public igRCMode As Integer
Public igRcfModel As Integer    'RcfCode to model from (zero=none)
Public igRcfChg As Integer
Public igRCReturn As Integer    '0=Cancelled
Public sgRCModelDate As String
Type RCMODELINFO
    iRdfCode As Integer
    iVefCode As Integer
End Type
'Public sgDPName As String
'Public igDPNameCallSource As Integer
'Public igVefCode As Integer
Public sgWkStartDate As String
Public sgWkEndDate As String
Public igCurNoWks As Integer
Public igNewNoWks As Integer
Public igAvgPrices As Integer   'True or False
'Reallocation of Dollars
Type CFFAUD
    iCffIndex As Integer
    iVefIndex As Integer
    lAud As Long
End Type
Public smPkgShow() As String * 30
Public smPkgSave() As String * 10
Public sgStdPkgName As String

Public igRCNoDollarColumns As Integer 'Number of dollar columns displayed

Public igDPNameCallSource As Integer
Public sgDPName As String
Public igVefCode As Integer
Public igDPAltered As Integer

