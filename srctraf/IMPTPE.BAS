Attribute VB_Name = "IMPTPESubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptRC.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
Public igBrowserType As Integer '0=Bmp (bit map); 1=Csv (comma delimited); 2= Txt (text)
Public igBrowserReturn As Integer   '0=Cancelled; 1=Ok
Public sgBrowserFile As String  'Drive\Path\FileName if igBrowserReturn=1 of the file
Public tmRateCard() As SORTCODE
Public smRateCardTag As String
Public igShowHelpMsg As Integer
Public igDPAltered As Integer
Type RIFREC
    sKey As String * 60 'Vehicle Name; Days Allowed (A=Yes; B=No); Time
    tRif As RIF
    iType As Integer    'Used for Comparison only(0=Base Rate Card Record; 1= Average Price to make Budget; 2= Difference)
    iStatus As Integer
    lRecPos As Long
    iLkYear As Integer  'Link to matching record but for different years
                        'This allows for multi-year rif records
End Type
Type RDFREC
    tRdf As RDF
    iStatus As Integer
    lRecPos As Long
End Type
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
Public sgDPName As String
Public igDPNameCallSource As Integer
Public igVefCode As Integer
Public sgWkStartDate As String
Public sgWkEndDate As String
Public igCurNoWks As Integer
Public igNewNoWks As Integer
Public igAvgPrices As Integer   'True or False
