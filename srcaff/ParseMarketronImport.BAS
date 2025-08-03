Attribute VB_Name = "ParseMarketronImportSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ParseMarketronImport.BAS
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
Public igShowHelpMsg As Integer
