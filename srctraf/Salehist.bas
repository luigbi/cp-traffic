Attribute VB_Name = "SALEHISTSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Salehist.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SaleHist.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the SaleHist subs and functions
Option Explicit
Option Compare Text
Type PHFREC
    sKey As String * 80 'Advt Product Contract # Transaction Date
    tPhf As RVF
    iStatus As Integer
    lRecPos As Long
    sProduct As String * 35
End Type

'Type SSPARTSH
'    iMnfSSCode As Integer   'Sales Source
'    iMnfGroup As Integer    'Participant
'    iVefIndex As Integer
'    iSSPartLp As Integer
'    iProdPct As Integer
'    sUpdateRvf As String * 1
'End Type
'Current
Public tgPhfRec() As PHFREC  'Sales History
Public tgPhfDel() As PHFREC
