Attribute VB_Name = "ExpNySubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPNY.BAS on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ExpMcFd subs and functions
Option Explicit
Option Compare Text
Type SPOTTIMESNY
    iVefCode As Integer
    sCodeStn As String * 2
    sBus As String * 5
    lAirDate As Long   'Air Date
    lAvailTime As Long  'Avail Time
    lNextSpotTime As Long   'Next spot time
End Type
