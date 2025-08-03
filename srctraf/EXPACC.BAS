Attribute VB_Name = "EXPACCSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expacc.bas on Wed 6/17/09 @ 12:56 PM 
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
Type SALESOFFICE
    sOffice As String * 20
    sRegion As String * 20
    iCode As Integer
End Type
