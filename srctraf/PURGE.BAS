Attribute VB_Name = "PURGESUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Purge.bas on Wed 6/17/09 @ 12:56 PM *
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
Type SORTCIF
    sKey As String * 75 'Advertiser; Rotation End date; Cart # Or
                        'Rotation End date; Advertiser; Cart #
    sAdvtName As String * 42
    sInvName As String * 12
    sDate As String * 5
    sISCI As String * 20        '3-19-20   for sort by ISCI
    iChg As Integer     'True=Cif.sPurged changed
    tCif As CIF
End Type
