Attribute VB_Name = "EXPCMCHGSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expcmchg.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCmChg.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Type TYPESORTCC
    sKey As String * 100
    sStnCode As String * 5
    iSelVefIndex As Integer
    lRecPos As Long
End Type
Type STNCODECC
    iVefcode As Integer
    sStnCode As String * 5
End Type
Type SELVEFCC
    iVefcode As Integer     'Selected vehicle
    iWithData As Integer    'Data shown for vehicle
End Type
