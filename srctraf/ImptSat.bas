Attribute VB_Name = "ImptSatSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ImptSat.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptMark.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions to be converted
Option Explicit

Type SATDPINFO
    iType As Integer    '0=Vehicle; 1=Daypart
    iBkNm As Integer    '0=Daypart Book name; 1=Exact Time Book Name
    iStartCol As Integer
    iSY As Integer
    iEY As Integer
    sDays(0 To 6) As String * 1
    lStartTime As Long
    lEndTime As Long
    lStartTime2 As Long
    lEndTime2 As Long
End Type

'Not used
'Type SATDEMO
'    iDPIndex As Integer
'    iVefCode As String
'    lDrfCode As Long
'    lDemo(1 To 18) As Long
'    lPlus(1 To 18) As Long
'End Type

Type SATEXTRAPOP
    lPop As Long
    iMnfDemo As Integer
End Type
