Attribute VB_Name = "IMPTARSubs"
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
Public igShowHelpMsg As Integer

Type VEFMAP
    iFromVefCode As Integer
    iToVefCode As Integer
End Type

Type MISSEDVEHNAMES
    iVefCode As Integer
    sName As String * 40
    iStatus As Integer  '0=Not shown in list box; 1=Shown (don't show again)
End Type

Type MISSEDCNTR
    lCntrNo As Long
    lGross As Long
    lNet As Long
    iCount As Integer
End Type
