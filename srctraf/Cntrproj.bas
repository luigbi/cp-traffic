Attribute VB_Name = "CNTRPROJSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Cntrproj.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CntrProj.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
Type PJF1REC
    sKey As String * 62 'Advertiser; Product; Year
    tPjf As PJF
    sAdvtName As String * 42
    sProdName As String * 20
    iStatus As Integer
    lRecPos As Long
    iSaveIndex As Integer   'Index into smSave
    i2RecIndex As Integer   'Second record index
End Type
Type PJF2REC
    tPjf As PJF
    iStatus As Integer
    lRecPos As Long
End Type
'Type USERVEH
'    iCode As Integer
'    sName As String * 40
'End Type
Type PJSALEOFFICE
    iCode As Integer
    sName As String * 40
End Type
Type PJPDGROUPS
    iYear As Integer
    iStartWkNo As Integer
    iNoWks As Integer
    iTrueNoWks As Integer
    iFltNo As Integer
    sStartDate As String
    sEndDate As String
End Type
'Current
Public tgPjf1Rec() As PJF1REC  'Projection
Public tgPjf2Rec() As PJF2REC  'Projection
Public tgPjfDel() As PJF2REC
'Original Current
Public tgOPjf1Rec() As PJF1REC  'Projection
Public tgOPjf2Rec() As PJF2REC  'Projection
'Prior Week
Public tgPPjf1Rec() As PJF1REC  'Projection
Public tgPPjf2Rec() As PJF2REC  'Projection
'Public sgWkStartDate As String
'Public sgWkEndDate As String
'Public igCurNoWks As Integer
'Public igNewNoWks As Integer
