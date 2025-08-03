Attribute VB_Name = "ImptMarkSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Imptmark.bas on Wed 6/17/09 @ 12:56 P
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
Public tgDnfBook() As DNF
Public tgMnfCDemo() As MNF
Public tgMnfSDemo() As MNF
Type RDFINFO
    sKey As String * 60 'Sort Code
    iRdfIndex As Integer
End Type
Type DRFINFO
    tDrf As DRF
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
    iDemoGender As Integer
End Type
Type DPFINFO
    iCol As Integer
    tDpf As DPF
End Type
Public tgDpfInfo() As DPFINFO

Type DNFSORT
    sKey As String * 80 'Date; Book Name
    tDnf As DNF
End Type
