Attribute VB_Name = "ImptRadSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Imptrad.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptRad.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions to be converted
Option Explicit
Public tgVefRad() As VEF
'Public tgDnfBook() As DNF
Public tgMnfSocEcoRad() As MNF
'Global tgRdf() As RDF
Type VEHMERGE
    tDrf As DRF
    iCount As Integer   'Number of records merged
End Type
Public tgVehMerge() As VEHMERGE
