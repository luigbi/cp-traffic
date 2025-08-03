Attribute VB_Name = "COPYINVSUB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyinv.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyInv.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the File subs and functions
Option Explicit
Option Compare Text
Public sgISCI As String
Public sgCreativeTitle As String
Public igSortCart As Integer
Public tgCif As CIF

Type FDCRTNOINFO
    lEndDate As Long
    sListInfo As String * 120   'Advertiser|Advertiser AddrID|Status|Date Range|Product
End Type
