Attribute VB_Name = "ExpProjSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of expproj.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'       Projection Export
' Release: 5.3
'
Option Explicit
Option Compare Text
'Export Projection: projection data for a day, hour, & spot length
Type PROJINFO
    lDate As Long
    iLen As Integer
    iInventory(1 To 24) As Integer  'inventory defined for this spot length
    lMinRate(1 To 24) As Long       'lowest spot rate (zero included)
    lMaxRate(1 To 24) As Long       'highest spot rate
    lSchedRev(1 To 24) As Long      'total $ scheduled (missed included)
    lMissedRev(1 To 24) As Long     'total $ missed
    iSchedUnits(1 To 24) As Integer 'total scheduled  units (missed included)
    iNCUnits(1 To 24) As Integer    'total no charge (zero)  units
    iMissedUnits(1 To 24) As Integer    'total missed units
End Type

