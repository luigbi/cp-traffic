Attribute VB_Name = "ExpAW"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPNY.BAS on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpAW.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ExpAW subs and functions
Option Explicit
Option Compare Text

'******************************************************************************
' AWF_AirWave Record Definition
' 
'******************************************************************************
Type AWF
    lCode                 As Long            ' Auto Increment assigned by Pervasive
    iDateExported(0 To 1) As Integer         ' Date Exported
    iSeqNo                As Integer         ' Sequence number starting at 1
    iUrfCode              As Integer         ' User Reference code
    iEnteredDate(0 To 1)  As Integer         ' Entered date
    iEnteredTime(0 To 1)  As Integer         ' Entered Time
    sUnused               As String * 10     ' Unused
End Type


'Type AWFKEY0
'    lCode                 As Long
'End Type

Type AWFKEY1
    iDateExported(0 To 1) As Integer
End Type
