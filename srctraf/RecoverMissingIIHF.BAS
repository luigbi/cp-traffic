Attribute VB_Name = "GENASFSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of GENASF.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  AVAILINFO                     ASFKEY2                                                 *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCnCSA.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Export Selling to Airing Links subs and functions
Option Explicit
Option Compare Text


Type AVAILINFO 'VBC NR
    iRdfIndex As Integer 'VBC NR
    iAvail60(0 To 26) As Integer 'VBC NR    'index zero ignored
    iAvail30(0 To 26) As Integer 'VBC NR
    iAvail10(0 To 26) As Integer 'VBC NR
End Type 'VBC NR

'Global igShowHelpMsg As Integer
Type ASF
    iStartWkDate(0 To 1) As Integer 'Start Week date
    iUnused1 As Integer
    iRdfCode As Integer
    iUnused2 As Integer
    iVefCode As Integer
    sAvail60(0 To 25) As String * 1 '0-199 available, 200=200+available; 201-249=1-49 Oversold, 250=50+oversold; 255=not set, index zero ignored
    sAvail30(0 To 25) As String * 1 '0-199 available, 200=200+available; 201-249=1-49 Oversold, 250=50+oversold; 255=not set
    sAvail10(0 To 25) As String * 1 '0-199 available, 200=200+available; 201-249=1-49 Oversold, 250=50+oversold; 255=not set
    lCode As Long                   '1-24-01 Auto increment
    iRcfCode As Integer             '1-24-01 Rate Card code
    sNRProcessed          As String * 1      ' N-R Processed (Y or N)
    'sUnused               As String * 23     ' 1-24-01 chged from 30 to 24
    iCreatedDate(0 To 1) As Integer 'Date asf record was created
    iCreatedTime(0 To 1) As Integer 'Time asf record was created
    sUnused               As String * 15     ' 1-24-01 chged from 30 to 24
End Type

'Type ASFKEY0
'    iVefCode              As Integer
'End Type

'Type ASFKEY1
'    lCode                 As Long
'End Type

Type ASFKEY2 'VBC NR
    sNRProcessed          As String * 1 'VBC NR
End Type 'VBC NR


'Type RUINTKEY
'    iRemoteID As Integer
'    iAutoCode As Integer
'End Type
