Attribute VB_Name = "FEEDSPOTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of FeedSpot.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  FSFKEY1                       FSFKEY2                                                 *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FeedSpot.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for Sales Commsion
Option Explicit


Public igPdCodeFpf As Integer 'Mode from
Public igPdReturn As Integer   '0=Cancel; 1=Model


'********************************************************
'
'Feed Names file definition
'
'*********************************************************
'Fnf record layout
Type FNF
    iCode As Integer            'Internal code number
    sName As String * 40        'Name like Paul Harvey, Information,..
    sFTP As String * 70         'FTP or URL address to get Spots from
    sPW As String * 10          'Password
    sChkDays As String * 7      'Days to check for Web Spots
    lChkInterval As Long        'Time interval to check for spots (in minutes).  Like every 2 hours (enter 120), every 15 min (enter 15), ..
    iChkStartHr(0 To 1) As Integer  'Start hour for check
    iChkEndHr(0 To 1) As Integer    'End hour for checking
    iProdArfCode As Integer     'Producer ARF pointer
    iNetArfCode As Integer      'Network/Rep ARF Pointer
    ianfCode As Integer         'Avail Name to book into
    iMcfCode As Integer         'Media Code to use for Auto-Cart assigment
    sPledgeTime As String * 1   'Log Pledge Time:  I=Insertion Order; L = Log needs Conversion (FPF);  P=Pre-Coverted Log
    sState As String * 1    'A=Active; D=Dormant
    sUnused As String * 20
End Type
'Fnf key record layout
'Type FNFKEY0
'    iCode As Integer
'End Type

'********************************************************
'
'Feed Pledge file definition
'
'*********************************************************
'Fpf record layout
Type FPF
    iCode As Integer            'Internal Code
    iFnfCode As Integer         'Feed Name Pointer
    iVefCode As Integer         'Vehicle Name Pointer
    iEffStartDate(0 To 1) As Integer    'Effective Start Date of this Pledge
    iEffEndDate(0 To 1) As Integer      'Effective end date of this pledge
    iUrfCode As Integer         'User
    sUnused As String * 20
End Type

'Fpf key record layout
'Type FPFKEY0
'    iCode As Integer
'End Type

Type FPFKEY1
    iFnfCode As Integer
    iVefCode As Integer
    iEffEndDate(0 To 1) As Integer
End Type

Type FPFKEY2
    iFnfCode As Integer
    iVefCode As Integer
    iEffStartDate(0 To 1) As Integer    'Effective Start Date of this Pledge
End Type


'********************************************************
'
'Feed Pledge Data file definition
'
'*********************************************************
'Fpf record layout
Type FDF
    iCode As Integer            'Internal Code
    iFpfCode As Integer         'Feed Name Pointer
    iFeedStartTime(0 To 1) As Integer    'Feed Time of Spot
    iFeedEndTime(0 To 1) As Integer    'Feed Time of Spot
    sFeedDays(0 To 6) As String * 1     'Feed Day of spot(for M-f=1111100)
    iPledgeStartTime(0 To 1) As Integer 'Pledge Start Time
    iPledgeEndTime(0 To 1) As Integer   'Pledge End Time
    sPledgeDays(0 To 6) As String * 1   'Pledge Days to air (Sa-Su=0000011)
    sUnused As String * 20
End Type
'Fdf key record layout
'Type FDFKEY0
'    iCode As Integer
'End Type

'Fdf key record layout
Type FDFKEY1
    iFpfCode As Integer
End Type


'********************************************************
'
'Feed Spot file definition
'
'*********************************************************
'Fsf record layout
Type FSF
    lCode As Long           'Internal Code
    iFnfCode As Integer     'Feed Name Pointer
    iVefCode As Integer     'Vehicle Name Pointer
    sRefID As String * 10   'Reference # like contract # or ASTCode if importing
    iRevNo As Integer       'Revision #, to be used to get history of spots
    iAdfCode As Integer     'Advertiser Pointer
    lPrfCode As Long     'Product Pointer
    iMnfComp1 As Integer    'Product protection Pointer
    iMnfComp2 As Integer    'Product Protection Pointer
    iLen As Integer         'Spot Length
    iStartDate(0 To 1) As Integer   'Start Date to air
    iEndDate(0 To 1) As Integer     'End Date to Air
    iRunEvery As Integer            'Run Every (0=Every week, 1=every other week,..)
    sDyWk As String * 1     'Daily or Weekly (D=Daily, W=Weekly)
    iNoSpots As Integer     'Number of Spots if Weekly
    iDays(0 To 6) As Integer    'If Daily then number of Spots each day, If weekly then 0(no) or 1(Yes) to indicate if airing day
    iStartTime(0 To 1) As Integer   'Airing start time or Feed time
    iEndTime(0 To 1) As Integer     'Airing end time or feed time
    lCifCode As Long            'Copy inventory
    sSchStatus As String * 1    'Schedul status (N=New or Changed, F=Fully scheduled, P=Partially scheduled, D=Deleted)
    iEnterDate(0 To 1) As Integer   'Entered Date
    iEnterTime(0 To 1) As Integer   'Entered Time
    iUrfCode As Integer         'User
    lPrevFsfCode As Long        'Previous Feed Spot pointer.  Used to obtain history of changes
    lAstCode As Long            'Affiliate Spot pointer from Web System
    sUnused As String * 20
End Type
'Fsf key record layout
'Type FSFKEY0
'    lCode As Long
'End Type
Type FSFKEY1 'VBC NR
    sRefID As String * 10 'VBC NR
    iRevNo As Integer 'VBC NR
End Type 'VBC NR
Type FSFKEY2 'VBC NR
    sSchStatus As String * 1 'VBC NR
End Type 'VBC NR
Type FSFKEY3
    iFnfCode As Integer     'Feed Name Pointer
    iVefCode As Integer     'Vehicle Name Pointer
    iStartDate(0 To 1) As Integer   'Start Date to air
    iEndDate(0 To 1) As Integer     'End Date to Air
End Type
Type FSFKEY4
    lPrevFsfCode As Long        'Previous Feed Spot pointer.  Used to obtain history of changes
End Type
Type FSFKEY5
    lAstCode As Long            'Affiliate Spot pointer from Web System
End Type

Type FSFREC
    sKey As String * 80 'Advt Product Contract # Transaction Date
    tFsf As FSF
    iStatus As Integer
    iDateChg As Integer 'Date changed (True)
    lRecPos As Long
End Type
'Current
Public tgFsfRec() As FSFREC  'Sales Commission
Public tgFsfDel() As FSFREC

Type FDFREC
    sKey As String * 80 'Advt Product Contract # Transaction Date
    tFdf As FDF
    iStatus As Integer
    iDateChg As Integer 'Date changed (True)
    lRecPos As Long
End Type

'Current
Public tgPlgeRec() As FDFREC  'Sales Commission
Public tgPlgeDel() As FDFREC

'Pass values to Pledge
Public igPledgeFnfCode As Integer
Public igPledgeVefCode As Integer

