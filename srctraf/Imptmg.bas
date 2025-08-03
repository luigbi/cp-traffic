Attribute VB_Name = "ImptMGSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Imptmg.bas on Wed 6/17/09 @ 12:56 PM 
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  tgMoveRec                                                                             *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  ERRMSG                                                                                *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ConvFile.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions to be converted
Option Explicit
'Declare Sub HMemCpy Lib "kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%, ByVal lpFileName$)
'Declare Function GetOffSetForInt Lib "CPS.DLL" (i As Any, j As Any) As Integer
'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
'Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
'Declare Function GetVersion Lib "Kernel" () As Long
Public tgVef() As VEF
'********************************************************
'
'MG Track file definition
'
'*********************************************************
'Mtf record layout
Type MTF
    lCode As Long   'AutoIncrement
    lChfCode As Long    'Contract ChfCode
    iLineNo As Integer  'Line Number
    lTrackID As Long    'Tracking ID
    lRefTrackID As Long 'Reference Tracking ID (if 'From' spot was a MG, then its TrackID)
    lTransGpID As Long  'Transaction Group ID
    iEnteredDate(0 To 1) As Integer 'Entered Date
    iSdfVefCode As Integer  'SDF vehicle code
    iSdfDate(0 To 1) As Integer 'SDF Date
    iSdfTime(0 To 1) As Integer 'SDF Time
    sSdfSchStatus As String * 1 'SDF Schedule Status
    lFromPrice As Long      'Original Spot Price (xxxxxxx.xx)
    lToPrice As Long      'To Spot Price (xxxxxxx.xx)
    lPrevTransGpID As Long  'Transaction Group ID
    iToStartTime(0 To 1) As Integer 'SDF Time
    iToEndTime(0 To 1) As Integer 'SDF Time
    iToRdfCode As Integer
    iDays(0 To 6) As Integer    '1=Air Day; 0=Not ari daye; Index 0=Monday; 1=Tuesday,...
    sUnused As String * 4
End Type
'Mtf key record layout- use LONGKEY0
'Type MTFKEY0
'    lCode As Long   'AutoIncrement
'End Type
'Type MTFKEY1- use LONGKEY0
'    lTrackID As Long    'Tracking ID
'End Type
'Type MTFKEY2- use LONGKEY0
'    lRefTrackID As Long 'Reference Tracking ID (if 'From' spot was a MG, then its TrackID)
'End Type
Type ERRMSG 'VBC NR
    lCntrNo As Long 'VBC NR
    iLineNo As Integer 'VBC NR
    iRdfCode As Integer 'VBC NR
End Type 'VBC NR
Type MGMOVEREC
    iStatus As Integer  '0=Not Processed; 1= Processed
    lCreatedDate As Long    'Created date
    lRefTrackID As Long
    lTrackID As Long
    lTransGpID As Long
    lCntrNo As Long         'Contract #
    iLineNo As Integer      'Line #
    iFromVefCode As Integer
    lFromDate As Long
    lFromSTime As Long
    lFromETime As Long
    iToVefCode As Integer
    lToDate As Long
    lToSTime As Long
    lToETime As Long
    sOper As String * 1
    lFromPrice As Long
    lToPrice As Long
    sOpResult As String * 1
    lCreatedTime As Long
    sFromVehicle As String * 40
    sToVehicle As String * 40
    iFromAnfCode As Integer
    iToAnfCode As Integer
    iSpotLen As Integer
    iDays(0 To 6) As Integer
    lMoveID As Long
    lRecNo As Long
End Type
