Attribute VB_Name = "SMFCHECKSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SMFCHECK.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SmfCheck.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions to be converted
Option Explicit
Option Compare Text
'Declare Sub HMemCpy Lib "kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%, ByVal lpFileName$)
'Declare Function GetOffSetForInt Lib "CPS.DLL" (i As Any, j As Any) As Integer
'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
'Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
'Declare Function GetVersion Lib "Kernel" () As Long
'In ImptMGSubs
'Public sgVefStamp As String
'Public tgVef() As VEF
Type SDFINFO
    lCode As Long
    lChfCode As Long
    iLineNo As Integer
    iDate(0 To 1) As Integer
    iTime(0 To 1) As Integer
    sSchStatus As String * 1
    lSmfCode As Long
    iMatchFound As Integer
End Type
Type DELSMF
    lRecPos As Long
    iReason As Integer
End Type
Type SMFMAKE
    sInfo As String * 100
    lSdfCode As Long
    lRecPos As Long 'If created
End Type
