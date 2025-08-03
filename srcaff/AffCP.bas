Attribute VB_Name = "modCP"
'******************************************************
'*  modCP - various global declarations for Log and CP
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Public igCPOrLog As Integer '0=CP; 1=Log

Type CPINFO
    iStatus As Integer  '0=Show; 1=Bypass
    iVefCode As Integer
    sVefName As String * 42
    sDate As String * 10
    iCycle As Integer
    sRnfName As String * 8
    iRnfPlayCode As Integer
    sRnfOther As String * 8
    sVefState As String * 1
End Type

Type CMMLSUM
    iVefCode As Integer
    sZone As String * 3  '0=EST; 1=CST; 2=MST; 3=PST
    iAdfCode As Integer
    sProduct As String * 35     'changed from 20 to 35 2/27/98
    iLen As Integer
    iMFEarliest As Integer  'added 8/30/99
    iMFEarly As Integer     'added 2/1/98
    iMFAM As Integer
    iMFMid As Integer
    iMFPM As Integer
    iMFEve As Integer
    iSaEarliest As Integer  'added 8/30/99
    iSaEarly As Integer     'added 2/1/98
    iSaAM As Integer
    iSaMid As Integer
    iSaPM As Integer
    iSaEve As Integer
    iSuEarliest As Integer  'added 8/30/99
    iSuEarly As Integer     'added 2/1/98
    iSuAM As Integer
    iSuMid As Integer
    iSuPM As Integer
    iSuEve As Integer
    iTotal As Integer
    iDay(0 To 6)  As Integer
End Type

