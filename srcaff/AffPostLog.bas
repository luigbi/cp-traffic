Attribute VB_Name = "modPostLog"
'******************************************************
'*  modPostLog - various global declarations for Pre-Log and Post Log
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Type POSTINFO
    iType As Integer '0=Spot; 1=Avail; 2=New Spot (either converted avail or inserted)
    lSdfCode As Long
    sDateZone(0 To 3) As String * 10
    sTimeZone(0 To 3) As String * 11
    sProd As String * 35
    lCntrNo As Long
    lLstCodeZone(0 To 3) As Long
    lCifZone(0 To 3) As Long 'Cart or ISCI
    sCartZone(0 To 3) As String * 30 'Cart or ISCI
    iWkNoZone(0 To 3) As Integer
    iBreakNoZone(0 To 3) As Long
    iPositionNoZone(0 To 3) As Integer
    iSeqNoZone(0 To 3) As Integer
    iLen As Integer
    iUnits As Integer
    iStatus As Integer
    sAdfName As String * 35
    iAdfCode As Integer
    iAgfCode As Integer
    iAnfCode As Integer
    iChgd As Integer
End Type

Type CNTRINFO
    lCntrNo As Long
    sProd As String * 35
    lChfCode As Long
    iAdfCode As Integer
    iAgfCode As Integer
End Type

Type COPYINFO
    lCifCode As Long
    lCpfCode As Long
    sCart As String * 7
    sISCI As String * 20
End Type

Public igPreOrPost As Integer
