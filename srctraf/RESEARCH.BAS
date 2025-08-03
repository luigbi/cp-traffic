Attribute VB_Name = "RESEARCHSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Research.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Research.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Research subs and functions
Option Explicit
Option Compare Text
Public igDnfModel As Integer    '>0 is DnfCode value
Public igReturn As Integer  '0=Cancel; 1=Done
Public igDuplDnfCode As Integer
Public sgPercentChg As String
Public igResearchModelMethod As Integer '0=Book; 1=Vehicle
Public bgResearchByImpressions As Boolean
Type RESEARCHADJUSTVEHICLE
    sVehicleName As String * 40
    iVefCode As Integer
    bAdjust As Boolean
End Type
Public tgResearchAdjustVehicle() As RESEARCHADJUSTVEHICLE

Type DRFREC
    sKey As String * 60     'Vehicle
    tDrf As DRF
    iStatus As Integer  '0=New; 1=Old; 2=Delete Old
    lRecPos As Long
    lIndex As Long   'Index into tgAllDrf
    iModel As Integer   'True=Created in model (retain values in SaveRec)
    '8/14/18: retain the drfcode so that matching up with dpf
    lModelDrfCode As Long
    lLink As Long
    iCustInfoIndex As Integer
End Type
Type SAVESHOW
    'sSave(1 To 23) As String
    'sShow(1 To 23) As String
    'Index zero ignored in arrays below
    sSave(0 To 27) As String
    sShow(0 To 27) As String
End Type
Type DPFREC
    sKey As String * 10 'Demo
    iStatus As Integer
    lDpfCode As Long
    lDrfCode As Long
    sDemo As String * 10                   '# of people in thousands
    sPop As String * 10                    'Population for plus demos, used when getting population for non-standard demos (base 16)
    lIndex As Long
    sSource As String * 1                   'B=Base to Non-Base; M = Model; D=Dynamic
    iRdfCode As Integer                     'Used with Source = B to match up
End Type

Type DEFREC
    sKey As String * 10 'StartDate
    iStatus As Integer
    lDefCode As Long
    sStartDate As String * 10              'StartDate
    sPop As String * 10                    'Population for Estimate, used when getting audience numbers
    sEstPct As String * 10
End Type

Type DRFMAP
    lInitDrfCode As Long
    lNewDRfCode As Long
    iRdfCode As Integer     'Used with Base to Non-base duplication
End Type

Type RSCHBYCUSTINFO
    sKey As String * 10
    iDnfCode As Integer
    lPopDrf As Long
    lCustDrf As Long
    lCustDemo As Long
    lPopDemo As Long
    sDate As String * 10
End Type

Type CUSTINFO
    sName As String * 20
    sDataType As String * 1
    iDemoIndex As Integer
End Type

Public tgMnfSocEco() As MNF
Public tgDnf As DNF
Public tgDrfRec() As DRFREC        'DRf record image
Public tgSDrfPop As DRF            'Population record for standard demos
Public tgCDrfPop() As DRF            'Population record for custom demos
Public tgGDrfPop() As DRF
Public tgDrfDel() As DRFREC
Public tgAllDrf() As DRFREC
Public tgDpfRec() As DPFREC
Public tgDpfDel() As DPFREC
Public tgAllDpf() As DPFREC
Public tgLinkDrfRec() As DRFREC

Public tgDefRec() As DEFREC
Public tgDefDel() As DEFREC

