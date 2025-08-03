Attribute VB_Name = "SLSPCOMMSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Slspcomm.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SlspComm.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for Sales Commsion
Option Explicit


Public tmSlspCommSalesperson() As SORTCODE
Public smSlspCommSalespersonTag As String


'********************************************************
'
'Sales Commission file definition
'
'*********************************************************
'Scf record layout
Type SCF
    lCode As Long        'Internal code number for Sales Commission
    iSlfCode As Integer 'Salesperson Code
    iVefCode As Integer 'Vehicle Code
    iStartDate(0 To 1) As Integer   'Start date
    iEndDate(0 To 1) As Integer 'End Date (TFN allowed)
    iUnderComm As Integer      'Commission paid if under sales Goal (xx.xx)
    iRemUnderComm As Integer      'Commission paid if under sales Goal for remnants(xx.xx)
    iDateEntrd(0 To 1) As Integer 'End Date (TFN allowed)
    iUrfCode As Integer 'User Code
    sUnused As String * 8
End Type
'Scf key record layout- use INTKEY0
'Type SCFKEY0
'    iCode As Integer
'End Type
'Type SCFKEY1
'    iSlfCode As Integer
'End Type
Type SCFREC
    sKey As String * 80 'Advt Product Contract # Transaction Date
    tScf As SCF
    iStatus As Integer
    iDateChg As Integer 'Date changed (True)
    lRecPos As Long
End Type
'Current
Public tgScfRec() As SCFREC  'Sales Commission
Public tgScfDel() As SCFREC
Public igModReturn As Integer    '0=Cancelled
Public tgScfAdd() As SCF
