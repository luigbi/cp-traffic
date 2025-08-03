Attribute VB_Name = "VEHOPTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Vehopt.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VehOpt.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for VOF
Option Explicit

'********************************************************
'
'Vehicle Log Option file definition
'
'*********************************************************
'Vof record layout
Type VOF
    iVefCode As Integer     'Vehicle Code
    sType As String * 1     'L=Log; C=C.P.; O=Other
    lHd1CefCode As Long     'Header comment
    lFt1CefCode As Long     'Foot 1 Note
    lFt2CefCode As Long     'Foot 2 Note
    iNoDaysCP As Integer    'No of Days to return C.P.
    sShowLen As String * 1  'Show Length
    sShowProduct As String * 1  'Show Product
    sShowCreative As String * 1 'Show creative title
    sShowISCI As String * 1 'Show ISCI
    sShowDP As String * 1   'Show Daypart Description
    sShowAirTime As String * 1  'Show Air Time
    sShowAirLine As String * 1  'Show Air Line
    sShowHour As String * 1 'Show Hour
    sSkipPage As String * 1 'Skip Page by day/week
    iLoadFactor As Integer
    sUnused As String * 18
End Type
'Vof key record layout
Type VOFKEY0
    iVefCode As Integer     'Vehicle Code
    sType As String * 1     'L=Log; C=C.P.; O=Other
End Type


'******************************************************************************
' VBF_Vehicle_Barter Record Definition
'
'******************************************************************************
Type VBF
    lCode                 As Long            ' AutoIncrement number
    iVefCode              As Integer         ' Vehicle reference code
    iStartDate(0 To 1)    As Integer         ' Start Date (Monday date only)
    iEndDate(0 To 1)      As Integer         ' End Date
    'iSpotLen(1 To 10)     As Integer         'Spot Length
    'lDefAcqCost(1 To 10)  As Long            ' Default Acquisition Cost function of spot length
                                             ' (xxxxxx.xx)
    'lActAcqCost(1 To 10)  As Long            ' Actual Acquisition Cost function of spot length
                                             ' (xxxxxx.xx)
    iSpotLen(0 To 9)     As Integer         'Spot Length
    lDefAcqCost(0 To 9)  As Long            ' Default Acquisition Cost function of spot length
                                             ' (xxxxxx.xx)
    lActAcqCost(0 To 9)  As Long            ' Actual Acquisition Cost function of spot length
                                             ' (xxxxxx.xx)
    sMethod               As String * 1      ' Barter Method (A=All Cash; M=Pay
                                             ' when Minutes exceeded; U=Pay when
                                             ' Units Exceeeded; X=X Free for Y
                                             ' Sold)
    iThreshold            As Integer         ' Minutes or Unit threshold
    iXFree                As Integer         ' Number of X free spots/week for Y
                                             ' sold spots/week
    iYSold                As Integer         ' Number of Y sold spots/week to
                                             ' get X free spots/week
    lInsertionCefCode     As Long            ' Insertion comment
    iAcqCommPct           As Integer         ' Acquisition commission (xx.xx)
    lBalance              As Long         ' Balance
    iBalanceDate(0 To 1)  As Integer         ' Date balance last computed
    sPerPeriod            As String * 1      ' Per period: W=Weekly; M=Monthly;
                                             ' Y=Yearly
    sUnused               As String * 10     ' Unused
End Type


'Type VBFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type VBFKEY1
    iVefCode              As Integer
    iStartDate(0 To 1)    As Integer         ' Start Date (Monday date only)
End Type


'******************************************************************************
' PIF_Participant_Info Record Definition
'
'******************************************************************************
Type PIF
    lCode                 As Long            ' Auto increment field set by
                                             ' Pervasive
    iVefCode              As Integer         ' Vehicle Reference Code
    iStartDate(0 To 1)    As Integer         ' Start Date of participant
                                             ' Percentages (must be start of
                                             ' standard Broadcast month)
    iEndDate(0 To 1)      As Integer         ' End Date (TFN = 12/31/2069)
    iSeqNo                As Integer         ' Sequence number (used to retain
                                             ' the order)
    iMnfSSCode            As Integer         ' Sales Source Reference code.
                                             ' MnfType = S
    iMnfGroup             As Integer         ' Participant (Owner):  Mnf
                                             ' Reference Code. MnfType = H and
                                             ' UnitType = 1 (H1).
    sUpdateRVF            As String * 1      ' Participant-Update Rvf(Y) or
                                             ' Phf(N) flag
    sExtUpdateRvf         As String * 1      ' Participant-External Update
                                             ' Rvf(Y) or Phf(N) flag
    iProdPct              As Integer         ' Producers Percentage (xx.xx)
    sUnused               As String * 10     ' Unused
End Type


'Pdf key record layout- use LONGKEY0
'Type PIFKEY0
'    lCode                 As Long
'End Type

Type PIFKEY1
    iVefCode              As Integer
    iStartDate(0 To 1)    As Integer
    iSeqNo                As Integer
End Type

Type PIFDATES
    sStartDate As String * 10
    sEndDate As String * 10
End Type

Type PIFREC
    tPif As PIF
    iStatus As Integer  '1=Old, 2=New Added to old; 3=All New
End Type
