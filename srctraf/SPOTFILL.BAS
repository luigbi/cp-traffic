Attribute VB_Name = "SPOTFILLSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Spotfill.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SpotFill.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the extra record definitions for contracts
Option Explicit
'Rate card dayparts
Type RCDP
    lStartTime As Long  'Start time of daypart
    lEndTime As Long    'End time of daypart
    'iDay(1 To 7) As Integer '0=False;1=True. Index 1=Monday; 2=Tuesday;....
    iDay(0 To 6) As Integer '0=False;1=True. Index 1=Monday; 2=Tuesday;....
End Type
Type CNTSPOT
    sKey As String * 90     'Sort by type, advertiser, Contract #, Vehicle
    sType As String * 1     'ChfType: C=Commercial Spot, S=PSA; M=Promo
    iAdfCode As Integer     'Advertiser Code
    lChfCode As Long        'Contract Code
    lFsfCode As Long
    iVefCode As Integer     'Spot Vehicle Code
    iLineNo As Integer      'Line number
    iLnVefCode As Integer   'Line vehicle (used to bypass allow day test)
    iNoSSpots As Integer     'Number of spots scheduled
    iNoGSpots As Integer     'Number of MG or Outside spots
    iNoMSpots As Integer     'Number of missed spots
    iNoESpots As Integer    'Number of Extra spots
    lSdfRecPos As Long        'Record image to be used when creating Fills
    sProduct As String * 35
    sLen As String * 3
    sDate As String * 17    'Date range of line
    iAllowedDays(0 To 6) As Integer
    sPriceType As String * 1
    lPrice As Long
    'lAllowedSTime(1 To 7) As Long
    'lAllowedETime(1 To 7) As Long
    lAllowedSTime(0 To 6) As Long
    lAllowedETime(0 To 6) As Long
    iNoTimesUsed As Integer
    iMnfComp0 As Integer
    iMnfComp1 As Integer
    iSpotLkIndex As Integer
    iRdfcode As Integer
    lAud As Long            'Audience
    iUpdateIndex As Integer
    iGameNo As Integer
    lSepLength As Long
End Type
Type SPOTLINKS
    iStatus As Integer
    lSdfRecPos As Long        'Record image to be used when creating Fills
    iSpotLkIndex As Integer       'Next Spot Link index
End Type
Type MGBOOKINFO
    iVefIndex As Integer
    iRdfIndex As Integer
End Type
Type AVAILINDEX
    lSsfRecPos As Long
    iAvEvt As Integer
    lEvtTime As Long
    iExcl As Integer    'True = violate exclusion
    iAvail As Integer   'True = violate avail name
    lFill1SdfCode As Long
    lFill2SdfCode As Long
End Type
Type AVAILMIX
    sKey As String * 5
    iRandomNo As Integer
    iColNo As Integer
    iRowNo As Integer
    iIndex As Integer
End Type
Public tgCntSpot() As CNTSPOT
Public tgTCntSpot() As CNTSPOT
Public tgDCntSpot() As CNTSPOT  'Detail
Public tgSCntSpot() As CNTSPOT  'Summary
Public tgSvCntSpot() As CNTSPOT
Public sgFillStartDate As String    'Start date for fill
Public sgFillEndDate As String    'End date for fill (used for games)
Public lgFillAllowDate As Long      'Allowed date for fill
Public igFillVefCode As Integer     'Fill vehicle code
Public igFillDW As Integer          '0=Daily view; 1=Weekly view
Public igSpotFillReturn As Integer    '0=No Change; >0 = Week changed
Public tgSpotLinks() As SPOTLINKS
'For SpotWks
Public igSpotWksReturn As Integer    '0=Cancelled; >0 = Done
Public sgSpotWksVehName As String
Public sgSpotWksStartDate As String
Type VEFRDFINFO
    sKey As String * 80    'Vehicle|Daypart
    iVefIndex As Integer
    iRdfIndex As Integer
    iDnfCode As Integer
    iAvail As Integer
    lAvgAud As Long
End Type
'For MGs +
Public sgMGStartDate As String    'Start date
Public lgMGEarliestDate As Long
Public lgMGAllowDate As Long      'Allowed date
Public lgMGLatestDate As Long
Public igMGVefCode As Integer     'vehicle code
Public igMFVpfIndex As Integer
Public lgChfMGCode As Long
Public lgFsfMGCode As Long
Public igMGSpotLen As Integer
Public igSpotMGReturn As Integer    '0=No Change; >0 = Week changed
Public tgVefRdfInfo() As VEFRDFINFO
'For Spots
Public tgSdfExt() As SDFEXT    'Spot scheduled for a contract
Public tgSdfExtSort() As SDFEXTSORT
Public igHourCount(0 To 23) As Integer 'Hour count
Public igDayCount(0 To 6) As Integer   'Day count
'Global tgAdMdManSchCode() As SORTCODE
'Global tgL2ManSchCode() As SORTCODE
Public tgMoVcf0() As VCF      'M-F Vcf records
Public tgMoVcf6() As VCF      'Sa Vcf records
Public tgMoVcf7() As VCF      'Su Vcf records
Public tgTuVcf0() As VCF      'M-F Vcf records
Public tgTuVcf6() As VCF      'Sa Vcf records
Public tgTuVcf7() As VCF      'Su Vcf records
Public tgWeVcf0() As VCF      'M-F Vcf records
Public tgWeVcf6() As VCF      'Sa Vcf records
Public tgWeVcf7() As VCF      'Su Vcf records
Public tgThVcf0() As VCF      'M-F Vcf records
Public tgThVcf6() As VCF      'Sa Vcf records
Public tgThVcf7() As VCF      'Su Vcf records
Public tgFrVcf0() As VCF      'M-F Vcf records
Public tgFrVcf6() As VCF      'Sa Vcf records
Public tgFrVcf7() As VCF      'Su Vcf records
Public tgSaVcf0() As VCF      'M-F Vcf records
Public tgSaVcf6() As VCF      'Sa Vcf records
Public tgSaVcf7() As VCF      'Su Vcf records
Public tgSuVcf0() As VCF      'M-F Vcf records
Public tgSuVcf6() As VCF      'Sa Vcf records
Public tgSuVcf7() As VCF      'Su Vcf records
Public tgAvailMix() As AVAILMIX

Public tgSpotAdvertiser() As SORTCODE
Public sgSpotAdvertiserTag As String

'  smShow will hold all of the data values from the files
Public tgMoEvt() As EVTINFO  'Event information for monday
Public tgTuEvt() As EVTINFO  'Event information for Tuesday
Public tgWeEvt() As EVTINFO  'Event information for Wednesday
Public tgThEvt() As EVTINFO  'Event information for Thursday
Public tgFrEvt() As EVTINFO  'Event information for Friday
Public tgSaEvt() As EVTINFO  'Event information for Saturday
Public tgSuEvt() As EVTINFO  'Event information for Sunday
Public tgMoRowIndex() As Integer 'Event information for monday index
Public tgTuRowIndex() As Integer  'Event information for Tuesday
Public tgWeRowIndex() As Integer  'Event information for Wednesday
Public tgThRowIndex() As Integer  'Event information for Thursday
Public tgFrRowIndex() As Integer  'Event information for Friday
Public tgSaRowIndex() As Integer  'Event information for Saturday
Public tgSuRowIndex() As Integer  'Event information for Sunday
Public igAMax As Integer
Public tgL1Evt() As EVTINFO  'Event information for monday
Public tgL1RowIndex() As Integer 'Event information for monday index
Public tgL1Vcf0() As VCF      'M-F Vcf records
Public tgL1Vcf6() As VCF      'Sa Vcf records
Public tgL1Vcf7() As VCF      'Su Vcf records
Public tgL2Evt() As EVTINFO  'Event information for monday
Public tgSdfMdExt() As SDFMDEXT
Public tgL2RowIndex() As Integer 'Event information for monday index
Public tgL2Vcf0() As VCF      'M-F Vcf records
Public tgL2Vcf6() As VCF      'Sa Vcf records
Public tgL2Vcf7() As VCF      'Su Vcf records
'Global tgManSchSdfExt() As SDFEXT    'Spot scheduled for a contract
'Global lgL2ManSchCode() As Long  'Contract code number
'Global igL2ManSchInfo() As Integer   'Index:1=Line #; 2=Len; 3=# Spots in week; 4=Max for week or day (-1=invalid week); 5=lbcL2ManSchCode index; 6=lbcL2ManSch index
'Global lgAdMdManSchCode() As Long  'Contract code number
'Global igAdMdManSchInfo() As Integer  'Index:1=Line #; 2=Len; 3=# Spots in week; 4=Max for week or day; 5=lbcAdMdManSchCode index; 6=lbcAdMdManSch index
                                   'For Extra Index:1=Unused; 2=Len; 3=# in Week; 4=Unused; 5=lbcAdMdManSchCode index; 6=lbcAdMdManSch index
Public tgAdMdSdfMdExt() As SDFMDEXT
'Public tgAdMdRowIndex(1 To 5) As Integer 'Event information for monday index
Public tgAdMdRowIndex(0 To 5) As Integer 'Event information for monday index
Public tgWAMdExt() As SDFMDEXT  'Event information for monday
'Global tgTrash() As SDFMDEXT
Public tgAddEvt As EVTINFO
Public tgSptMve() As SPOTMOVE

'Spot Action
Type SPOTACTIONINFO
    lChfCode As Long
    iSchdSpotCount As Integer
    iSchdFillCount As Integer
    iMissedSpotCount As Integer
    iCancelSpotCount As Integer
    iHiddenSpotCount As Integer
End Type

'For Spot Action
Public igSpotActionVefCode As Integer     'Fill vehicle code
Public sgSpotActionDefaultDate As String
Public igSpotActionReturn As Integer    '0=No Change; >0 = Week changed


'Find Contract
Type CNTRFINDSPOT
    sKey As String * 30
    iLineNo As Integer
    iGameNo As Integer
    sAirDate As String * 10
    sAirTime As String * 11
    sStatus As String * 1
    iAirVefCode As Integer  'MG/Outside
    sMissedDate As String * 10
    sMissedTime As String * 11
    iMissedGameNo As Integer
    sPtType As String * 1
    lCopyCode As Long
    sISCI As String * 20
End Type
