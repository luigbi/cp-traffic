Attribute VB_Name = "ArbitronImportSUBS"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ArbitronImport.Bas
'
' Release: 1.0
'
' Description:
Option Explicit
Option Compare Text
'Global igShowHelpMsg As Integer

'Public tgCAgf() As AGFADFCOLL
'Public sgCAgfStamp As String 'Date/Time stamp of file
'Public tgCAdf() As AGFADFCOLL   'Direct only
'Public sgCAdfStamp As String 'Date/Time stamp of file
'Public tgCAdfAll() As AGFADFCOLL
'Public sgCAdfAllStamp As String 'Date/Time stamp of file
'Public tgCSlf() As SLFEXT
'Public sgCSlfStamp As String 'Date/Time stamp of file
'Public tgCCdf() As CDFLIST
'Public sgCCdfStamp As String 'Date/Time stamp of file
'Public tgCommUrf() As URFEXT
'Public sgCommUrfStamp As String 'Date/Time stamp of file
'Public tgTransactionCode() As SORTCODE
'Public sgTransactionCodeTag As String

Public tgColAdvertiser() As SORTCODE
Public sgColAdvertiserTag As String

Public tgCopyAdvertiser() As SORTCODE
Public sgCopyAdvertiserTag As String

Public tgSpotAdvertiser() As SORTCODE
Public sgSpotAdvertiserTag As String

Public tmCntAdvertiser() As SORTCODE
Public smCntAdvertiserTag As String

Public tmCntAgency() As SORTCODE
Public smCntAgencyTag As String

Public tmCntRateCard() As SORTCODE
Public smCntRateCardTag As String

Public tmVehicleCode() As SORTCODE
Public smVehicleCodeTag As String

Public tmCntSalesperson() As SORTCODE
Public smCntSalespersonTag As String

Public tmCntBuyerCode() As SORTCODE
Public smCntBuyerCodeTag As String

'Type SAF
'    iCode As Integer
'    sSchRNG As String * 1   'Scheduling: Use T=Clock As Seed; F=Fixed Value As Seed
'    sSchMaL As String * 1   'Major Loop on: H=Hour; D=Day; Q=Quarter Hour
'    sSchMdL As String * 1   'Middle Loop on: H=Hour; D=Day; Q=Quarter Hour
'    sSchMiL As String * 1   'Minor Loop on: H=Hour; D=Day; Q=Quarter Hour
'    sSchCycle As String * 1 'Schedule cycle pattern: 1=Iterate on equal counts (major; middle; minor); 2=Iterate on minor in count order
'    sSchMove As String * 1  'Schedule Move Pass: N=N/A; A=each iteration; B=Each Count; C=Minor Loop; D=Middle Loop; E=Major Loop
'    sSchCompact As String * 1  'Schedule Compact Pass: N=N/A; A=each iteration; B=Each Count; C=Minor Loop; D=Middle Loop; E=Major Loop
'    sSchPreempt As String * 1  'Schedule Preempt Pass: N=N/A; A=each iteration; B=Each Count; C=Minor Loop; D=Middle Loop; E=Major Loop
'    sSchHour As String * 1   'For Hour Loop: N=N/A; R=Randomize Counts; S=Sort Counts Only; B=Sort Counts and randomize = counts
'    sSchDay As String * 1   'For Day Loop: N=N/A; R=Randomize Counts; S=Sort Counts Only; B=Sort Counts and randomize = counts
'    sSchQH As String * 1    'For Quarter Hour Loop: N=N/A; R=Randomize Counts; S=Sort Counts Only; B=Sort Counts and randomize = counts
'    iImptSeqNo As Integer   'Contract Import Sequence Number
'    sName As String * 20
'    iCurSeqNo As Integer    'Current Sequence number so that move can be reimported
'    lLastMoveID As Long     'ID of last move ID successfully completed
'    iVefcode As Integer
'    lLowPrice As Long
'    lLevelToPrice(1 To 11) As Long
'    lHighPrice As Long
'    iPreferredPct As Integer    'Prefereed percentage of spots to try to book into the preferred days/times (min # spots will be 1)
'    iWk1stSoloIndex As Integer  'Last Week, 1st Position and Solo Avail (.xx)
'    sUnused As String * 10  'was 8
'End Type
'Public tgSaf() As SAF

Declare Function btrCreFile Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal RecLength&, ByVal PageSize%, ByVal NumIndexes%, ByVal uFileFlags%, ByVal uAllocation%) As Integer
Declare Function btrCreKey Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal KeyPosition&, ByVal KeyLength&, ByVal KeyFlags%, ByVal ExtendKeyType%, ByVal NullValue%) As Integer
Declare Function btrCreCreate Lib "cbtrv432.dll" (ByVal Ohnd%, ByVal fileName$, ByVal fOverWrite%) As Integer

Type Arbitron
    lCode As Long
    sCallLetters As String * 4
    sBand As String * 2
    sDataType As String * 1     '1=All Data; 2=Ethnic Black Only; 3= Ethnic Hispanic Only
    sAreaCovered As String * 1  'M=Metro; D=DMA
    sDemoType As String * 1     'P=Population; A=Audience
    iStartHour As Integer       'Start Hour (12m-1a=0; 1a-2a=1;...11p-12a=23)
    iEndHour As Integer         'End Hour
    iDayIndex As Integer        '1=Mon (used for Mon-Fri); 6=Saturday; 7=Sunday
    sForm As String * 1         'Number of demo data fields (6=16; 8=18)
    lDemo(1 To 18) As Long
End Type

Type ARBKEY1
    sCallLetters As String * 4
    sBand As String * 2
    sDataType As String * 1     '1=All Data; 2=Ethnic Black Only; 3= Ethnic Hispanic Only
    iDayIndex As Integer        '1=Mon (used for Mon-Fri); 6=Saturday; 7=Sunday
    iStartHour As Integer       'Start Hour (12m-1a=0; 1a-2a=1;...11p-12a=23)
End Type

Type ARBAUDCHAR
    lID As Long
    sName As String * 10
    iBucketID As Integer
End Type

Type ARBCALLLETTERS
    lID As Long
    sName As String * 4
    sBand As String * 2
End Type

Type ARBDAYPART
    lID As Long
    iDayIndex As Integer
    iStartHour As Integer
    iEndHour As Integer
End Type


'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterModalForm                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center modal form within        *
'*                     Traffic Form                    *
'*                                                     *
'*******************************************************
Sub gCenterModalForm(FrmName As Form)
'
'   gCenterModalForm FrmName
'   Where:
'       FrmName (I)- Name of modal form to be centered within Traffic form
'
    Dim flLeft As Single
    Dim flTop As Single
    flLeft = ArbitronImport.Left + (ArbitronImport.Width - FrmName.Width) / 2
    flTop = ArbitronImport.Top + (ArbitronImport.Height - FrmName.Height - 60) / 2
    FrmName.Move flLeft, flTop
End Sub
