Attribute VB_Name = "csiArchive3"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: csiArchive3.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Global Variables for Set Advertiser/Agency Credit
Option Explicit
Option Compare Text

    
'Report Queue
    'Report Queue
    Public hgDBRqf As Integer
    Public tgDBRqf As RQF
    Public tgDBRqfSrchKey0 As LONGKEY0
    Public igDBRqfRecLen As Integer

    'Report Parameters
    Public hgDBRff As Integer
    Public tgDBRff As RFF
    Public tgDBRffSrchKey0 As LONGKEY0
    Public tgDBRffSrchKey1 As RFFKEY1
    Public igDBRffRecLen As Integer
    
Type CRFARCHIVEINFO
    sKey As String * 10
    lCrfCode As Long
    lDate As Long
End Type

Type DBVEHICLE
    sType As String * 1
    iVefCode As Integer
End Type

'Public tgStations() As SHTTINFO
'Public sgStationsStamp As String


Type STATUSTYPES
    sName As String * 30
    iPledged As Integer '0=Live; 1=Delayed; 2=Not Carried; 3=No Pledge
    iStatus As Integer
End Type

'Public tgStatusTypes(0 To 12) As STATUSTYPES
Public tgStatusTypes(0 To 14) As STATUSTYPES

Type ADFCREDIT
    iCode As Integer
    sUnbilled As String * 12
    sProjected As String * 12
    s1Wk As String * 12
    s4Wks As String * 12
    s13Wks As String * 12
End Type
Type AGFCREDIT
    iCode As Integer
    sUnbilled As String * 12
    sProjected As String * 12
    s1Wk As String * 12
    s4Wks As String * 12
    s13Wks As String * 12
End Type
Type SDFCREDITEXT
    lCode As Long           'Spot Reference Number
    iVefCode As Integer     'Vehicle Code (combos not allowed)
    lChfCode As Long        'Contract code (used to get Agency; Advertiser)
    iLineNo As Integer      'Line number
    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
    sPriceType As String * 1 'L=Line; P=Post Log N/C  old=>T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
    sBill As String * 1     'Billed status (Y= Yes; N=No)
End Type

Type SCR
    lCode As Long                   'auto code
    iGenDate(0 To 1) As Integer     'Report generation date
    lGenTime As Long                'report generation time
    iStrLen As Integer  'String length (required by LVar)
    sScript As String * 5002   'Last two bytes after the comment must be 0

End Type

Type SCRKEY1
    iGenDate(0 To 1) As Integer   'Generation Date
    lGenTime As Long
End Type

'******************************************************************************
' ept Record Definition
'
'******************************************************************************
Type EPT
    lCode                 As Long            ' Estimate Pledge table reference
                                             ' code
    lDatCode              As Long            ' Pledge Daypart/Avail table
                                             ' (Dat.mkd) reference code
    iSeqNo                As Integer         ' Sequence number by DAT
    lAttCode              As Long            ' Agreement reference code
    iShttCode             As Integer         ' Station Reference code
    iVefCode              As Integer         ' Vehicle reference code
    sFdAvailDay           As String * 2      ' Feed Avail Day
    iFdAvailTime(0 To 1)  As Integer         ' Feed Avail Time
    sEstimatedDay         As String * 2      ' Estimated Day
    iEstimatedTime(0 To 1) As Integer        ' Estimated Time
    sUnused               As String * 20     ' Unused
End Type


Type EPTKEY2
    lAttCode              As Long
End Type


Public hgDBEpt As Integer
Public tgDBEpt As EPT
Public tgDBEptSrchKey2 As EPTKEY2
Public igDBEptRecLen As Integer

'******************************************************************************
' pet Record Definition
'
'******************************************************************************
Type PET
    lCode                 As Long            ' Pledge by Event table reference
                                             ' auto code
    lAttCode              As Long            ' Agreement reference code
    iVefCode              As Integer         ' Vehicle reference code
    iShttCode             As Integer         ' Station reference code
    lGsfCode              As Long            ' Event schedule reference code
    sDeclaredStatus       As String * 1      ' Declared status. Y=will Carry;
                                             ' N=will Not Carry; U (or
                                             ' blank)=Undedcided. Test for Y or
                                             ' N
    sClearStatus          As String * 1      ' Clearance status. Y=Carried;
                                             ' N=did Not Carry; U (or
                                             ' Blank)=Unknown. Test for Y or N
    iUstCode              As Integer         ' User reference code that last
                                             ' added or changed setting.
    iEnteredDate(0 To 1)  As Integer         ' Date last added or changed
    iEnteredTime(0 To 1)  As Integer         ' Time last added or changed
    sUnused               As String * 20     ' Unused
End Type


Type PETKEY1
    lAttCode              As Long
End Type

Public hgDBPet As Integer
Public tgDBPet As PET
Public tgDBPetSrchKey1 As PETKEY1
Public igDBPetRecLen As Integer

'******************************************************************************
' VAT_Vendor_Agreement Record Definition
'
'******************************************************************************
Type VAT
    lAttCode              As Long            ' Agreement internal reference code
    iWvtVendorID          As Integer         ' Vendor ID
    sSentToWeb            As String * 1      ' Sent to Web (Y/N). Blank or N =
                                             ' No; Y = Yes; M=Modified
    sUnused               As String * 10     ' Unused
End Type


Type VATKEY0
    lAttCode              As Long
End Type

Public hgDBVat As Integer
Public tgDBVat As VAT
Public tgDBVatSrchKey0 As VATKEY0
Public igDBVatRecLen As Integer

