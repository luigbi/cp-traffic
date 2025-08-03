Attribute VB_Name = "RECEXTRA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Recextra.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RecExtra.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the File subs and functions
Option Explicit
Option Compare Text

'********************************************************
'
'One day log file definition
'  10/29/98 - removed 2 bytes from unused (19 to 17)
'             insert SortSeq for Crystal output (integer)
'  01/07/99 - changed unused from 17 to 13, inserted 4 bytes
'             for Avail comment ptr to be placed in spots
'  01/17/99 - inserted (before unused 13 bytes) the following:
'             odfrdfSortCode - sort code for DP sorting (2 bytes)
'             odfDPDesc - DP description (or override) 20 bytes
'             odfchfcxfCode - contr hdr "Other" comment code (4 bytes)
'             increased unused from 13 to 20 bytes
'  03/09/99 - Inserted (before unused 12 bytes) the following:
'             odfEvtIFCefCode - Event ID from Program Library (4 bytes)
'             odfDupeAvailID - Duplicated Avail ID to be appended to Event ID (5 bytes)
'             decreased unused to 3
'   5-30-01   Replace 2 byte string with integer for DP crystal sort field
'*********************************************************
'Odf record layout
Type ODF
    iUrfCode As Integer 'User who entered this record
    iVefCode As Integer 'Vehicle code number
    iAirDate(0 To 1) As Integer    'Start (Effective) date
    iAirTime(0 To 1) As Integer 'Air Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSeqNo As Integer   'Seq number to kept same time events in order
    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iFeedTime(0 To 1) As Integer 'Feed Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sZone As String * 3     'Time zone
    iEtfCode As Integer         'Event type code
    iEnfCode As Integer         'Event name code
    sProgCode As String * 5 'Program code #
    iMnfFeed As Integer     'Multi-name: Feed
    iDPSort As Integer      'replaced 5-30-01 sUnused1 As String * 2      'Unused
    iWkNo As Integer 'Week number of 1/5/70 (Monday- can't use 12/29/69 which is first monday of the year)
    ianfCode As Integer 'Avail name code
    iUnits As Integer   'Units
    iLen(0 To 1) As Integer     'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iAdfCode As Integer 'Advertiser code
    lCifCode As Long 'Copy
    sProduct As String * 35 'Product (either from contract or copy)
    iMnfSubFeed As Integer
    lCntrNo As Long 'Contract number
    'iType As Integer    '1=Prog; 2=Avail; 3=Spot; 9=Comment (other event type)
    'iPageNo As Integer  'Page number
    iBreakNo As Integer 'Reset at start of each program
    iPositionNo As Integer  'Reset at start of each avail
    iType As Integer    '1=Prog; 2=Comment; 3=Avail; 4=Spot
    lCefCode As Long    'Event comment
    sShortTitle As String * 15  'Short title
    iGameNo As Integer        'GameNumber (only set to game number if generating separate logs for each game)
    sPageEjectFlag As String * 1    'Page Eject(Y/N)
    iSortSeq As Integer         'seq # for Crystal sorting
    lAvailcefCode As Long       'comment ptr from avail placed in the spot (for log)
    iRdfSortCode As Integer     'DP sort code (1/17/99)
    sDPDesc As String * 20      'DP description (with override)
    lchfcxfCode As Long         'contr hdr Other comment code
    iDaySort As Integer         '1=M-F; 6=Sat; 7=Sunday
    lEvtCefCode As Long         'Other comment (14) (Stored into each Record)
    iEvtCefSort As Integer      'Incremented within a day each time a new Other Comment record found
    lEvtIDCefCode As Long       'Event ID from Program Library
    sDupeAvailID As String * 5  'Duplicated Avail ID from Delivery Links (this will be appended to EvtIDCefCode )
    sLogType As String * 1      'P=Preliminary Log; F or Blank = Final Log; R = Reprint
    '10-9-01
    iGenDate(0 To 1) As Integer    'Generation date
    'iGenTime(0 To 1) As Integer 'Generation Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    lGenTime As Long            'generation time
    lHd1CefCode As Long       'Header comment 1 from vehicle options
    lFt1CefCode As Long       'Footer comment 1 from vehicle options
    lFt2CefCode As Long       'Footer comment 2 from vehicle options
    'lVehNmCefCode As Long   'Vehicle Log Name
    lClfCode As Long           '1-17-14 name change
    sBonus As String * 1    '2-15-01 bonus flag (B)
    imnfSeg As Integer  '6-19-01
    '6/29/06
    sSplitNetwork         As String * 1      ' Split Network Flag:  N=Not a
                                             ' split metwork spot; P=Primary spot of
                                             ' a split netwrok; S=Secondary spot
                                             ' of a split network
    'sUnused As String * 19  '2-15-01 chged from 20 to 19
    'sUnused As String * 17   '6-19-01 chged from 19 to 17
    sBBDesc As String * 6       'BB description:  OBB, CBB, O/C BB, BB
    lRafCode As Long            '5-16-08 region code
    iAvailLen As Integer    'Original Avail Length (Set only for first spot in Break)
    sAvailLock As String * 1    'Avail Lock: A=Avail, S=Spot; B= Both; N or Blank = Neither
                                '7-25-13 dual defined field: C = Not locked but split copy, D = Avail+spot locked plus split copy,
                                'E = Avail locked + split copy, F =
    lCode As Long           'Autoincrement Code
    'sUnused As String * 3   '5/16/08 chg from 10 to6; 6/29/07 chg from 16 to 10
    iAlternateVefCode As Integer    '1-17-14 replaced unused: for log vehicle, this is the conventional vefcode
    sNewCopyInventory As String * 1 '1-17-14 replaced unusedNew copy inventory (y/n)
End Type
'Odf key record layout
Type ODFKEY0
    iVefCode As Integer 'Vehicle code number
    iAirDate(0 To 1) As Integer    'Start (Effective) date
    iLocalTime(0 To 1) As Integer 'Local Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sZone As String * 3     'Time zone
    iSeqNo As Integer   'Seq number to kept same time events in order
End Type
'Odf key record layout
Type ODFKEY1
    iMnfFeed As Integer     'Multi-name: Feed
    iAirDate(0 To 1) As Integer    'Start (Effective) date
    iFeedTime(0 To 1) As Integer 'Feed Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sZone As String * 3     'Time zone
    iSeqNo As Integer   'Seq number to kept same time events in order
End Type
Type ODFKEY2
    '10-9-01
    iGenDate(0 To 1) As Integer    'Generation date
    'iGenTime(0 To 1) As Integer 'Generation Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    lGenTime As Long                'generation time
End Type
'Odf key record layout- use LONGKEY0
'Type ODFKEY3
'    lCode As Long    'Internal code number (AutoInc)
'End Type

Type RSF
    lCode As Long           'Autoincrement Code
    lSdfCode As Long        'SDF Code number
    sPtType As String * 1   'Copy code type (0=none assigned; 1=cif code; 2= ccfcode)
    lCopyCode As Long        'Copy inventory code
    iRotNo As Integer       'Rotation #.
    lRafCode As Long        'Region Code
    lSBofCode As Long       'Suppress Blackout reference code if type = B
    lRBofCode As Long       'Replace Blackout reference code if type = B
    sType As String * 1     'B=Blackout; R (or blank)=Regional
    iBVefCode As Integer    'Vef Code for Type = B
    lRChfCode As Long       'Replace Blackout ChfCode (Type = B)
    lCrfCode As Long        'Crf Reference code
    iDateAdded(0 To 1) As Integer 'Date Added.  This is used with Alert to determine if export required(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    sUnused As String * 10
End Type

'Region info for L78 when generating logs by region
Type REGIONINFO
    lRafCode As Long
    sName As String
End Type

'Rsf key record layout- use LONGKEY0
'Type RSFKEY0
'    lCode As Long    'Internal code number (AutoInc)
'End Type
'Rsf key record layout- use LONGKEY0
'Type RSFKEY1
'    lSdfCode As Long    'SDF Code
'End Type
Type RSFKEY3
    lSdfCode              As Long
    lRafCode              As Long
End Type

Type RSFKEY4
    lSdfCode              As Long
    sType                 As String * 1
    iBVefCode             As Integer
End Type


'******************************************************************************
' emf_EMail_Content Record Definition
'
'******************************************************************************
Type EMF
    lCode                 As Long            ' E-Mail Content internal reference
                                             ' code
    sName                 As String * 20     ' E-Mail content name
    sContent              As String * 2000    ' Content of the E-Mail
    sType                 As String * 1      ' Content Type (I=Insertion Order;
                                             ' P=Proposal; C=Contract)
    iUrfCode              As Integer         ' User that added the content or
                                             ' the last user that changed it.
                                             ' User reference code
    iEnteredDate(0 To 1)  As Integer         ' Last date changed
    iEnteredTime(0 To 1)  As Integer         ' Last time changed
    sUnused               As String * 20     ' Unused
End Type


Type EMFKEY0
    lCode                 As Long
End Type

Type EMFKEY1
    sType                 As String * 1
End Type
Public sgEMailContentType As String * 1 'I=Insertion; P=Proposal; C=Contract
Public sgEMailContentName As String
Public igEMailContentCallSource As Integer

Type EMAILINFO
    sType As String * 1     'I=Insertion
    sRouteTo As String * 1  'S=Station; A=Agency; N=Network
    lChfCode As Long        'Used to obtain Contract #, Advertiser Name; Contract Product; Estimate Number
    lCntrNo As Long         'contract #
    iVefCode As Integer     'Used to obtain vehicle name
    sStartDate As String * 10 'contract start date
    sEndDate As String * 10     'contract end date
    lEmfCode As Long
    sAgyEstNo As String * 20        '8-20-15   chfagyestno + chftitle
    sAdvtName As String * 30
    sProduct As String * 35
    sResponseDate As String * 10
    sPDFFileName As String * 100
End Type

'TTP 10937 - Contract Dashboard: new screen to show contract header information
Type APPLIEDFILTER
    sType As String         'Filter Type (Example: Adf = Advertiser)
    lValue As Long          'Filter Value (Example: AdvertiserID)
    sValue As String          'Filter Value (Example: if Order status, "F" for Fully Scheduled)
End Type

'*******************************************************
'*                                                     *
'*      Procedure Name:gFieldOffsetExtra               *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the offset to a field *
'*                     within a record                 *
'*      5-25-01 Add ODFGenDate & ODFGenTime to offset table
'*                                                     *
'*******************************************************
Function gFieldOffsetExtra(slFile As String, slField As String) As Integer
'
'   sFile = "ADF"
'   sField = "slfFirstName"
'   ilOffset = gFieldOffsetExtra(sFile, sField)
'   Where:
'       sFile (I)- Name of the file
'       sField (I)- Name of the field as in the DDF
'       ilOffset (O)- The offset of the start of the field from 0 (-1 if not found)
'
    Dim slMsg As String
    Dim ilOffSet As Integer
    slMsg = ""
    ilOffSet = csiGetOffset(UCase$(slFile), UCase$(slField))
    If ilOffSet >= 0 Then
        gFieldOffsetExtra = ilOffSet
    Else
        slMsg = "Offset to field " & slField & " missing from file " & slFile
    End If

    'slMsg = ""
'    Select Case UCase$(slFile)
'        Case "ODF"
'            Select Case UCase$(slField)
'                Case "ODFURFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iUrfCode) '0
'                Case "ODFVEFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iVefCode) '0
'                Case "ODFAIRDATE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iAirDate(0)) '0
'                Case "ODFSEQNO"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iSeqNo) '0
'                Case "ODFLOCALTIME"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iLocalTime(0)) '0
'                Case "ODFZONE"
'                    gFieldOffsetExtra = 22 '0
'                Case "ODFETFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iEtfCode) '4
'                Case "ODFENFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iEnfCode) '4
'                Case "ODFPROGCODE"
'                    gFieldOffsetExtra = 29 '0
'                Case "ODFMNFFEED"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iMnfFeed) '4
'                Case "ODFANFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iAnfCode) '4
'                Case "ODFLENGTH"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iLen(0)) '4
'                Case "ODFADFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iAdfCode) '4
'                Case "ODFCIFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.lCifCode) '4
'                Case "ODFPRODUCT"
'                    gFieldOffsetExtra = 54 '0
'                Case "ODFMNFSUBFEED"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iMnfSubFeed) '4
'                Case "ODFBREAKNO"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iBreakNo) '4
'                Case "ODFPOSITIONNO"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.iPositionNo) '4
'                Case "ODFCEFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.lCefCode) '0
'                Case "ODFSHORTTITLE"
'                    gFieldOffsetExtra = 105 '0
'                Case "ODFEVTIDCEFCODE"
'                    gFieldOffsetExtra = GetOffSetForInt(tmOdf, tmOdf.lEvtIDCefCode) '0
'                Case "ODFDUPEAVAILID"
'                    gFieldOffsetExtra = 167 '0
'                Case "ODFGENDATE"
'                    gFieldOffsetExtra = 173    '5-25-01
'                Case "ODFGENTIME"
'                    gFieldOffsetExtra = 177    '5-25-01
'                Case "ODFMNFSEG"
'                    gFieldOffsetExtra = 198     '6-19-01
'                Case Else
'                    slMsg = "Offset to field " & slField & " missing from file " & slFile
'            End Select
'    End Select
    If slMsg = "" Then
        Exit Function
    End If
    MsgBox slMsg, vbOKOnly + vbCritical + vbApplicationModal, "Offset Error"
    gFieldOffsetExtra = -1
End Function
