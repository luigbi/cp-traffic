Attribute VB_Name = "RecDefXFSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RecDefXF.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions to be converted
Option Explicit
'********************************************************
'
'Log History file definition
'
'*********************************************************
'Lhf record layout
Type LHF
    iVefCode As Integer       'Vehicle code
    iDate(0 To 1) As Integer  'Date of error
    iTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSeqNo As Integer           'Sequency number (used to order events at same time)
    iEtfCode As Integer         'Event type code
    ianfCode As Integer         'Avail name code to book into (Event type = 2,3,4,5,6,7,8)
    iUnits As Integer        'Max units (zero if only length used for avails)
    iLen As Integer             'length or zero if only by units
End Type
'Lhf key record layout
Type LHFKEY0
    iVefCode As Integer       'Vehicle code
    iDate(0 To 1) As Integer  'Date of error
    iTime(0 To 1) As Integer 'Start Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iSeqNo As Integer           'Sequency number (used to order events at same time)
End Type
''********************************************************
''
''Spot history file definition
''
''*********************************************************
''Shf record layout
'Type SHF
'    lCode As Long       'AutoInc
'    iVefCode As Integer       'Vehicle code
'    lChfCode As Long        'Contract code
'    iLineNo As Integer      'Line number
'    iAdfCode As Integer     'Advertiser code number
'    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
'                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
'    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    iSeqNo As Integer       'Schedule spot seq #
'    lSpotID As Long     'lSdfCode
'    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready?, U=Unscheduled,
'                                'G=Makegood, A=on alternate log but not MG, B=on alternate Log and MG,
'                                'C=Cancelled; H=Hidden
'    sType As String * 1
'    sPriceType As String * 1
''    iMnfMissed As Integer   'Missed reason
''    sBB As String * 1     'BB/CP: N=N/A; O=Open BB; C=Close BB; B= Open/Close BB;
'                            'F=Floating; 1=Single Open; 2=Single Close; 3=Single Both;
'                            '4=Single Floating; d+Donut; K=Bookend; V=Vignette;
'                            'I=Infomercial; P=Comml Promo (create avails if not found)
''    iUrfCode As Integer     'Last user who modified spot
'End Type
''Shf key record layout
'Type SHFKEY0
'    lChfCode As Long        'Contract code
'    iLineNo As Integer      'Line number
'    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
'    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    iSeqNo As Integer       'Schedule spot seq #
'End Type
''Shf key record layout
'Type SHFKEY1
'    iVefCode As Integer       'Vehicle code
'    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
'    iTime(0 To 1) As Integer 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
'    iSeqNo As Integer       'Schedule spot seq #
'End Type
''Shf key record layout- use LONGKEY0
''Type SHFKEY2
''    lCode As Long    'Internal code number (AutoInc)
''End Type
'Dim tmShf As SHF
Dim tmLongSrchKey As LONGKEY0
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetByKeyForUpdate              *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get record for update           *
'*                     Used after a GetDirect to       *
'*                     position master and slave       *
'*                     for update                      *
'*                                                     *
'*******************************************************
Function gGetByKeyForUpdateXF(slFileName, hlFile As Integer, tlRec As LPOPREC) As Integer
    Dim slName As String
    Dim ilPos As Integer
    Dim ilRecLen As Integer
    Dim ilRet As Integer
    ilPos = InStr(slFileName, ".")
    If ilPos > 0 Then
        slName = Left$(slFileName, ilPos - 1)
    Else
        slName = slFileName
    End If
    Select Case UCase$(slName)
'        Case "SHF"
'            tmShf = tlRec
'            tmLongSrchKey.lCode = tmShf.lCode
'            ilRecLen = Len(tmShf)
'            ilRet = btrGetEqual(hlFile, tlRec, ilRecLen, tmLongSrchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        Case Else
            ilRet = BTRV_ERR_FILE_NOT_FOUND
    End Select
    gGetByKeyForUpdateXF = ilRet
End Function
