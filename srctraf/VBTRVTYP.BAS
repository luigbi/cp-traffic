' Copyright 1992, 1993 Classic Software, Inc. All rights reserved.
' Btrieve is a registered trademark of Novell, Inc.
'
' File Name: BTRVTYPE.H
'
' Release: 2.0
'
' Description:
'   This provides a lot of useful definitions for use with Btrieve. An
'   application developer who is using the VBTRV library does NOT
'   need to include this for his compilations. With the exception of the
'   Stat operation none of these defines should be needed. In the case of
'   the Stat operation only the file and key specification defines are
'   needed. The File, supplemental index and extended operation functions
'   provided by the vbtrv library are an easier way of handling the creation
'   and interpretation of these complex structures.


Const BTRV_MAX_KEYS = 24

' structs for creating files.

'Type sFILESPEC
'    RecLength   As Integer
'    PageSize    As Integer
'    NumIndexes  As Integer
'    ulNumRecs   As Long
'    uFileFlags  As Integer
'    uReserved   As Integer
'    uAllocation As Integer
'End Type

'Type sKEYSPEC
'    KeyPosition    As Integer
'    KeyLength      As Integer
'    KeyFlags       As Integer
'    ulNumKeys      As Long
'    cExtendKeyType As String * 1
'    cNullValue     As String * 1
'    ulReserved     As Long
'End Type

'Type sCOLLATE
'    SignatureByte  As String * 1     ' always 0xac
'    SequenceName   As String * 8
'    CollateSeq     As String * 256
'    align          As String * 1     ' Force even allignment
'End Type


'Type sFILECREATE
'    sFILESPEC As sFILESPEC
'    sKeySpecs(24) As sKEYSPEC
'    sCOLLATE  As sCOLLATE
'End Type

' Extended operation structures

' Header and filter fixed portion
'Type sEXTEND
'    DataBuffSize  As Integer
'    ConstE        As String * 1   '  always 'E'
'    ConstG        As String * 1   ' always 'G'
'    MaxSkipped    As Integer
'    NumLogicTerms As Integer
'End Type

' Logic Expressions
'Type sLOGIC
'    DataType          As String * 1
'    FieldLength       As Integer
'    FieldOffset       As Integer
'    ComparisonCode    As String * 1
'    AndOrLogic        As String * 1
'    SecondFieldOffset As Integer
'End Type

'Type sLOGIC_CONST
'    DataType          As String * 1
'    FieldLength       As Integer
'    FieldOffset       As Integer
'    ComparisonCode    As String * 1
'    AndOrLogic        As String * 1
'End Type

' Field extraction

Type sEXTRACTOR
    RecsRetrieved   As Integer
    FieldsExtracted As Integer
End Type

Type sFIELD_EXTRACTOR
    FieldLength As Integer
    FieldOffset As Integer
End Type

Type sREC_BUFF
    RecordLength As Integer
    RecordPosition As Long
End Type

Type sEXT_DATA
    RecsReturned As Integer
End Type

Type sBTRV_VERSION
    VersionNo   As Integer     ' Version Number
    RevisionNo  As Integer     ' Revision Number
    OS_Type     As String * 1  ' N=Server and DOS, P=OS/2, W=Windows
End Type

'Type sDATE
'   Day As String * 1
'   Month As String * 1
'   Year As Integer
'End Type

'Type sTIME
'   Hundredths As String * 1
'   Seconds    As String * 1
'   Minutes    As String * 1
'   Hours      As String * 1
'End Type

