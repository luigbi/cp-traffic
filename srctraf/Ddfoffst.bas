Attribute VB_Name = "DDFOFFSTSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: API.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the API declarations
Option Explicit
'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function SendMessage& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Declare Function SendMessageByNum& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam&)
Declare Function SendMessageByString& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam$)

Declare Sub HMemCpy Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Copyright « 1991-1996 Crescent Software, Inc.
'Sort
'Setup fnAV helper function
Const G_MAX_ARRAYDIMS = 60      'VB limit on array dimensions

Type tagAV                      'Array and Vector in 1 compact unit
    PPSA As Long                'Address of pointer to SAFEARRAY
    NumDims As Long         'Number of dimensions
    SCode As Long               'Error info
    Flags As Long               'Reserved
    Subscripts(1 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
End Type

Declare Function fnAV Lib "QPRO32.DLL" (ByRef A() As Any, ParamArray SubscriptsVector()) As tagAV
Public Declare Sub SortT2 Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal ElSize As Long, ByVal MbrOff As Long, ByVal MbrSiz As Long) ' Sub
Public Declare Sub ArraySortStr Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal CaseSensitive As Long)            ' Sub
Public Declare Sub ArraySortTyp Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal ElSize As Long, ByVal MbrOff As Long, ByVal MbrSiz As Long, ByVal CaseSensitive As Long)



Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
End Type
Type DDFNAMES
    sKey As String * 20
    tDDFFile As DDFFILE
End Type
Type DDFFIELD
    iFieldID As Integer         'Field ID
    iFileID As Integer          'File ID from FILE.DDF
    sName As String * 20        'Field Name
    sDataType As String * 1     'Data Type Code (Use ASC to convert)
                                'Code  Description
                                '  0   String
                                '  1   Integer
                                '  2   IEEE Float
                                '  3   Btrieve Date
                                '  4   Btrieve Time
                                '  5   COBOL Decimal COMP-3
                                '  6   COBOL Money
                                '  7   Logical
                                '  8   COBOL Numeric
                                '  9   BASIC BFloat
                                ' 10   Pascal LString
                                ' 11   C ZString
                                ' 12   Variable Length Note
                                ' 13   LVar
                                ' 14   Unsigned Binary
                                ' 15   AutoIncrement
                                ' 16   Bit
                                ' 17   COBOL Numeric STS
    iOffset As Integer          'Field Offset
    iSize As Integer            'Field Size
    sDec As String * 1          'Decimal/Delimiter/Bit Position
    iFlags As Integer           'Case Flag for String data type
End Type
Type DDFFIELD1
    iFileID As Integer          'File ID from FILE.DDF
End Type
Type DDFINDEX
    iFileID As Integer          'File ID from FILE.DDF
    iFieldID As Integer         'Field ID from FIELD.DDF
    iNumber As Integer          'Index Number
    iPart As Integer            'Segment Part Number
    iFlag As Integer            'Btrieve Index Flag
End Type
Type DDFINDEX0
    iFileID As Integer          'File ID from FILE.DDF
End Type
Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type
Type BASEREC
    sChar(1 To 32000) As Byte 'Record
End Type

'*******************************************************
'*                                                     *
'*      Procedure Name:mFileNameFilter                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name                           *
'*                                                     *
'*******************************************************
Function gMakeReplaceName(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    Do
        ilFound = False
        ilPos = InStr(1, slName, " ", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "-", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    gMakeReplaceName = slName
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeUpArrowName                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name                           *
'*                                                     *
'*******************************************************
Function gMakeUpArrowName(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    Do
        ilFound = False
        ilPos = InStr(1, slName, " ", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "^"
            ilFound = True
        End If
    Loop While ilFound
    gMakeUpArrowName = slName
End Function
