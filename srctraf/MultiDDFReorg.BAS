Attribute VB_Name = "MultiDDFReorgSubs"
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

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Sub HMemCpy Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
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


Type BASEREC
    sChar(1 To 32000) As Byte 'Record
End Type


'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Type ZIPNAME
    sKey As String * 60     'Date, Time, Name
    sName As String * 40
    sDate As String * 10
    sTime As String * 10
End Type
Type UNZIPNAME
    sName As String * 40
    sDate As String * 10
    sTime As String * 10
    lSize As Long
    sOldDate As String * 10
    sOldTime As String * 10
    lOldSize As Long
    iIndex As Integer
    sZipName As String * 40
End Type
Type ZIPPEDNAMES
    sName As String * 40
End Type
Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type

Global Const STILL_ACTIVE = &H103
Global Const PROCESS_QUERY_INFORMATION = &H400


Public Sub gShellAndWait(Frm As Form, ByVal slFilePath As String, ilWinStyle As Integer)
    Dim llProcess As Long
    Dim llReturn As Long
    
    llProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(slFilePath, ilWinStyle))
    If llProcess <> 0 Then
        Frm.Enabled = False
        Do
            GetExitCodeProcess llProcess, llReturn
            Sleep 50
            DoEvents
        Loop While llReturn = STILL_ACTIVE
        Frm.Enabled = True
    Else
        MsgBox "Unable to Shell to " & slFilePath, vbOkOnly, "Shell Error"
    End If
End Sub

Public Function gGetTrueFileExt(slItemName As String) As String

    Dim ilPos As Integer
    Dim slTempStr As String
    Dim slName As String
    
    slName = slItemName
    ilPos = InStrRev(slItemName, ".", Len(slItemName) - 1)
    slTempStr = UCase(Mid$(slItemName, ilPos + 1))
    
    If slTempStr = "FWE" Then
        slName = left$(slItemName, ilPos) & "exe"
    End If
    
    If slTempStr = "FWL" Then
        slName = left$(slItemName, ilPos) & "dll"
    End If

    If slTempStr = "FWA" Then
        slName = left$(slItemName, ilPos) & "oca"
    End If

    If slTempStr = "FWX" Then
        slName = left$(slItemName, ilPos) & "ocx"
    End If

    gGetTrueFileExt = slName
    
End Function



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

