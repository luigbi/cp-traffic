Attribute VB_Name = "EXPSTNFDSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Expstnfd.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpAffFd.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ExpStnFd subs and functions
Option Explicit
Option Compare Text
'Public igBrowserType As Integer '0=Bmp (bit map); 1=Csv (comma delimited); 2= Txt (text)
'Public igBrowserReturn As Integer   '0=Cancelled; 1=Ok
'Public sgBrowserFile As String  'Drive\Path\FileName if igBrowserReturn=1 of the file
Type SCHSPOTINFO
    sKey As String * 20     'Date; Time; SeqNo
    tCpr As CPR
End Type
Public tgSchSpotInfo() As SCHSPOTINFO
Type STNINFO
    'sType As String * 1  'S=Station Record (mGetStnInfo); G=General (mNonRotFileName)
    'sCallFreq As String * 20
    'iAirVeh As Integer
    'lRegionCode As Long
    'sSiteID As String * 10
    'sEDAS As String * 20
    'sTransportal As String * 20
    'lRafCode As Long
    'sFileName As String * 20
    'sFdZone As String * 3
    'iLkStnInfo As Integer
    sType As String * 1  'S=Station Record (mGetStnInfo); G=General (mNonRotFileName)
    sCallLetter As String * 4
    sBand As String * 2
    iAirVeh As Integer
    lRegionCode As Long
    sSiteID As String * 10
    sEDAS(0 To 5) As String * 10
    sTransportal(0 To 5) As String * 10
    sKCNo As String * 10
    lRafCode As Long
    sFileName As String * 20
    sStnFdCode As String * 2
    sFdZone As String * 3
    iAirPlays As Integer
    sCmmlLogReq As String * 1
    sCmmlLogPledge(0 To 9) As String * 30
    sKCEnvCopy As String * 1
    sCmmlLogDPType As String * 1
    sCmmlLogCart As String * 1
    iLkStnInfo As Integer
    iLkCartInfo1 As Integer
    iLkCartInfo2 As Integer
End Type
Type CARTSTNXREF
    lCifCode As Long
    sShortTitle As String * 15
    sISCI As String * 20
    iLkCartInfo1 As Integer
    iLkCartInfo2 As Integer
    iAdfCode As Integer
    iLen As Integer
    iFdDateNew As Integer
End Type
Public tgCartStnXRef() As CARTSTNXREF
Public igSGOrKC As Integer  '0=StarGuide; 1=KenCast


Type DUPLCOMMENT
    iNoteNo As Integer
    iAdfCode As Integer
    lChfCode As Long
    sISCI As String * 20
End Type



'*******************************************************
'*                                                     *
'*      Procedure Name:gAdvtNameFilter                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Replace , with space           *
'*                                                     *
'*******************************************************
Function gAdvtNameFilter(slInName As String) As String
    Dim slName As String
    'slName = slInName
    'Do
    '    ilFound = False
    '    ilPos = InStr(1, slName, ",", 1)
    '    If ilPos > 0 Then
    '        Mid$(slName, ilPos, 1) = " "
    '        ilFound = True
    '    End If
    'Loop While ilFound
    slName = gFileNameFilter(slInName)
    gAdvtNameFilter = slName
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFileNameFilter                 *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove illegal characters from *
'*                      name                           *
'*                                                     *
'*******************************************************
Function gFileNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    slName = slInName
    'Remove " and '
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        'Dan M 10/27/14 added &
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    gFileNameFilter = slName
End Function
Public Function gXDSShortTitle(tlAdf As ADF, slProduct As String, blNoAbbreviation As Boolean, blAddProduct As Boolean) As String
    '7219
    'matches affiliate, except uses tlAdf instead of llAdfSearch
    'blAddProduct only if blNoAbbreviation is false (blNoAbbreviation is for 'CU')
    Dim slShortTitle As String
'    Dim llAdf As Long
    
'    slShortTitle = Trim$(slProduct)
'    llAdf = gBinarySearchAdf(llAdfSearch)
'    If llAdf <> -1 Then
        If blNoAbbreviation Then
            slShortTitle = Trim$(tlAdf.sName)
        Else
            slShortTitle = Trim$(Left$(tlAdf.sAbbr, 6))
            If slShortTitle = "" Then
                slShortTitle = Trim$(Left$(tlAdf.sName, 6))
            End If
            If blAddProduct Then
                '7555 got rid of space  ", "
                slShortTitle = slShortTitle & "," & Trim$(slProduct)
            End If
        End If
'    End If
    gXDSShortTitle = UCase$(gFileNameFilter(slShortTitle))
End Function
