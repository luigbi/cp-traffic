Attribute VB_Name = "LINKDLVYSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Linkdlvy.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: LinkDlvy.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for LinkDlvy
Option Explicit

Public igPreFeedType As Integer     '0=Delivery; 1=Engineering
Public sgPreFeedScreenCaption As String
Public sgPreFeedDate As String
Public igPreFeedVefCode As Integer
Public sgPreFeedDay As String


'Delivery file (DLF)
Dim hmDlf As Integer        'Delivery link file
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image
'Prefeed
Dim hmPff As Integer
Dim tmPFF As PFF        'GSF record image
Dim imPffRecLen As Integer        'GSF record length
Dim tmPffSrchKey0 As LONGKEY0
Dim tmPffSrchKey1 As PFFKEY1

'******************************************************************************
' PFF_Pre-Feed Record Definition
'
'******************************************************************************
Type PFF
    lCode                 As Long            ' Pre-Feed Auto Increment reference
                                             ' code
    sType                 As String * 1      ' D=Delivery; E=Engineering
    iVefCode              As Integer         ' Vehicle reference code
    sAirDay               As String * 1      ' 0=M-F; 6=Sa; 7=Su.  Same as
                                             ' defined in Links
    iStartDate(0 To 1)    As Integer         ' Start Date of definitions.  Match
                                             ' Link Dates.
    iFromStartTime(0 To 1) As Integer        ' From Start Time of spots to be
                                             ' Pre-Feed. For Delivery, use Local
                                             ' Time. For Engineering, use Feed
                                             ' Time.
    iFromEndTime(0 To 1)  As Integer         ' End time of spots to be Pre-Feed
    iFromDay              As Integer         ' 0=Mo; 1= Tu; 2= We;.....
    sFromZone             As String * 1      ' E; C; M; P or A (All).  Test for
                                             ' E; C; M and P.
    iToStartTime(0 To 1)  As Integer         ' Start Time to map spots to.
    iToDay                As Integer         ' 0=Mo; 1=Tu; 2=We;....
    sUnused               As String * 20     ' Unused
End Type


'Type PFFKEY0
'    lCode                 As Long
'End Type

Type PFFKEY1
    sType                 As String * 1
    iVefCode              As Integer
    sAirDay               As String * 1
    iStartDate(0 To 1)    As Integer
End Type

Type PFFINFO
    sKey As String * 10
    tPff As PFF
    lTiePffToPbfID As Long
End Type


'******************************************************************************
' PBF_Pre-Feed_Bus Record Definition
'
'******************************************************************************
Type PBF
    lCode                 As Long            ' Pre-Feed Bus Auto Increment
                                             ' Reference code
    lPffCode              As Long            ' PFF Reference Code
    sFromBus              As String * 5      ' From Bus
    sToBus                As String * 5      ' To Bus
    sUnused               As String * 10
End Type


'Type PBFKEY0
'    lCode                 As Long
'End Type

'Type PBFKEY1
'    lPffCode              As Long
'End Type

Type PBFINFO
    sKey As String * 20
    tPbf As PBF
End Type

Type PREFEEDEXPT
    iLogDate0 As Integer
    iLogDate1 As Integer
    iDlfDate0 As Integer
    iDlfDate1 As Integer
    lPffCode As Long
    lFStartTime As Long
    lFEndTime As Long
    iFDay As Integer
    sFZone As String * 1
    lAdjTime As Long        'ToStartTime-FromStartTime
    iTDay As Integer
End Type

'Public igShowHelpMsg As Integer
Type DLFLIST
    DlfRec As DLF
    lDlfCode As Long
    iStatus As Integer  '0=New; 1=old and retain, 2=old and delete; -1= New but not used
    sVehicle As String * 40
    sFeed As String * 20
    sAirTime As String * 10
    sLocalTime As String * 10
    sFeedTime As String * 10
    sSubfeed As String * 20
    sEventName As String * 20
    sEventType As String * 20
End Type


Type DALLASEXPORTSORT
    sKey As String * 20
    sRecord As String * 104
End Type


Public Sub gBuildExportDates(hlDlf As Integer, hlPff As Integer, ilVefCode As Integer, slType As String, llSDate As Long, ilDateAdj As Integer, tlPrefeedExpt() As PREFEEDEXPT)

    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llSuEndDate As Long
    Dim llDate As Long
    Dim slDay As String
    Dim ilDay As Integer
    Dim ilAirDay As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilUpper As Integer
    Dim ilRet As Integer
    Dim llTStartTime As Long
    Dim slSDate As String
    
    ReDim tlPrefeedExpt(0 To 0) As PREFEEDEXPT
    slSDate = Format(llSDate, "m/d/yy")
    ilAirDay = gWeekDayStr(slSDate)
    llStartDate = gDateValue(slSDate)
    llEndDate = llStartDate
    Do While gWeekDayLong(llEndDate) <> 6
        llEndDate = llEndDate + 1
    Loop
    If llEndDate = llStartDate Then
        llEndDate = llEndDate + 1
    End If
    If ilDateAdj = -1 Then
        llStartDate = llStartDate - 1
    Else
        llStartDate = llStartDate - 1   'To handle prefeed that cross mid night
        llEndDate = llEndDate + 1
    End If
    llSuEndDate = llSDate
    Do While gWeekDayLong(llSuEndDate) <> 6
        llSuEndDate = llSuEndDate + 1
    Loop
    imDlfRecLen = Len(tmDlf)
    imPffRecLen = Len(tmPFF)
    For llDate = llStartDate To llEndDate Step 1
        gPackDateLong llDate, ilLogDate0, ilLogDate1
        ilDay = gWeekDayLong(llDate)
        If (ilDay >= 0) And (ilDay <= 4) Then
            slDay = "0"
        ElseIf ilDay = 5 Then
            slDay = "6"
        Else
            slDay = "7"
        End If
        'Obtain the start date of Dlf
        tmDlfSrchKey.iVefCode = ilVefCode
        tmDlfSrchKey.sAirDay = slDay
        tmDlfSrchKey.iStartDate(0) = ilLogDate0
        tmDlfSrchKey.iStartDate(1) = ilLogDate1
        tmDlfSrchKey.iAirTime(0) = 0
        tmDlfSrchKey.iAirTime(1) = 6144 '24*256
        ilRet = btrGetLessOrEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVefCode) And (tmDlf.sAirDay = slDay) Then
            ilDlfDate0 = tmDlf.iStartDate(0)
            ilDlfDate1 = tmDlf.iStartDate(1)
            If (llDate = llStartDate) Or (llDate = llStartDate + 1) Or ((llDate = llStartDate + 2) And (ilDateAdj > 0)) Then
                ilUpper = UBound(tlPrefeedExpt)
                tlPrefeedExpt(ilUpper).iLogDate0 = ilLogDate0
                tlPrefeedExpt(ilUpper).iLogDate1 = ilLogDate1
                tlPrefeedExpt(ilUpper).iDlfDate0 = ilDlfDate0
                tlPrefeedExpt(ilUpper).iDlfDate1 = ilDlfDate1
                tlPrefeedExpt(ilUpper).lPffCode = 0
                tlPrefeedExpt(ilUpper).lFStartTime = 0
                tlPrefeedExpt(ilUpper).lFEndTime = 86400
                tlPrefeedExpt(ilUpper).iFDay = -1
                tlPrefeedExpt(ilUpper).sFZone = ""
                tlPrefeedExpt(ilUpper).lAdjTime = 0
                tlPrefeedExpt(ilUpper).iTDay = -1
                ReDim Preserve tlPrefeedExpt(0 To ilUpper + 1) As PREFEEDEXPT
            End If
            tmPffSrchKey1.sType = slType
            tmPffSrchKey1.iVefCode = ilVefCode
            tmPffSrchKey1.sAirDay = slDay
            tmPffSrchKey1.iStartDate(0) = ilDlfDate0
            tmPffSrchKey1.iStartDate(1) = ilDlfDate1
            ilRet = btrGetEqual(hlPff, tmPFF, imPffRecLen, tmPffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            Do
                If (ilRet = BTRV_ERR_NONE) And (tmPFF.sType = slType) And (tmPFF.iVefCode = ilVefCode) And (tmPFF.sAirDay = slDay) And (tmPFF.iStartDate(0) = ilDlfDate0) And (tmPFF.iStartDate(1) = ilDlfDate1) Then
                    ''If (tmPff.iToDay = ilAirDay) And (gWeekDayLong(llDate) = tmPff.iFromDay) And (llDate >= llSDate) And (llDate <= llSuEndDate) Then
                    'If (tmPff.iToDay >= ilAirDay - 1) And (tmPff.iToDay <= ilAirDay + 1) And (gWeekDayLong(llDate) = tmPff.iFromDay) And (llDate >= llSDate - 1) And (llDate <= llSuEndDate + 1) Then
                    If (gWeekDayLong(llDate) = tmPFF.iFromDay) And (llDate >= llSDate - 1) And (llDate <= llSuEndDate + 1) Then
                        ilUpper = UBound(tlPrefeedExpt)
                        tlPrefeedExpt(ilUpper).iLogDate0 = ilLogDate0
                        tlPrefeedExpt(ilUpper).iLogDate1 = ilLogDate1
                        tlPrefeedExpt(ilUpper).iDlfDate0 = ilDlfDate0
                        tlPrefeedExpt(ilUpper).iDlfDate1 = ilDlfDate1
                        tlPrefeedExpt(ilUpper).lPffCode = tmPFF.lCode
                        gUnpackTimeLong tmPFF.iFromStartTime(0), tmPFF.iFromStartTime(1), False, tlPrefeedExpt(ilUpper).lFStartTime
                        gUnpackTimeLong tmPFF.iFromEndTime(0), tmPFF.iFromEndTime(1), True, tlPrefeedExpt(ilUpper).lFEndTime
                        tlPrefeedExpt(ilUpper).iFDay = tmPFF.iFromDay
                        tlPrefeedExpt(ilUpper).sFZone = tmPFF.sFromZone
                        gUnpackTimeLong tmPFF.iToStartTime(0), tmPFF.iToStartTime(1), False, llTStartTime
                        tlPrefeedExpt(ilUpper).lAdjTime = llTStartTime - tlPrefeedExpt(ilUpper).lFStartTime
                        tlPrefeedExpt(ilUpper).iTDay = tmPFF.iToDay
                        ReDim Preserve tlPrefeedExpt(0 To ilUpper + 1) As PREFEEDEXPT
                    End If
                    ilRet = btrGetNext(hlPff, tmPFF, imPffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    Exit Do
                End If
            Loop
        End If
    Next llDate
    If UBound(tlPrefeedExpt) <= LBound(tlPrefeedExpt) Then
        If ilDateAdj = -1 Then
            llDate = llStartDate - 1
        Else
            '5/31/16: Check date before
            'llDate = llStartDate
            llDate = llStartDate - 1
        End If
        gPackDateLong llDate, ilLogDate0, ilLogDate1
        ilUpper = UBound(tlPrefeedExpt)
        tlPrefeedExpt(ilUpper).iLogDate0 = ilLogDate0
        tlPrefeedExpt(ilUpper).iLogDate1 = ilLogDate1
        tlPrefeedExpt(ilUpper).iDlfDate0 = 0
        tlPrefeedExpt(ilUpper).iDlfDate1 = 0
        tlPrefeedExpt(ilUpper).lPffCode = 0
        tlPrefeedExpt(ilUpper).lFStartTime = 0
        tlPrefeedExpt(ilUpper).lFEndTime = 86400
        tlPrefeedExpt(ilUpper).iFDay = -1
        tlPrefeedExpt(ilUpper).sFZone = ""
        tlPrefeedExpt(ilUpper).lAdjTime = 0
        tlPrefeedExpt(ilUpper).iTDay = -1
        ReDim Preserve tlPrefeedExpt(0 To ilUpper + 1) As PREFEEDEXPT
        If ilDateAdj = -1 Then
            llDate = llStartDate
        Else
            llDate = llStartDate + 1
        End If
        gPackDateLong llDate, ilLogDate0, ilLogDate1
        ilUpper = UBound(tlPrefeedExpt)
        tlPrefeedExpt(ilUpper).iLogDate0 = ilLogDate0
        tlPrefeedExpt(ilUpper).iLogDate1 = ilLogDate1
        tlPrefeedExpt(ilUpper).iDlfDate0 = 0
        tlPrefeedExpt(ilUpper).iDlfDate1 = 0
        tlPrefeedExpt(ilUpper).lPffCode = 0
        tlPrefeedExpt(ilUpper).lFStartTime = 0
        tlPrefeedExpt(ilUpper).lFEndTime = 86400
        tlPrefeedExpt(ilUpper).iFDay = -1
        tlPrefeedExpt(ilUpper).sFZone = ""
        tlPrefeedExpt(ilUpper).lAdjTime = 0
        tlPrefeedExpt(ilUpper).iTDay = -1
        ReDim Preserve tlPrefeedExpt(0 To ilUpper + 1) As PREFEEDEXPT
    End If
End Sub



