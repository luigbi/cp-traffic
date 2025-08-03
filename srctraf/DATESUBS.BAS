Attribute VB_Name = "DATESUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of DATESUBS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: DateSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
Public sgSysDate As String
Public Function gCSIGetName() As String
    gCSIGetName = sgSysDate
End Function
Public Function gCsiSetName(slSYSName As String) As Integer
    sgSysDate = Trim$(slSYSName)
    gCsiSetName = 0
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAddDayToDate                   *
'*                                                     *
'*             Created:7/27/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Day of week to string date  *
'*                                                     *
'*******************************************************
Function gAddDayToDate(slInpDate As String) As String
'
'   sRet = gAddDayToDate(sDate)
'   Where:
'       sDate (I)- String date (Format as "m/d/yy")
'       sRet (O)- String Date with day
                '
'    Dim dtlDate As Double
'    dtlDate = gDateValue(slAnyDate)
'    gAddDayToDate = Format$(dtlDate, "ddd, m/d/yy")
    Dim slDate As String
    Dim slAnyDate As String



    slAnyDate = Trim$(slInpDate)
    On Error GoTo gAddDayToDateErr
    slDate = Format$(gAdjYear(slAnyDate), "m/d/yy")
    On Error GoTo 0
    Select Case gWeekDayStr(slAnyDate)
        Case 0  'Monday
            slDate = "Mo, " & slDate
        Case 1  'Tuesday
            slDate = "Tu, " & slDate
        Case 2  'Wednesday
            slDate = "We, " & slDate
        Case 3  'Thursday
            slDate = "Th, " & slDate
        Case 4  'Friday
            slDate = "Fr, " & slDate
        Case 5  'Saturday
            slDate = "Sa, " & slDate
        Case 6  'Sunday
            slDate = "Su, " & slDate
    End Select
    gAddDayToDate = slDate
    On Error GoTo 0
    Exit Function
gAddDayToDateErr:
    gAddDayToDate = ""
    On Error GoTo 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAddLengths                     *
'*                                                     *
'*             Created:10/1/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute the sum of two lengths  *
'*                                                     *
'*******************************************************
Sub gAddLengths(slLength1 As String, slLength2 As String, slFormat As String, slSum As String)
'
'   gAddLengths slLen1, slLen2, slFormat, slLen
'   Where:
'       slLen1 (I)- First length to be added
'       slLen2 (I)- Second length to be added
'       slFormat (I)-"1" = hh:mm:ss
'                           "2" = hh mm'ss"
'                           "3" = hhHmmMssS
'       slLen (O) - Length of Time in format
'
    ReDim ilLength1(0 To 1) As Integer
    ReDim ilLength2(0 To 1) As Integer
    Dim ilSec1 As Integer    'Seconds
    Dim ilMin1 As Integer    'Minutes
    Dim ilHour1 As Integer   'Hours
    Dim ilSec2 As Integer    'Seconds
    Dim ilMin2 As Integer    'Minutes
    Dim ilHour2 As Integer   'Hours
    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer

    gPackLength Trim$(slLength1), ilLength1(0), ilLength1(1)
    gPackLength Trim$(slLength2), ilLength2(0), ilLength2(1)
    ilSec1 = ilLength1(0) \ 256    'Obtain seconds
    ilMin1 = ilLength1(1) And &HFF 'Obtain Minutes
    ilHour1 = ilLength1(1) \ 256   'Obtain month
    ilSec2 = ilLength2(0) \ 256    'Obtain seconds
    ilMin2 = ilLength2(1) And &HFF 'Obtain Minutes
    ilHour2 = ilLength2(1) \ 256   'Obtain month
    ilSec = ilSec1 + ilSec2
    ilMin = ilMin1 + ilMin2
    ilHour = ilHour1 + ilHour2
    Do While ilSec > 59
        ilMin = ilMin + 1
        ilSec = ilSec - 60
    Loop
    Do While ilMin > 59
        ilHour = ilHour + 1
        ilMin = ilMin - 60
    Loop
    ilHsSec = ilSec * 256 'High order byte = hundredths of seconds; low order byte = seconds
    ilMinHr = ilMin + ilHour * 256 'High order byte = minute; low order byte = hours
    gUnpackLength ilHsSec, ilMinHr, slFormat, False, slSum
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gAddTimeLength                  *
'*                                                     *
'*             Created:10/1/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute the sum of time plus    *
'*                     length                          *
'*                                                     *
'*******************************************************
Sub gAddTimeLength(slTime1 As String, slLength2 As String, slStyle As String, slFormat As String, slSum As String, slXMid As String)
'
'   gAddTimeLength slTime1, slLen2, slStyle, slFormat, slTime, slXMid
'   Where:
'       slTime1 (I)- Time to be added
'       slLen2 (I)- length to be added to time
'       slStyle (I)- "A" = AM or PM style
'                       "M" = Military style
'       slFormat (I)- "1" = hours, min, sec
'                     "2" = hours, minutes (no seconds)
'                     "3" = minutes, seconds (no hours)
'       slTime (O) - Time formatted
'       slXMid(O)- Cross Midnight (N or Y)
'
    ReDim ilTime1(0 To 1) As Integer
    ReDim ilLength2(0 To 1) As Integer
    Dim ilSec1 As Integer    'Seconds
    Dim ilMin1 As Integer    'Minutes
    Dim ilHour1 As Integer   'Hours
    Dim ilSec2 As Integer    'Seconds
    Dim ilMin2 As Integer    'Minutes
    Dim ilHour2 As Integer   'Hours
    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer

    gPackTime Trim$(slTime1), ilTime1(0), ilTime1(1)
    gPackLength Trim$(slLength2), ilLength2(0), ilLength2(1)
    ilSec1 = ilTime1(0) \ 256    'Obtain seconds
    ilMin1 = ilTime1(1) And &HFF 'Obtain Minutes
    ilHour1 = ilTime1(1) \ 256   'Obtain hour
    ilSec2 = ilLength2(0) \ 256    'Obtain seconds
    ilMin2 = ilLength2(1) And &HFF 'Obtain Minutes
    ilHour2 = ilLength2(1) \ 256   'Obtain month
    ilSec = ilSec1 + ilSec2
    ilMin = ilMin1 + ilMin2
    ilHour = ilHour1 + ilHour2
    Do While ilSec > 59
        ilMin = ilMin + 1
        ilSec = ilSec - 60
    Loop
    Do While ilMin > 59
        ilHour = ilHour + 1
        ilMin = ilMin - 60
    Loop
    slXMid = "N"
    If ilHour >= 24 Then
        slXMid = "Y"
        ilHour = ilHour - 24
    End If
    ilHsSec = ilSec * 256 'High order byte = hundredths of seconds; low order byte = seconds
    ilMinHr = ilMin + ilHour * 256 'High order byte = minute; low order byte = hours
    gUnpackTime ilHsSec, ilMinHr, slStyle, slFormat, slSum
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjYear                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Adjust year (70-99 maps to 19xx;*
'*                       00-69 maps to 20xx)           *
'*                                                     *
'*******************************************************
Function gAdjYear(slInpDate As String) As String
'
'   sRetDate = gAdjYear(sDate)
'   Where:
'       sDate (I)- Date to adjust year
'       sRetDate (O)- Date with adjusted year
'
    Dim slMonthDay As String
    Dim slTMonthDay As String
    Dim slYear As String
    Dim llYear As Long
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slAnyDate As String
    Dim slChar As String
    Dim ilLoop As Integer

    slAnyDate = Trim$(slInpDate)
    On Error GoTo gAdjYearErr
    ilPos1 = InStr(1, slAnyDate, "/")
    If ilPos1 <= 0 Then
        gAdjYear = slAnyDate
        Exit Function
    End If
    ilPos2 = InStr(ilPos1 + 1, slAnyDate, "/")
    If ilPos2 < 0 Then
        gAdjYear = slAnyDate
        Exit Function
    ElseIf ilPos2 = 0 Then
        slTMonthDay = slAnyDate
        slYear = str$(Year(gNow()))
    Else
        slTMonthDay = Left$(slAnyDate, ilPos2 - 1)
        slYear = right$(slAnyDate, Len(slAnyDate) - ilPos2)
        'The above line could also be:
        'slYear = Mid$(slAnyDate, ilPos2+1)
        'I don't know which is the fastest
    End If
    '1/14/08:  Remove charaters after blank as they might contain time (mm/dd/yy hh:mm:ss a/p)
    ilPos1 = InStr(1, slYear, " ", vbTextCompare)
    If ilPos1 > 2 Then
        slYear = Left$(slYear, ilPos1 - 1)
    End If
    '8-20-01 remove blanks from date or it will error out in Day function
    slMonthDay = ""
    For ilLoop = 1 To Len(slTMonthDay) Step 1
        slChar = Mid$(slTMonthDay, ilLoop, 1)
        If slChar <> " " Then
            slMonthDay = slMonthDay & slChar
        End If
    Next ilLoop

    llYear = Val(slYear)
    If (llYear >= 0) And (llYear <= 69) Then
        llYear = 2000 + llYear
    ElseIf (llYear >= 70) And (llYear <= 99) Then
        llYear = 1900 + llYear
    End If
    gAdjYear = slMonthDay & "/" & Trim$(str$(llYear))
    Exit Function
gAdjYearErr:
    On Error GoTo 0
    gAdjYear = slAnyDate
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gConvertTime                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format time so it can be used   *
'*                     VB time procedures              *
'*                                                     *
'*******************************************************
Function gConvertTime(slInpTime As String) As String
'
'   sRetTime = gConvertTime(sTime)
'   Where:
'       sTime (I)- Time string to be checked and formatted
'       sRetTime (O)- Formatted time
'
    Dim slFixedTime As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim ilPos3 As Integer
    Dim slTime As String

    slTime = Trim$(slInpTime)
    If slTime = "" Then
        slTime = "12:00AM"
    End If
    slFixedTime = UCase$(slTime)
    ilPos1 = InStr(slFixedTime, "N")
    If ilPos1 <> 0 Then
        slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "PM" & Mid$(slFixedTime, ilPos1 + 1)
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    If (ilPos1 = 0) And (ilPos2 = 0) Then
        ilPos1 = InStr(slFixedTime, "M")
        If ilPos1 <> 0 Then
            slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "AM" & Mid$(slFixedTime, ilPos1 + 1)
        End If
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    ilPos3 = InStr(slFixedTime, "M")
    If (ilPos3 = 0) And ((ilPos1 <> 0) Or (ilPos2 <> 0)) Then
        slFixedTime = slFixedTime & "M"
    End If
    If InStr(slTime, ":") = 0 Then
        If ilPos1 <> 0 Then
            If Len(slFixedTime) <= 4 Then
                slFixedTime = Left$(slFixedTime, ilPos1 - 1) & ":00AM"
            Else
                If Len(slFixedTime) <= 6 Then
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                Else
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & right$(slFixedTime, 6)
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                End If
            End If
        Else
            If ilPos2 <> 0 Then
                If Len(slFixedTime) <= 4 Then
                    slFixedTime = Left$(slFixedTime, ilPos2 - 1) & ":00PM"
                Else
                    If Len(slFixedTime) <= 6 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & right$(slFixedTime, 6)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                    End If
                End If
            Else
                If Len(slFixedTime) <= 2 Then
                    slFixedTime = slFixedTime & ":00"
                Else
                    If Len(slFixedTime) <= 4 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & right$(slFixedTime, 2)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & right$(slFixedTime, 2)
                    End If
                End If
            End If
        End If
    End If
    gConvertTime = slFixedTime
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCurrencyToLength               *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert currency to length      *
'*                                                     *
'*******************************************************
Function gCurrencyToLength(clAnyLength As Currency) As String
'
'   sRetLength = gCurrencyToLength(clLength)
'   Where:
'       clLength (I)- Length as currency to be converted to a string
'       sRetLength (O)- Length (hhHmmMssS format)
'
    Dim slLength As String
    Dim clLength As Currency
    Dim llLength As Long
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilSec As Integer
    On Error GoTo gCurrencyToLengthErr
    clLength = clAnyLength
    Do While clLength < 0
        clLength = clLength + 86400
    Loop
    llLength = clLength '+ .5 not required as conversion rounds up
    ilHour = llLength \ 3600
    llLength = llLength Mod 3600
    ilMin = llLength \ 60
    ilSec = llLength Mod 60
    slLength = ""
    If (ilHour <> 0) Or (ilMin <> 0) Or (ilSec <> 0) Then
        If ilHour <> 0 Then
            slLength = Trim$(str$(ilHour)) & "h"
        End If
        If ilMin <> 0 Then
            slLength = slLength & Trim$(str$(ilMin)) & "m"
        End If
        If ilSec <> 0 Then
            slLength = slLength & Trim$(str$(ilSec)) & "s"
        End If
    Else
        slLength = "0s"
    End If
    gCurrencyToLength = slLength
    Exit Function
gCurrencyToLengthErr:
    On Error GoTo 0
    gCurrencyToLength = ""
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCurrencyToTime                 *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert currency to time        *
'*                                                     *
'*******************************************************
Function gCurrencyToTime(clAnyTime As Currency) As String
'
'   sRetTime = gCurrencyToTime(clTime)
'   Where:
'       clTime (I)- Time as currency to be converted to a string
'       sRetTime (O)- Time (hh:mm:ssAM or PM format)
'
    Dim slTime As String
    Dim clTime As Currency
    Dim llTime As Long
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilSec As Integer

    On Error GoTo gCurrencyToTimeErr
    clTime = clAnyTime
    Do While clTime < 0
        clTime = clTime + 86400
    Loop
    llTime = clTime '+ .5 not required as conversion round number
    ilHour = llTime \ 3600
    Do While ilHour > 23
        ilHour = ilHour - 24
    Loop
    llTime = llTime Mod 3600
    ilMin = llTime \ 60
    ilSec = llTime Mod 60
    slTime = ""
    If (ilHour = 0) Or (ilHour = 12) Then
        slTime = "12"
    Else
        If ilHour < 12 Then
            slTime = Trim$(str$(ilHour))
        Else
            slTime = Trim$(str$(ilHour - 12))
        End If
    End If
    If ilMin <> 0 Then
        If ilMin < 10 Then
            slTime = slTime & ":0" & Trim$(str$(ilMin))
        Else
            slTime = slTime & ":" & Trim$(str$(ilMin))
        End If
    End If
    If ilSec <> 0 Then
        If ilMin = 0 Then
            slTime = slTime & ":00:"
        Else
            slTime = slTime & ":"
        End If
        If ilSec < 10 Then
            slTime = slTime & "0" & Trim$(str$(ilSec))
        Else
            slTime = slTime & Trim$(str$(ilSec))
        End If
    End If
    If ilHour < 12 Then
        slTime = slTime & "AM"
    Else
        slTime = slTime & "PM"
    End If
    gCurrencyToTime = slTime
    Exit Function
gCurrencyToTimeErr:
    On Error GoTo 0
    gCurrencyToTime = ""
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gDateValue                      *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain serial date              *
'*                     (in VB format)                  *
'*                                                     *
'*******************************************************
Function gDateValue(slInpDate As String) As Long
'
'   lRetDate = gDateValue(sDate)
'   Where:
'       sDate (I)- Date for which to obtain serial value
'       lRetDate (O)- Date as serial number
'
    Dim slDate As String
    Dim slAnyDate As String
    Dim llDate As Long



    slAnyDate = Trim$(slInpDate)
    If Trim$(slAnyDate) = "" Then
        gDateValue = 0
        Exit Function
    End If
    'If Not gValidDate(slAnyDate) Then
    '    gDateValue = 0
    '    Exit Function
    'End If

    slDate = gAdjYear(slAnyDate)

    'The following code was taken from gValidDate
    On Error GoTo gDateValueErr
    llDate = Day(slDate)
    llDate = Month(slDate)
    llDate = Year(slDate)
    If (llDate < 1970) Or (llDate > 2069) Then
        gDateValue = 0
        Exit Function
    End If
    'End if code taken from gValidDate

    gDateValue = DateValue(slDate)
    Exit Function
gDateValueErr:
    On Error GoTo 0
    gDateValue = 0
    Exit Function

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mDayNames                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Days                     *
'*                                                     *
'*******************************************************
Function gDayNames(ilCffDay() As Integer, slCffXDay() As String * 1, ilCharsNo As Integer, slEDIDays As String) As String
    Dim slStr As String
    Dim slDayImage As String
    Dim ilDay As Integer
    Dim ilNoChars As Integer
    Dim ilSvNoChars As Integer

    ilNoChars = ilCharsNo
    ilSvNoChars = ilNoChars
    slStr = ""
    slEDIDays = slStr
    For ilDay = 0 To 6 Step 1
        If ilCffDay(ilDay) > 0 Then
            slStr = slStr & "Y"
            slEDIDays = slEDIDays & "Y"
            Select Case ilDay
                Case 0
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left("Monday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Monday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 1
                    If slDayImage = "" Then
                        slDayImage = Left("Tuesday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Tuesday", ilNoChars)
                    End If
                Case 2
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left$("Wednesday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Wednesday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 3
                    If slDayImage = "" Then
                        slDayImage = Left$("Thursday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Thursday", ilNoChars)
                    End If
                Case 4
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left$("Friday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Friday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 5
                    If slDayImage = "" Then
                        slDayImage = Left$("Saturday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Saturday", ilNoChars)
                    End If
                Case 6
                    If slDayImage = "" Then
                        slDayImage = Left$("Sunday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Sunday", ilNoChars)
                    End If
            End Select
        Else
            slStr = slStr & " "
            slEDIDays = slEDIDays & "N"
        End If
    Next ilDay
    For ilDay = 0 To 6 Step 1
        If slCffXDay(ilDay) = "1" Then
            Mid$(slStr, ilDay + 1, 1) = "Y"
            Mid$(slEDIDays, ilDay + 1, 1) = "Y"
            Select Case ilDay
                Case 0
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left("Monday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Monday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 1
                    If slDayImage = "" Then
                        slDayImage = Left("Tuesday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Tuesday", ilNoChars)
                    End If
                Case 2
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left$("Wednesday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Wednesday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 3
                    If slDayImage = "" Then
                        slDayImage = Left$("Thursday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Thursday", ilNoChars)
                    End If
                Case 4
                    If ilNoChars = 2 Then
                        ilNoChars = 1
                    End If
                    If slDayImage = "" Then
                        slDayImage = Left$("Friday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Friday", ilNoChars)
                    End If
                    ilNoChars = ilSvNoChars
                Case 5
                    If slDayImage = "" Then
                        slDayImage = Left$("Saturday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Saturday", ilNoChars)
                    End If
                Case 6
                    If slDayImage = "" Then
                        slDayImage = Left$("Sunday", ilNoChars)
                    Else
                        slDayImage = slDayImage & "," & Left$("Sunday", ilNoChars)
                    End If
            End Select
        End If
    Next ilDay
'    If slStr = "YYYYYYY" Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
'        Else
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", 2)
'        End If
'    ElseIf slStr = "YYYYYY " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
'        Else
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", 2)
'        End If
'    ElseIf slStr = "YYYYY  " Then
'        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
'    ElseIf slStr = "YYYY   " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
'        Else
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", 2)
'        End If
'    ElseIf slStr = "YYY    " Then
'        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
'    ElseIf slStr = "YY     " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", ilNoChars)
'        Else
'            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", 2)
'        End If
'    ElseIf slStr = " YYYYYY" Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
'        Else
'            slStr = Left$("Tuesday", 2) & "-" & Left$("Sunday", 2)
'        End If
'    ElseIf slStr = " YYYYY " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
'        Else
'            slStr = Left$("Tuesday", 2) & "-" & Left$("Saturday", 2)
'        End If
'    ElseIf slStr = " YYYY  " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
'        Else
'            slStr = Left$("Tuesday", 2) & "-" & Left$("Friday", ilNoChars)
'        End If
'    ElseIf slStr = " YYY   " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
'        Else
'            slStr = Left$("Tuesday", 2) & "-" & Left$("Thursday", 2)
'        End If
'    ElseIf slStr = " YY    " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
'        Else
'            slStr = Left$("Tuesday", 2) & "-" & Left$("Wednesday", ilNoChars)
'        End If
'    ElseIf slStr = "  YYYYY" Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
'        Else
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", 2)
'        End If
'    ElseIf slStr = "  YYYY " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
'        Else
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", 2)
'        End If
'    ElseIf slStr = "  YYY  " Then
'        slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
'    ElseIf slStr = "  YY   " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
'        Else
'            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", 2)
'        End If
'    ElseIf slStr = "   YYYY" Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
'        Else
'            slStr = Left$("Thursday", 2) & "-" & Left$("Sunday", 2)
'        End If
'    ElseIf slStr = "   YYY " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
'        Else
'            slStr = Left$("Thursday", 2) & "-" & Left$("Saturday", 2)
'        End If
'    ElseIf slStr = "   YY  " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
'        Else
'            slStr = Left$("Thursday", 2) & "-" & Left$("Friday", ilNoChars)
'        End If
'    ElseIf slStr = "     Y " Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Saturday", ilNoChars)
'        Else
'            slStr = Left$("Saturday", 2)
'        End If
'    ElseIf slStr = "     YY" Then
'        slStr = Left$("Saturday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
'    ElseIf slStr = "      Y" Then
'        If ilNoChars > 1 Then
'            slStr = Left$("Sunday", ilNoChars)
'        Else
'            slStr = Left$("Sunday", 2)
'        End If
'    Else
'        slStr = slDayImage
'    End If
    gDayNames = gDayStrToNames(slStr, slDayImage, ilNoChars)
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gDecOneDay                      *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Decrement String Date by one day*
'*                                                     *
'*******************************************************
Function gDecOneDay(slInpDate As String) As String
'
'   sRet = gDecOneDay(sDate)
'   Where:
'       sDate (I)- String date (Format as "m/d/yy")
'       sRet (O)- String Date decremented by one day
'
    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate) - 1
    gDecOneDay = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gDecOneWeek                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Decrement String Date by one    *
'*                     week                            *
'*                                                     *
'*******************************************************
Function gDecOneWeek(slInpDate As String) As String
'
'   sRetDate = gDecOneWeek(sAnyDate)
'   Where:
'       sAnyDate (I)- String Date to be decremented formatted "m/d/yy"
'       sRetDate (O)- String date one week less then sAnyDate
'
    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate) - 7
    gDecOneWeek = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatDate                     *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format Date so it can be        *
'*                     viewed                          *
'*                                                     *
'*******************************************************
Function gFormatDate(slInpDate As String) As String
'
'   sRetDate = gFormatDate(sDate)
'   Where:
'       sDate (I)- Date string to be checked and formatted
'       sRetDate (O)- Formatted date
'

    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim slFormatDate As String
    Dim slDate As String

    slDate = Trim$(slInpDate)
    If Not gValidDate(slDate) Then
        gFormatDate = ""
        Exit Function
    End If
    gPackDate slDate, ilDyMn, ilYear
    gUnpackDate ilDyMn, ilYear, slFormatDate    'add year if missing
    gFormatDate = slFormatDate
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatLength                   *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format Length so it can be      *
'*                     viewed                          *
'*                                                     *
'*******************************************************
Function gFormatLength(slInpLength As String, slFormat As String, ilTest120 As Integer) As String
'
'   sRetLength = gFormatLength(sLength, slFormat, ilTest120)
'   Where:
'       sLength (I)- Length string to be checked and formatted
'       slFormat (I)-"1" = hh:mm:ss
'                           "2" = hh mm'ss"
'                           "3" = hhHmmMssS
'       ilTest120 (I)- True=Test if length is 120 or less and show as seconds
'       sRetLength (O)- Formatted length
'

    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slFormatLength As String
    Dim slLength As String

    slLength = Trim$(slInpLength)
    If Not gValidLength(slLength) Then
        gFormatLength = ""
        Exit Function
    End If
    gPackLength slLength, ilHsSec, ilMinHr
    gUnpackLength ilHsSec, ilMinHr, slFormat, ilTest120, slFormatLength
    gFormatLength = slFormatLength
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatTime                     *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format Time so it can be        *
'*                     viewed                          *
'*                                                     *
'*******************************************************
Function gFormatTime(slInpTime As String, slStyle As String, slFormat As String) As String
'
'   sRetTime = gFormatTime(sTime, slStyle, slFormat)
'   Where:
'       sTime (I)- Time string to be checked and formatted
'       slStyle (I)- "A" = AM or PM style
'                    "M" = Military style (HHMM:SS)
'       slFormat (I)- "1" = hours, min, sec
'                     "2" = hours, minutes (no seconds)
'                     "3" = minutes, seconds (no hours)
'       sRetTime (O)- Formatted time
'

    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slFormatTime As String
    Dim slTime As String

    slTime = Trim$(slInpTime)
    If Not gValidTime(slTime) Then
        gFormatTime = ""
        Exit Function
    End If
    gPackTime slTime, ilHsSec, ilMinHr
    gUnpackTime ilHsSec, ilMinHr, slStyle, slFormat, slFormatTime
    gFormatTime = slFormatTime
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatTime                     *
'*                                                     *
'*             Created:10/27/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format Time so it can be        *
'*                     viewed                          *
'*                                                     *
'*******************************************************
Function gFormatTimeLong(llInpTime As Long, slStyle As String, slFormat As String) As String
'
'   sRetTime = gFormatTime(llTime, slStyle, slFormat)
'   Where:
'       llTime (I)- Time as long to be checked and formatted
'       slStyle (I)- "A" = AM or PM style
'                    "M" = Military style (HHMM:SS)
'       slFormat (I)- "1" = hours, min, sec
'                     "2" = hours, minutes (no seconds)
'                     "3" = minutes, seconds (no hours)
'       sRetTime (O)- Formatted time
'

    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slFormatTime As String

    gPackTimeLong llInpTime, ilHsSec, ilMinHr
    gUnpackTime ilHsSec, ilMinHr, slStyle, slFormat, slFormatTime
    gFormatTimeLong = slFormatTime
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetCorpCalIndex                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine index into tmMCof     *
'*                                                     *
'*******************************************************
Function gGetCorpCalIndex(ilYear As Integer) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gObtainCorpCal()
    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
        If tgMCof(ilLoop).iYear = ilYear Then
            gGetCorpCalIndex = ilLoop
            Exit Function
        End If
    Next ilLoop
    gGetCorpCalIndex = -1
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIncOneDay                      *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Increment String Date by one day*
'*                                                     *
'*******************************************************
Function gIncOneDay(slInpDate As String) As String
'
'   sRet = gIncOneDay(sDate)
'   Where:
'       sDate (I)- String date (Format as "m/d/yy")
'       sRet (O)- String Date incremented by one day
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate) + 1
    gIncOneDay = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIncOneWeek                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Increment String Date by one   *
'*                      week                           *
'*                                                     *
'*******************************************************
Function gIncOneWeek(slInpDate As String) As String
'
'   sRetDate = gIncOneWeek(sAnyDate)
'   Where:
'       sAnyDate (I)- String Date to be incremented formatted "m/d/yy"
'       sRetDate (O)- String date one week greater then sAnyDate
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate) + 7
    gIncOneWeek = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLengthToCurrency               *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert length to currency      *
'*                     (for precision -                *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gLengthToCurrency(slInpLength As String) As Currency
'
'   clRetLength = gLengthToCurrency(slLength)
'   Where:
'       slLength (I)- Length as string to be converted to currency
'       clRetLength (O)- length as currency
'
    Dim slLength As String
    Dim clLength As Currency
    Dim ilPos As Integer
    Dim ilLoc As Integer
    Dim slAnyLength As String

    slAnyLength = Trim$(slInpLength)
    On Error GoTo gLengthToCurrencyErr
    ilPos = InStr(slAnyLength, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gLengthToCurrency = 0
            Exit Function
        End If
        slLength = UCase$(Mid$(slAnyLength, 2))
    Else
        slLength = UCase$(slAnyLength)
    End If
    clLength = 0
    ilLoc = InStr(1, slLength, "H", 1)
    If ilLoc <> 0 Then
        clLength = clLength + 3600 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "M", 1)
    If ilLoc <> 0 Then
        clLength = clLength + 60 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "S", 1)
    If ilLoc <> 0 Then
        clLength = clLength + Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    Else
        clLength = clLength + Val(slLength)
    End If
    If ilPos = 0 Then
        gLengthToCurrency = clLength
    Else
        gLengthToCurrency = -clLength
    End If
    Exit Function
gLengthToCurrencyErr:
    On Error GoTo 0
    gLengthToCurrency = 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gMonthName                      *
'*                                                     *
'*             Created:11/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain month name from date     *
'*                                                     *
'*******************************************************
Function gMonthName(slInpDate As String) As String
'
'   sRetDate = gMonthName(sDate)
'   Where:
'       sDate (I)- Date string for which month name is obtained
'       sRetDate (O)- Month name (not abbreviated)
'

    Dim ilMonth As Integer
    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slDate = gAdjYear(slAnyDate)
    On Error GoTo gMonthNameErr
    ilMonth = Month(slDate)
    Select Case ilMonth
        Case 1
            gMonthName = "January"
        Case 2
            gMonthName = "February"
        Case 3
            gMonthName = "March"
        Case 4
            gMonthName = "April"
        Case 5
            gMonthName = "May"
        Case 6
            gMonthName = "June"
        Case 7
            gMonthName = "July"
        Case 8
            gMonthName = "August"
        Case 9
            gMonthName = "September"
        Case 10
            gMonthName = "October"
        Case 11
            gMonthName = "November"
        Case 12
            gMonthName = "December"
    End Select
    Exit Function
gMonthNameErr:
    gMonthName = ""
    On Error GoTo 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gMonthYearFormat                *
'*                                                     *
'*             Created:11/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain month name, year from    *
'*                     Date                            *
'*                                                     *
'*******************************************************
Function gMonthYearFormat(slInpDate As String) As String
'
'   sRetDate = gMonthYearFormat(sDate)
'   Where:
'       sDate (I)- Date string for which month name and year are obtained
'       sRetDate (O)- Month name (3 characters), Year
'

    Dim ilMonth As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slDate = gAdjYear(slAnyDate)
    On Error GoTo gMonthYearFormatErr
    ilMonth = Month(slDate)
    Select Case ilMonth
        Case 1
            slStr = "Jan, "
        Case 2
            slStr = "Feb, "
        Case 3
            slStr = "Mar, "
        Case 4
            slStr = "Apr, "
        Case 5
            slStr = "May, "
        Case 6
            slStr = "June, "
        Case 7
            slStr = "July, "
        Case 8
            slStr = "Aug, "
        Case 9
            slStr = "Sept, "
        Case 10
            slStr = "Oct, "
        Case 11
            slStr = "Nov, "
        Case 12
            slStr = "Dec, "
    End Select
    gMonthYearFormat = slStr & Trim$(str$(Year(slDate)))
    Exit Function
gMonthYearFormatErr:
    gMonthYearFormat = ""
    On Error GoTo 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gNow                            *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain the system Date and Time*
'*                                                     *
'*******************************************************
Function gNow() As String
    Dim slDate As String
    Dim slDateTime As String
    Dim ilPos As Integer

    'Dan M 9/20 reports.exe issue, change storage of system date
    'slDate = Trim$(gGetCSIName("SYSDate"))
    slDate = gCSIGetName()
    If slDate = "" Then
        gNow = Now
    Else
        slDateTime = Trim$(Now)
        ilPos = InStr(1, slDateTime, " ", 1)
        If ilPos > 0 Then
            gNow = slDate & Mid$(slDateTime, ilPos)
        Else
            gNow = slDate
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainEndCal                   *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain calendar month end date *
'*                      of specified date              *
'*                                                     *
'*******************************************************
Function gObtainEndCal(slInpDate As String) As String
'
'   sRetDate = gObtainEndCal(sDate)
'   Where:
'       sDate (I)- Date for which the calendar month end date is to be obtained
'       sRetDate (O)- End date of the calendar month
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    Do While Month(llDate) = Month(llDate + 1)
        llDate = llDate + 1
    Loop
    gObtainEndCal = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainEndCorp                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain corporate month end date*
'*                      of specified date              *
'*                                                     *
'*******************************************************
Function gObtainEndCorp(slInpDate As String, ilTestStart As Integer) As String
'
'   sRetDate = gObtainEndCorp(sDate, ilTest)
'   Where:
'       sDate (I)- Date for which the corporate month end date is to be obtained
'       ilTest (I) - True = Test date with start; False = Don't test
'       sRetDate (O)- End date of the corporate month
'
    Dim llDate As Long
    Dim llTempDate As Long
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    'Dim llLastDate As Long
    'Dim slLastDate As String
    Dim slAnyDate As String
    'Dim ilDone As Integer

    slAnyDate = slInpDate
    'gObtainYearMonthDayStr slAnyDate, False, slYear, slMonth, slDay
    'ilYear = Val(slYear)
    llDate = gDateValue(slAnyDate)
    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
    '    If (tgMCof(ilLoop).iYear = ilYear) Or (tgMCof(ilLoop).iYear = ilYear + 1) Then
            'For ilIndex = 1 To 12 Step 1
            For ilIndex = 0 To 11 Step 1
                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llTempDate
                If llDate >= llTempDate Then
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llTempDate
                    If llDate <= llTempDate Then
                        gObtainEndCorp = Format$(llTempDate, "m/d/yy")
                        Exit Function
                    End If
                End If
            Next ilIndex
    '    End If
    Next ilLoop
    gObtainEndCorp = ""
    'ilDone = False
    'Do
    '    gObtainYearMonthDayStr slAnyDate, False, slYear, slMonth, slDay
    '    ilYear = Val(slYear)
    '    ilMonth = Val(slMonth)
    '    'Obtain the first date of the year
    '    slDate = "1/1/" & slYear
    '    llDate = gDateValue(slDate)
    '    If tgSpf.sRYEnd = "L" Then  'End corporate year on last sunday of the year
    '        Do While gWeekDayLong(llDate) <> 0   '0=monday
    '            llDate = llDate - 1
    '        Loop
    '    Else    'End year on 12/31/....
    '        'Count the number of weeks from the first monday
    '        llTempDate = llDate
    '        Do While gWeekDayLong(llTempDate) <> 0   '0=monday
    '            llTempDate = llTempDate + 1
    '        Loop
    '    End If
    '    ilIndex = 0
    '    For ilLoop = 2 To ilMonth Step 1
    '        If (ilLoop = 2) And (tgSpf.sRYEnd <> "L") Then
    '            llDate = llTempDate + 7 * tgSpf.iRCorp(ilIndex)
    '        Else
    '            llDate = llDate + 7 * tgSpf.iRCorp(ilIndex)
    '        End If
    '        ilIndex = ilIndex + 1
    '        If ilIndex > 2 Then
    '            ilIndex = 0
    '        End If
    '    Next ilLoop
    '    If ilMonth <> 12 Then
    '        If (ilMonth = 1) And (tgSpf.sRYEnd <> "L") Then
    '            llLastDate = llTempDate + 7 * tgSpf.iRCorp(ilIndex) - 1
    '        Else
    '            llLastDate = llDate + 7 * tgSpf.iRCorp(ilIndex) - 1
    '        End If
    '    Else
    '        slLastDate = "1/1/" & Trim$(Str$(ilYear + 1))
    '        llLastDate = DateValue(slLastDate)
    '        If tgSpf.sRYEnd = "L" Then
    '            Do While gWeekDayLong(llLastDate) <> 0   '0=monday
    '                llLastDate = llLastDate - 1
    '            Loop
    '        End If
    '        llLastDate = llLastDate - 1
    '    End If
    '    If ilTestStart Then
    '        slDate = slMonth & "/15/" & slYear
    '        slDate = gObtainStartCorp(slDate, False) 'Obtain corporate end to test if date is in next month
    '        If gDateValue(slAnyDate) > llLastDate Then
    '            slAnyDate = Format$(llLastDate + 10, "m/d/yy")    'Start of next month
    '        Else
    '            If gDateValue(slAnyDate) < gDateValue(slDate) Then  'Prior to start
    '                gObtainEndCorp = Format$(gDateValue(slDate) - 1, "m/d/yy")
    '                Exit Function
    '            Else
    '                gObtainEndCorp = Format$(llLastDate, "m/d/yy")
    '                Exit Function
    '            End If
    '        End If
    '    Else
    '        gObtainEndCorp = Format$(llLastDate, "m/d/yy")
    '        Exit Function
    '    End If
    'Loop While Not ilDone
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainEndStd                   *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain standard month end date *
'*                      of specified date              *
'*                                                     *
'*******************************************************
Function gObtainEndStd(slInpDate As String) As String
'
'   sRetDate = gObtainEndStd(sDate)
'   Where:
'       sDate (I)- Date for which the standard month end date is to be obtained
'       sRetDate (O)- End date of the standard month
'

    Dim llDate As Long
    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slDate = gObtainNextSunday(slAnyDate)
    llDate = gDateValue(slDate)
    Do While Month(llDate) = Month(llDate + 7)
        llDate = llDate + 7
    Loop
    gObtainEndStd = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMondayFromToday          *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the next monday date as  *
'*                     a string from today's date      *
'*                                                     *
'*******************************************************
Function gObtainMondayFromToday() As String
'
'   sRetDate = gObtainMondayFromToday()
'   Where:
'       sRetDate (O)- First Monday after today's date
'

    Dim slDate As String
    Dim llDate As Long
    Dim ilMatchDay As Integer
    slDate = Format$(gNow(), "m/d/yy") 'Date$
    llDate = gDateValue(gAdjYear(slDate)) + 1  'Not include todays date
    ilMatchDay = 0
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate + 1
    Loop
    gObtainMondayFromToday = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMonthYear                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Month and year given date*
'*                                                     *
'*******************************************************
Sub gObtainMonthYear(ilType As Integer, slInpDate As String, ilMonth As Integer, ilYear As Integer)
'
'   gObtainMonthYear ilType, sDate, ilMonth, ilYear
'   Where:
'       ilType (I)- 0=Standard month; 1= Regular month; 4=Corp on Jan-Dec; 5=Corp on Fiscal (Oct-Sept)
'       sDate (I)- Date to obtain Month and year
'       ilMonth (O) - Month
'       ilYear (O) - Year
'

    Dim llDate As Long
    Dim slDate As String
    Dim slAnyDate As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llSDate As Long
    Dim llEDate As Long

    slAnyDate = Trim$(slInpDate)
    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then  'Std
        slDate = gObtainEndStd(slAnyDate)
        slDate = gAdjYear(slDate)
    ElseIf ilType = 4 Then  'Corp on Jan-Dec year
        'slStartDate = gObtainStartCorp(slAnyDate, True)
        'slEndDate = gObtainEndCorp(slAnyDate, True)
        'slDate = Format$(gDateValue(slStartDate) + (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 2, "m/d/yy")
        'slDate = gAdjYear(slDate)
        llDate = gDateValue(slAnyDate)
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            For ilIndex = 1 To 12 Step 1
                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llSDate
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llEDate
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    ilYear = tgMCof(ilLoop).iYear
                    ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
                    If ilMonth > 12 Then
                        ilMonth = ilMonth - 12
                    Else
                        ilYear = ilYear - 1
                    End If
                    Exit Sub
                End If
            Next ilIndex
        Next ilLoop
        ilMonth = 0
        ilYear = 0
        Exit Sub
    ElseIf ilType = 5 Then  'Corp on Fiscal (Oct-Sept)
        'slStartDate = gObtainStartCorp(slAnyDate, True)
        'slEndDate = gObtainEndCorp(slAnyDate, True)
        'slDate = Format$(gDateValue(slStartDate) + (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 2, "m/d/yy")
        'slDate = gAdjYear(slDate)
        llDate = gDateValue(slAnyDate)
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            For ilIndex = 1 To 12 Step 1
                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex - 1), tgMCof(ilLoop).iStartDate(1, ilIndex - 1), llSDate
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex - 1), tgMCof(ilLoop).iEndDate(1, ilIndex - 1), llEDate
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    ilYear = tgMCof(ilLoop).iYear
                    ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
                    If ilMonth > 12 Then
                        ilMonth = ilMonth - 12
                    End If
                    Exit Sub
                End If
            Next ilIndex
        Next ilLoop
        ilMonth = 0
        ilYear = 0
        Exit Sub
    Else
        slDate = gAdjYear(slAnyDate)
    End If
    ilMonth = Month(slDate)
    ilYear = Year(slDate)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainNextMonday               *
'*                                                     *
'*             Created:8/31/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain first monday from        *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainNextMonday(slInpDate As String) As String
'
'   sRetDate = gObtainNextMonday(sDate)
'   Where:
'       sDate (I)- Date for which next monday is obtained
'       sRetDate (O)- Next monday including specified date
'

    Dim llDate As Long
    Dim ilMatchDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    ilMatchDay = 0
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate + 1
    Loop
    gObtainNextMonday = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainNextSunday               *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain first sunday from        *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainNextSunday(slInpDate As String) As String
'
'   sRetDate = gObtainNextSunday(sDate)
'   Where:
'       sDate (I)- Date to obtain next sunday from
'       sRetDate (O)- Next sunday including specified date
'

    Dim llDate As Long
    Dim ilMatchDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    ilMatchDay = 6  '6=sunday
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate + 1
    Loop
    gObtainNextSunday = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPrevtMonday               *
'*                                                     *
'*             Created:8/31/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain previous monday from     *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainPrevMonday(slInpDate As String) As String
'
'   sRetDate = gObtainPrevMonday(sDate)
'   Where:
'       sDate (I)- Date for which previous monday is obtained
'       sRetDate (O)- Previous monday including specified date
'

    Dim llDate As Long
    Dim ilMatchDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    ilMatchDay = 0
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate - 1
    Loop
    gObtainPrevMonday = Format$(llDate, "m/d/yy")
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPrevtMondayLong          *
'*                                                     *
'*             Created:8/31/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain previous monday from     *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainPrevMondayLong(llInpDate As Long) As Long
'
'   sRetDate = gObtainPrevMonday(sDate)
'   Where:
'       lDate (I)- Date for which previous monday is obtained
'       lRetDate (O)- Previous monday including specified date
'

    Dim llDate As Long
    Dim ilMatchDay As Integer

    llDate = llInpDate
    ilMatchDay = 0
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate - 1
    Loop
    gObtainPrevMondayLong = llDate
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPrevSunday               *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain previous sunday from     *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainPrevSunday(slInpDate As String) As String
'
'   sRetDate = gObtainPrevSunday(sDate)
'   Where:
'       sDate (I)- Date to obtain previous sunday from
'       sRetDate (O)- Previous sunday including specified date
'

    Dim llDate As Long
    Dim ilMatchDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    ilMatchDay = 6  '6=sunday
    Do Until gWeekDayLong(llDate) = ilMatchDay
        llDate = llDate - 1
    Loop
    gObtainPrevSunday = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStartCal                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain calendar month start     *
'*                     date of specified date          *
'*                                                     *
'*******************************************************
Function gObtainStartCal(slInpDate As String) As String
'
'   sRetDate = gObtainStartCal(sDate)
'   Where:
'       sDate (I)- Date for which the calendar month start date is to be obtained
'       sRetDate (O)- Start date of the calendar month
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    llDate = gDateValue(slAnyDate)
    Do While Day(llDate) <> 1
        llDate = llDate - 1
    Loop
    gObtainStartCal = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStartCorp                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain corporate month start   *
'*                      date of specified date         *
'*                                                     *
'*******************************************************
Function gObtainStartCorp(slInpDate As String, ilTestEnd As Integer) As String
'
'   sRetDate = gObtainStartCorp(sDate, ilTest)
'   Where:
'       sDate (I)- Date for which the corporate month start date is to be obtained
'       ilTest (I) - True = Test date with end; False = Don't test
'       sRetDate (O)- Start date of the corporate month
'
    Dim llDate As Long
    Dim llTempDate As Long
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slAnyDate As String
    'Dim ilDone As Integer

    'ilDone = False
    slAnyDate = slInpDate
    'gObtainYearMonthDayStr slAnyDate, False, slYear, slMonth, slDay
    'ilYear = Val(slYear)
    llDate = gDateValue(slAnyDate)
    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
    '    If (tgMCof(ilLoop).iYear = ilYear) Or (tgMCof(ilLoop).iYear = ilYear + 1) Then
            'For ilIndex = 1 To 12 Step 1
            For ilIndex = 0 To 11 Step 1
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llTempDate
                If llDate <= llTempDate Then
                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llTempDate
                    If llDate >= llTempDate Then
                        gObtainStartCorp = Format$(llTempDate, "m/d/yy")
                        Exit Function
                    End If
                End If
            Next ilIndex
    '    End If
    Next ilLoop
    gObtainStartCorp = ""
    'Do
    '    gObtainYearMonthDayStr slAnyDate, False, slYear, slMonth, slDay
    '    ilYear = Val(slYear)
    '    ilMonth = Val(slMonth)
    '    'Obtain the first date of the year
    '    slDate = "1/1/" & slYear
    '    llDate = gDateValue(slDate)
    '    If tgSpf.sRYEnd = "L" Then  'End corporate year on last sunday of the year
    '        Do While gWeekDayLong(llDate) <> 0   '0=monday
    '            llDate = llDate - 1
    '        Loop
    '    Else    'End year on 12/31/....
    '        'Count the number of weeks from the first monday
    '        llTempDate = llDate
    '        Do While gWeekDayLong(llTempDate) <> 0   '0=monday
    '            llTempDate = llTempDate + 1
    '        Loop
    '    End If
    '    ilIndex = 0
    '    For ilLoop = 2 To ilMonth Step 1
    '        If (ilLoop = 2) And (tgSpf.sRYEnd <> "L") Then
    '            llDate = llTempDate + 7 * tgSpf.iRCorp(ilIndex)
    '        Else
    '            llDate = llDate + 7 * tgSpf.iRCorp(ilIndex)
    '        End If
    '        ilIndex = ilIndex + 1
    '        If ilIndex > 2 Then
    '            ilIndex = 0
    '        End If
    '    Next ilLoop
    '    If ilTestEnd Then
    '        slDate = slMonth & "/15/" & slYear
    '        slDate = gObtainEndCorp(slDate, False) 'Obtain corporate end to test if date is in next month
    '        If gDateValue(slAnyDate) > gDateValue(slDate) Then  'Next month
    '            gObtainStartCorp = Format$(gDateValue(slDate) + 1, "m/d/yy")
    '            Exit Function
    '        Else
    '            If gDateValue(slAnyDate) < llDate Then
    '                'Compute start of previous month
    '                slAnyDate = Format$(llDate - 10, "m/d/yy")    'End date of previous month
    '            Else
    '                gObtainStartCorp = Format$(llDate, "m/d/yy")
    '                Exit Function
    '            End If
    '        End If
    '    Else
    '        gObtainStartCorp = Format$(llDate, "m/d/yy")
    '        Exit Function
    '    End If
    'Loop While Not ilDone
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStartDateForWkNo         *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Start Date               *
'*                                                     *
'*******************************************************
Function gObtainStartDateForWkNo(ilType As Integer, ilWkNo As Integer, ilYear As Integer) As String
'
'   slStartDate = gObtainStartDateForWkNo (ilType, ilWkNo, ilYear)
'   Where:
'       ilType (I)- 0=Standard month; 1= Regular month; 4=Corp on Jan-Dec but from standard start; 5=Corp on Fiscal (Oct-Sept)
'       ilWkNo(I)- Week number from 1 to 53 (-1 means last week)
'       ilYear(I)- Year
'       slStartDate (I)- Date for week number specified
'
    Dim slDate As String

    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then
        If ilWkNo > 0 Then
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearStartDate(ilType, slDate)
            gObtainStartDateForWkNo = Format$(gDateValue(slDate) + 7 * (ilWkNo - 1), "m/d/yy")
        Else
            slDate = "12/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearEndDate(ilType, slDate)
            gObtainStartDateForWkNo = gObtainPrevMonday(slDate)
        End If
    ElseIf ilType = 1 Then
        If ilWkNo > 0 Then
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearStartDate(ilType, slDate)
            gObtainStartDateForWkNo = Format$(gDateValue(slDate) + 7 * (ilWkNo - 1), "m/d/yy")
        Else
            slDate = "12/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearEndDate(ilType, slDate)
            gObtainStartDateForWkNo = gObtainPrevMonday(slDate)
        End If
    ElseIf ilType = 4 Then  'corp cal for Jan-dec year
        If ilWkNo > 0 Then
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearStartDate(ilType, slDate)
            gObtainStartDateForWkNo = Format$(gDateValue(slDate) + 7 * (ilWkNo - 1), "m/d/yy")
        Else
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearEndDate(ilType, slDate)
            gObtainStartDateForWkNo = gObtainPrevMonday(slDate)
        End If
    ElseIf ilType = 5 Then  'Corp cal on Fiscal (oct-sept)
        If ilWkNo > 0 Then
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearStartDate(ilType, slDate)
            gObtainStartDateForWkNo = Format$(gDateValue(slDate) + 7 * (ilWkNo - 1), "m/d/yy")
        Else
            slDate = "1/15/" & Trim$(str$(ilYear))
            slDate = gObtainYearEndDate(ilType, slDate)
            gObtainStartDateForWkNo = gObtainPrevMonday(slDate)
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStartStd                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain standard month start     *
'*                     date of specified date          *
'*                                                     *
'*******************************************************
Function gObtainStartStd(slInpDate As String) As String
'
'   sRetDate = gObtainStartStd(sDate)
'   Where:
'       sDate (I)- Date for which the standard month start date is to be obtained
'       sRetDate (O)- Start date of the standard month
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slAnyDate = gObtainEndStd(slAnyDate)
    llDate = gDateValue(slAnyDate)
    slAnyDate = gObtainEndStd(Format$(llDate - 40, "m/d/yy"))
    llDate = gDateValue(slAnyDate) + 1
    gObtainStartStd = Format$(llDate, "m/d/yy")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainWkNo                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Week Number              *
'*                                                     *
'*******************************************************
Sub gObtainWkNo(ilType As Integer, slInpDate As String, ilWkNo As Integer, ilFirstLastWk As Integer)
'
'   gObtainWkNo ilType, slDate, ilWkNo, ilFirstLastWk
'   Where:
'       ilType (I)- 0=Standard month; 1= Regular month; 4=Corp on Jan-Dec but from standard start; 5=Corp on Fiscal (Oct-Sept)
'       slDate (I)- Date to obtain week number
'       ilWkNo(O)- Week number from 1 to 53
'       ilFirstLastWk(O) If corp, indicates if this is the last week (week 52 and 53 should be added)
'                        or if this is the first week (week 0 and 1 should be added)
'
    '1996
    'Dec 29 Dec 30 Dec 31 Jan 1 Jan 2 Jan 3 Jan 4 Jan 5 Jan 6
    ' Sun    Mon    Tue    Wed   Thu   Fri   Sat   Sun   Mon
    'Std Week Index or Corp Last Sunday
    '         1      1      1     1     1     1     1     2
    'Corp Week Index- Jan 1 Start (index zero returned as 1, since if one return, week zero
    '                              should be included)
    '         0      0      1     1     1     1     1     2
    ' return  1      1      1     1     1     1     1     2
    '
    Dim slDate As String


    If ilType <> 4 Then
        slDate = gObtainYearStartDate(ilType, slInpDate)
    Else
        slDate = gObtainYearStartDate(0, slInpDate)
    End If
    'If ilType = 4 Then
    '    slDate = gObtainPrevMonday(slDate)
    'End If
    ilFirstLastWk = False
    ilWkNo = (gDateValue(slInpDate) - gDateValue(slDate)) \ 7 + 1
    'If ((ilType = 4) And (tgSpf.sRYEnd = "L") And (ilWkNo = 1)) Or ((ilType = 1) And (ilWkNo = 1)) Then
    If ((ilType = 4) And (ilWkNo = 1)) Or ((ilType = 5) And (ilWkNo = 1)) Or ((ilType = 1) And (ilWkNo = 1)) Then
        ilFirstLastWk = True
    End If
    'If (ilType = 4) And (tgSpf.sRYEnd <> "L") And (ilWkNo = 52) Then
    '    slDate = gObtainEndCorp(slInpDate, True)
    '    If gWeekDayStr(slDate) <> 6 Then
    '        ilFirstLastWk = True
    '    End If
    'End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainYearEndDate             *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain End Date of the year     *
'*                                                     *
'*******************************************************
Function gObtainYearEndDate(ilType As Integer, slInpDate As String) As String
'
'   slEndDate = gObtainYearEndDate(ilType, slInpDate)
'   Where:
'       ilType (I)- 0=Standard month;1= Regular month; 4=Corp on Jan-Dec; 5=Corp on Fiscal (Oct-Sept)
'       slInpDate (I)- Date to obtain end date
'       slEndDate(O) - Start Date of The Year
'
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llSDate As Long
    Dim llEDate As Long

    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then
        slDate = gObtainEndStd(slInpDate)
    ElseIf ilType = 1 Then
        slDate = slInpDate
    'This code does not work- changing to use same code as ilType=5.  That code will always
    'work as long as the corporate year matches the Jan Year.
    '  Oct      Dec  Jan                    Sept
    '  2002     2002 2003                   2003
    ' |-------------|---------------------------|
    '
    '  Feb                             Dec  Jan
    '  2002                            2002 2003
    ' |------------------------------------|----|
    '  If this is defined as 2003 year it will work.
    '  If this is defined as 2002 year, then it will fail
    '  (1/15/2002 year will not get this corporate calendar)
    '
    'ElseIf ilType = 4 Then  'corp cal for Jan-dec year
    '    slDate = slInpDate
    '    llDate = gDateValue(slDate)
    '    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
    ''        For ilIndex = 1 To 12 Step 1
    '        For ilIndex = 0 To 11 Step 1
    '            gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
    '            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
    '            If (llDate >= llSDate) And (llDate <= llEDate) Then
    '                'Find Jan
    '                ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
'    '                If ilMonth > 12 Then
'    '                    ilMonth = ilMonth - 12
    '                If ilMonth > 11 Then
    '                    ilMonth = ilMonth - 11

    '                End If
    '                If ilMonth >= tgMCof(ilLoop).iStartMnthNo Then
    '                    slDate = "12/15/" & Trim$(Str$(tgMCof(ilLoop).iYear - 1))
    '                Else
    '                    slDate = "12/15/" & Trim$(Str$(tgMCof(ilLoop).iYear))
    '                End If
    '                llDate = gDateValue(slDate)
    ''                For ilMonth = 1 To 12 Step 1
    '                For ilMonth = 0 To 11 Step 1
    '                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilMonth), tgMCof(ilLoop).iStartDate(1, ilMonth), llSDate
    '                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilMonth), tgMCof(ilLoop).iEndDate(1, ilMonth), llEDate
    '                    If (llDate >= llSDate) And (llDate <= llEDate) Then
    '                        gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slDate
    '                        gObtainYearEndDate = slDate
    '                        Exit Function
    '                    End If
    '                Next ilMonth
    '            End If
    '        Next ilIndex
    '    Next ilLoop
    '    gObtainYearEndDate = ""
    '    Exit Function
    ElseIf (ilType = 5) Or (ilType = 4) Then  'Corp cal on Fiscal (oct-sept)
        slDate = slInpDate
        llDate = gDateValue(slDate)
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            'For ilIndex = 1 To 12 Step 1
            For ilIndex = 0 To 11 Step 1
                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    'gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slDate
                    gUnpackDate tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), slDate
                    gObtainYearEndDate = slDate
                    Exit Function
                End If
            Next ilIndex
        Next ilLoop
        gObtainYearEndDate = ""
        Exit Function
    End If
    gObtainMonthYear ilType, slDate, ilMonth, ilYear  '0=Standard month
    slDate = "12/15/" & Trim$(str$(ilYear))
    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then
        gObtainYearEndDate = gObtainEndStd(slDate)
    ElseIf ilType = 1 Then
        gObtainYearEndDate = "12/31/" & Trim$(str$(ilYear))
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainYearMonthDay             *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain year, month and day      *
'*                     from specified date and front   *
'*                     fill with zero if specified     *
'*                                                     *
'*******************************************************
Sub gObtainYearMonthDayStr(slInpDate As String, ilZeroFill As Integer, slYear As String, slMonth As String, slDay As String)
'
'   gObtainYearMonthDayStr(sDate, ilFill, slYear, slMonth, slDay)
'   Where:
'       sDate (I)- Date
'       ilFill(I)- True=fill year to 4 digits; month to two digits; day to two digits
'       slYear(O)- Year adjusted (1970 thru 2069)
'       slMonth(O)- Month #
'       slDay(O)- Day number
'
    Dim ilRet As Integer
    Dim ilYear As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slYear = ""
    slMonth = ""
    slDay = ""
    On Error GoTo gObtainYearMonthDayStrErr
    ilRet = gParseItem(slAnyDate, 1, "/", slMonth)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    ilRet = gParseItem(slAnyDate, 2, "/", slDay)
    If ilRet <> CP_MSG_NONE Then
        Exit Sub
    End If
    ilRet = gParseItem(slAnyDate, 3, "/", slYear)
    If ilRet <> CP_MSG_NONE Then
        slYear = str$(Year(gNow()))
    End If
    ilYear = Val(slYear)
    If (ilYear >= 0) And (ilYear <= 69) Then
        ilYear = 2000 + ilYear
    ElseIf (ilYear >= 70) And (ilYear <= 99) Then
        ilYear = 1900 + ilYear
    End If
    slYear = Trim$(str$(ilYear))
    If ilZeroFill Then
        Do While Len(slYear) < 4
            slYear = "0" & slYear
        Loop
        Do While Len(slMonth) < 2
            slMonth = "0" & slMonth
        Loop
        Do While Len(slDay) < 2
            slDay = "0" & slDay
        Loop
    End If
    Exit Sub
gObtainYearMonthDayStrErr:
    Exit Sub

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainYearStartDate            *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Start Date of the year   *
'*                                                     *
'*******************************************************
Function gObtainYearStartDate(ilType As Integer, slInpDate As String) As String
'
'   slStartDate = gObtainYearStartDate(ilType, slInpDate)
'   Where:
'       ilType (I)- 0=Standard month;1= Regular month; 4=Corp on Jan-Dec; 5=Corp on Fiscal (Oct-Sept)
'       slDate (I)- Date to obtain start date
'       slStartDate(O) - Start Date of The Year
'
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llSDate As Long
    Dim llEDate As Long

    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then
        slDate = gObtainEndStd(slInpDate)
    ElseIf ilType = 1 Then
        slDate = slInpDate
    'This code does not work- changing to use same code as ilType=5.  That code will always
    'work as long as the corporate year matches the Jan Year.  9/28/02
    '  Oct      Dec  Jan                    Sept
    '  2002     2002 2003                   2003
    ' |-------------|---------------------------|
    '
    '  Feb                             Dec  Jan
    '  2002                            2002 2003
    ' |------------------------------------|----|
    '  If this is defined as 2003 year it will work.
    '  If this is defined as 2002 year, then it will fail
    '  (1/15/2002 year will not get this corporate calendar)
    'ElseIf ilType = 4 Then  'corp cal for Jan-dec year
    '    slDate = slInpDate
    '    llDate = gDateValue(slDate)
    '    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
    '        For ilIndex = 1 To 12 Step 1
    '            gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
    '            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
    '            If (llDate >= llSDate) And (llDate <= llEDate) Then
    '                'Find Jan
    '                ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
    '                If ilMonth > 12 Then
    '                    ilMonth = ilMonth - 12
    '                End If
    '                If ilMonth >= tgMCof(ilLoop).iStartMnthNo Then
    '                    slDate = "1/15/" & Trim$(Str$(tgMCof(ilLoop).iYear - 1))
    '                Else
    '                    slDate = "1/15/" & Trim$(Str$(tgMCof(ilLoop).iYear))
    '                End If
    '                llDate = gDateValue(slDate)
    '                For ilMonth = 1 To 12 Step 1
    '                    gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilMonth), tgMCof(ilLoop).iStartDate(1, ilMonth), llSDate
    '                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilMonth), tgMCof(ilLoop).iEndDate(1, ilMonth), llEDate
    '                    If (llDate >= llSDate) And (llDate <= llEDate) Then
    '                        gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slDate
    '                        gObtainYearStartDate = slDate
    '                        Exit Function
    '                    End If
    '                Next ilMonth
    '            End If
    '        Next ilIndex
    '    Next ilLoop
    '    gObtainYearStartDate = ""
    '    Exit Function
    ElseIf (ilType = 5) Or (ilType = 4) Then  'Corp cal on Fiscal (oct-sept)
        slDate = slInpDate
        llDate = gDateValue(slDate)
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            'For ilIndex = 1 To 12 Step 1
            For ilIndex = 0 To 11 Step 1
                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    gUnpackDate tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), slDate
                    gObtainYearStartDate = slDate
                    Exit Function
                End If
            Next ilIndex
        Next ilLoop
        gObtainYearStartDate = ""
        Exit Function
    End If
    gObtainMonthYear ilType, slDate, ilMonth, ilYear  '0=Standard month
    slDate = "1/15/" & Trim$(str$(ilYear))
    If (ilType = 0) Or (((ilType = 4) Or (ilType = 5)) And (tgSpf.sRUseCorpCal <> "Y")) Then
        gObtainYearStartDate = gObtainStartStd(slDate)
    ElseIf ilType = 1 Then
        gObtainYearStartDate = "1/1/" & Trim$(str$(ilYear))
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPackDate                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Pack date in btrieve format     *
'*                                                     *
'*******************************************************
Sub gPackDate(slInpDate As String, ilDyMn As Integer, ilYear As Integer)
'
'   gPackDate slDate, ilDyMn, ilYear
'   Where:
'       slDate (I) - Date to be packed in btrieve format
'       ilDyMn (O)- High order byte = Day; low order byte = month
'       ilYear (O)- Year
'

    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    If Len(slAnyDate) = 0 Then
        ilDyMn = 0
        ilYear = 0
    Else
        slDate = gAdjYear(slAnyDate)
        ilDyMn = Day(slDate) + Month(slDate) * 256 'byte 1= day, byte 2=month
        ilYear = Year(slDate)
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPackDateLong                   *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Pack date in btrieve format     *
'*                                                     *
'*******************************************************
Sub gPackDateLong(llInpDate As Long, ilDyMn As Integer, ilYear As Integer)
'
'   gPackDateLong llDate, ilDyMn, ilYear
'   Where:
'       llDate (I) - Date to be packed in btrieve format
'       ilDyMn (O)- High order byte = Day; low order byte = month
'       ilYear (O)- Year
'

    Dim slDate As String

    If llInpDate <= 0 Then
        slDate = ""
    Else
        slDate = Format$(llInpDate, "m/d/yy")
    End If
    gPackDate slDate, ilDyMn, ilYear
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPackLength                      *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Pack length in btrieve format   *
'*                                                     *
'*******************************************************
Sub gPackLength(slInpTime As String, ilHsSec As Integer, ilMinHr As Integer)
'
'   gPackLength slTime, ilHsSec, ilMinHr
'   Where:
'       slTime (I) - Length of Time to be packed in btrieve format
'       ilHsSec (O)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (O)- High order byte = minute; low order byte = hours
'

    Dim ilPos As Integer
    Dim slLen As String
    Dim slHour As String
    Dim ilHour As Integer
    Dim slMin As String
    Dim ilMin As Integer
    Dim slSec As String
    Dim ilSec As Integer
    Dim ilFormat As Integer
    Dim slTime As String

    slTime = Trim$(slInpTime)
    If Len(slTime) = 0 Then
        ilHsSec = 1 'High order byte = hundredths of seconds; low order byte = seconds
        ilMinHr = 0 'High order byte = minute; low order byte = hours
        Exit Sub
    End If
    slHour = ""
    slMin = ""
    slSec = ""
    slLen = Trim$(slTime)
    ilPos = InStr(1, slLen, ":")
    If ilPos > 0 Then
        ilFormat = 1
    Else
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            ilFormat = 3
        Else
            ilPos = InStr(1, slLen, "m", 1)
            If ilPos > 0 Then
                ilFormat = 3
            Else
                ilPos = InStr(1, slLen, "s", 1)
                If ilPos > 0 Then
                    ilFormat = 3
                Else
                    ilFormat = 2
                End If
            End If
        End If
    End If
    If ilFormat = 2 Then 'hh mm'ss"
        ilPos = InStr(1, slLen, " ")
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "'")
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            End If
         Else
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                End If
            End If
         End If
    ElseIf ilFormat = 3 Then 'hhHmmMssS
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "m", 1)
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            End If
         Else
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                End If
            End If
         End If
    Else    'format hh:mm:ss
        ilPos = InStr(slLen, ":")
        If ilPos > 0 Then   'Might be hour/min/sec or hour/min only
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, ":")
            If ilPos > 0 Then
                slMin = Left$(slLen, ilPos - 1)
                slSec = Mid$(slLen, ilPos + 1)
            Else
                slMin = slLen
            End If
        Else
            slHour = slLen
        End If
    End If
    If slHour <> "" Then
        ilHour = Val(slHour)
    Else
        ilHour = 0
    End If
    If slMin <> "" Then
        ilMin = Val(slMin)
    Else
        ilMin = 0
    End If
    If slSec <> "" Then
        ilSec = Val(slSec)
    Else
        ilSec = 0
    End If
    If ilSec > 59 Then
        ilMin = ilMin + ilSec \ 60
        ilSec = ilSec Mod 60
    End If
    If ilMin > 59 Then
        ilHour = ilHour + ilMin \ 60
        ilMin = ilMin Mod 60
    End If
    ilHsSec = ilSec * 256 'High order byte = hundredths of seconds; low order byte = seconds
    ilMinHr = ilMin + ilHour * 256 'High order byte = minute; low order byte = hours
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPackTime                        *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Pack time in btrieve format     *
'*                                                     *
'*******************************************************
Sub gPackTime(slInpTime As String, ilHsSec As Integer, ilMinHr As Integer)
'
'   gPackTime slTime, ilHsSec, ilMinHr
'   Where:
'       slTime (I) - Time to be packed in btrieve format
'       ilHsSec (O)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (O)- High order byte = minute; low order byte = hours
'

    Dim slTime As String
    slTime = Trim$(slInpTime)
    If Len(slTime) = 0 Then
        ilHsSec = 1 'indicate blank with hundredths of seconds set to 1 only
        ilMinHr = 0
    Else
        slTime = gConvertTime(slTime)
        ilHsSec = Second(slTime) * 256 'High order byte = hundredths of seconds; low order byte = seconds in the record
        ilMinHr = Minute(slTime) + Hour(slTime) * 256 'High order byte = minute; low order byte = hours in the record
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackTimeLong                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack time in btrieve format   *
'*                                                     *
'*******************************************************
Sub gPackTimeLong(llTime As Long, ilHsSec As Integer, ilMinHr As Integer)
'
'   gPackTimeLong llTime, ilHsSec, ilMinHr
'   Where:
'       llTime (I) - Time as long
'       ilHsSec (O)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (O)- High order byte = minute; low order byte = hours
'

    Dim llSec As Long    'Seconds
    Dim llMin As Long    'Minutes
    Dim llHour As Long   'Hours

    If llTime = -1 Then
        ilHsSec = 1
        ilMinHr = 0
        Exit Sub
    End If
    If (llTime = 0) Or (llTime = 86400) Then
        ilHsSec = 0
        ilMinHr = 0
        Exit Sub
    End If
    'llTime = llHour * 3600 + llMin * 60 + llSec
    llHour = llTime \ 3600
    llMin = llTime Mod 3600
    llSec = llMin Mod 60
    llMin = llMin \ 60
    ilHsSec = llSec * 256
    ilMinHr = llMin + llHour * 256
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gTimeToCurrency                 *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert time to currency (for   *
'*                     precision-                      *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gTimeToCurrency(slInpTime As String, ilChk12M As Integer) As Currency
'
'   clRetTime = gTimeToCurrency(slTime, ilChk12M)
'   Where:
'       slTime (I)- Time as string to be converted to currency
'       ilChk12M(I)- True=If 12M (0) convert to 86400 (24*3600)- handle end time
'                    False=Leave 12m as (0)
'       clRetTime (O)- time as currency
'
    Dim slTime As String
    Dim clTime As Currency
    Dim ilPos As Integer
    Dim slAnyTime As String

    slAnyTime = Trim$(slInpTime)
    On Error GoTo gTimeToCurrencyErr
    ilPos = InStr(slTime, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gTimeToCurrency = 0
            Exit Function
        End If
        slTime = Mid$(slAnyTime, 2)
    Else
        slTime = slAnyTime
    End If
    slTime = gConvertTime(slTime)
    clTime = Hour(slTime) * 3600
    clTime = clTime + Minute(slTime) * 60
    clTime = clTime + Second(slTime)
    If (clTime = 0) And ilChk12M Then
        clTime = 86400
    End If
    If ilPos = 0 Then
        gTimeToCurrency = clTime
    Else
        gTimeToCurrency = -clTime
    End If
    Exit Function
gTimeToCurrencyErr:
    On Error GoTo 0
    gTimeToCurrency = 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gTimeToCurrency                 *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert time to currency (for   *
'*                     precision-                      *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gTimeToLong(slInpTime As String, ilChk12M As Integer) As Long
'
'   llRetTime = gTimeToLong(slTime, ilChk12M)
'   Where:
'       slTime (I)- Time as string to be converted to currency
'       ilChk12M(I)- True=If 12M (0) convert to 86400 (24*3600)- handle end time
'                    False=Leave 12m as (0)
'       llRetTime (O)- time as Long
'
    Dim slTime As String
    Dim llTime As Long
    Dim ilPos As Integer
    Dim slAnyTime As String

    slAnyTime = Trim$(slInpTime)
    On Error GoTo gTimeToLongErr
    ilPos = InStr(slTime, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gTimeToLong = 0
            Exit Function
        End If
        slTime = Mid$(slAnyTime, 2)
    Else
        slTime = slAnyTime
    End If
    slTime = gConvertTime(slTime)
    llTime = CLng(Hour(slTime)) * 3600
    llTime = llTime + Minute(slTime) * 60
    llTime = llTime + Second(slTime)
    If (llTime = 0) And ilChk12M Then
        llTime = 86400
    End If
    If ilPos = 0 Then
        gTimeToLong = CLng(llTime)
    Else
        gTimeToLong = -llTime
    End If
    Exit Function
gTimeToLongErr:
    On Error GoTo 0
    gTimeToLong = 0
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnformatTime                   *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove trialing sec and/or      *
'*                     minutes if zero from time       *
'*                                                     *
'*******************************************************
Function gUnformatTime(slInpTime As String) As String
'
'   sRetTime = gUnformatTime(sTime)
'   Where:
'       sTime (I) - Time from which zero sec and/or minutes will be removed
'       sRetTime (O)- time with zeros removed
'

    Dim ilPos As Integer
    Dim slUnTime As String
    Dim slTime As String

    slTime = Trim$(slInpTime)
    slUnTime = slTime
    ilPos = InStr(slTime, ":0:0")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00:00")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":0A")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":0P")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00A")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00P")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    gUnformatTime = slUnTime
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackDate                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack date in btrieve format   *
'*                                                     *
'*******************************************************
Sub gUnpackDate(ilDyMn As Integer, ilYear As Integer, slDate As String)
'
'   gUnpackDate ilDyMn, ilYear, slDate
'   Where:
'       ilDyMn (I)- High order byte = Day; low order byte = month
'       ilYear (I)- Year
'       slDate (O) - Date as a string (MM/DD/YY)
'

    Dim ilDy As Integer 'Day #
    Dim ilMn As Integer 'Month #
    Dim dlDateSerial As Double
    If (ilDyMn = 0) And (ilYear = 0) Then
        slDate = ""
        Exit Sub
    End If
    'TFN Log Calendar date
    If ((ilDyMn >= 1) And (ilDyMn <= 7)) And (ilYear = 0) Then
        slDate = ""
        Exit Sub
    End If

    ilDy = ilDyMn And &HFF 'Obtain day #
    ilMn = ilDyMn \ 256  'Obtain month
    If ((ilDy < 1) Or (ilDy > 31)) Or ((ilMn < 1) Or (ilMn > 12)) Or ((ilYear < 1900) Or (ilYear > 2100)) Then
        slDate = ""
        Exit Sub
    End If
    dlDateSerial = DateSerial(ilYear, ilMn, ilDy)
    slDate = Trim$(Format$(dlDateSerial, "m/d/yy"))  '"m/d/yy"))
'    slDate = Trim$(Str$(ilMn)) & "/" & Trim$(Str$(ilDy)) & "/" & Trim$(Str$(ilYear))
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackDateForSor               *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack date in btrieve format   *
'*                     as a sort string                *
'*                                                     *
'*******************************************************
Sub gUnpackDateForSort(ilDyMn As Integer, ilYear As Integer, slDate As String)
'
'   gUnpackDateForSort ilDyMn, ilYear, slDate
'   Where:
'       ilDyMn (I)- High order byte = Day; low order byte = month
'       ilYear (I)- Year
'       slDate (O) - Date as a string (MM/DD/YYYY)
'

    Dim ilDy As Integer 'Day #
    Dim ilMn As Integer 'Month #
    Dim dlDateSerial As Double
    Dim ilAdjYear As Integer
    Dim slSortDate As String

    If (ilDyMn = 0) And (ilYear = 0) Then
        slDate = ""
        Exit Sub
    End If
    ilDy = ilDyMn And &HFF 'Obtain day #
    ilMn = ilDyMn \ 256  'Obtain month
    ilAdjYear = ilYear
    If (ilAdjYear >= 0) And (ilAdjYear <= 69) Then
        ilAdjYear = 2000 + ilAdjYear
    ElseIf (ilAdjYear >= 70) And (ilAdjYear <= 99) Then
        ilAdjYear = 1900 + ilAdjYear
    End If
    dlDateSerial = DateSerial(ilAdjYear, ilMn, ilDy)
    slSortDate = Format$(dlDateSerial, "mm/dd/yyyy")
    slSortDate = Trim$(str$(gDateValue(slSortDate)))
    Do While Len(slSortDate) < 5
        slSortDate = "0" & slSortDate
    Loop
    slDate = slSortDate
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackDateLong                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack date in btrieve format   *
'*                                                     *
'*******************************************************
Sub gUnpackDateLong(ilDyMn As Integer, ilYear As Integer, llDate As Long)
'
'   gUnpackDate ilDyMn, ilYear, slDate
'   Where:
'       ilDyMn (I)- High order byte = Day; low order byte = month
'       ilYear (I)- Year
'       llDate (O) - Date as a long
'

    Dim ilDy As Integer 'Day #
    Dim ilMn As Integer 'Month #
    Dim dlDateSerial As Double
    Dim slDate As String

    If (ilDyMn = 0) And (ilYear = 0) Then
        llDate = 0
        Exit Sub
    End If
    'TFN Log Calendar date
    If ((ilDyMn >= 1) And (ilDyMn <= 7)) And (ilYear = 0) Then
        llDate = 0
        Exit Sub
    End If
    ilDy = ilDyMn And &HFF 'Obtain day #
    ilMn = ilDyMn \ 256  'Obtain month
    dlDateSerial = DateSerial(ilYear, ilMn, ilDy)
    slDate = Trim$(Format$(dlDateSerial, "m/d/yy"))
    llDate = gDateValue(slDate)
'    slDate = Trim$(Str$(ilMn)) & "/" & Trim$(Str$(ilDy)) & "/" & Trim$(Str$(ilYear))
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackDateLong                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack date in btrieve format   *
'*                                                     *
'*******************************************************
Sub gUnpackDateLongError(ilDyMn As Integer, ilYear As Integer, llDate As Long, slErrorMsg As String)
'
'   gUnpackDate ilDyMn, ilYear, slDate
'   Where:
'       ilDyMn (I)- High order byte = Day; low order byte = month
'       ilYear (I)- Year
'       llDate (O) - Date as a long
'

    Dim ilDy As Integer 'Day #
    Dim ilMn As Integer 'Month #
    Dim dlDateSerial As Double
    Dim slDate As String
    On Error GoTo gUnpackDateLongErrorErr:
    igUnpackDateError = False
    If (ilDyMn = 0) And (ilYear = 0) Then
        llDate = 0
        Exit Sub
    End If
    'TFN Log Calendar date
    If ((ilDyMn >= 1) And (ilDyMn <= 7)) And (ilYear = 0) Then
        llDate = 0
        Exit Sub
    End If
    ilDy = ilDyMn And &HFF 'Obtain day #
    ilMn = ilDyMn \ 256  'Obtain month
    dlDateSerial = DateSerial(ilYear, ilMn, ilDy)
    slDate = Trim$(Format$(dlDateSerial, "m/d/yy"))
    llDate = gDateValue(slDate)
'    slDate = Trim$(Str$(ilMn)) & "/" & Trim$(Str$(ilDy)) & "/" & Trim$(Str$(ilYear))
    Exit Sub
gUnpackDateLongErrorErr:
    igUnpackDateError = True
    gLogMsg slErrorMsg & "-Error in gUnpackDateLong: Year =" & ilYear & "  Month =" & ilMn & " Day =" & ilDy & " Passed ilDyMn =" & ilDyMn, "InvError" & CStr(tgUrf(0).iCode) & ".Txt", False
    llDate = 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackLength                   *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack length of time in btrieve*
'*                     format                          *
'*                                                     *
'*******************************************************
Sub gUnpackLength(ilHsSec As Integer, ilMinHr As Integer, slFormat As String, ilTest120 As Integer, slLength As String)
'
'   gUnpackLength ilHsSec, ilMinHr, slFormat, ilTest120, slTime
'   Where:
'       ilHsSec (I)- High order byte = hundredths of seconds; low order byte = seconds
'       ilMinHr (I)- High order byte = minute; low order byte = hours
'       slFormat (I)-"1" = hh:mm:ss
'                           "2" = hh mm'ss"
'                           "3" = hhHmmMssS
'       ilTest120 (I)- True=Test if length is 120 or less and show as seconds
'       slTime (O) - Length of Time in format
'

    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    Dim slSec As String     'Seconds
    Dim slMin As String     'Minutes
    Dim slHour As String    'Hours
    Dim llTotalSec As Long

    If (ilHsSec = 1) And (ilMinHr = 0) Then
        slLength = ""
        Exit Sub
    End If
    If (ilHsSec = 0) And (ilMinHr = 0) Then
        If slFormat = "1" Then
            slLength = "0"
            Exit Sub
        ElseIf slFormat = "3" Then
            slLength = "0s"
            Exit Sub
        Else
            slLength = "0" & """"
            Exit Sub
        End If
    End If
    ilSec = ilHsSec \ 256    'Obtain seconds
    slSec = Trim$(str$(ilSec))
    If Len(slSec) = 1 Then
        slSec = "0" & slSec
    End If
    ilMin = ilMinHr And &HFF 'Obtain Minutes
    slMin = Trim$(str$(ilMin))
    ilHour = ilMinHr \ 256   'Obtain month
    slHour = Trim$(str$(ilHour))
    If ilTest120 Then
    llTotalSec = 60 * ilHour
        llTotalSec = 60 * llTotalSec + 60 * ilMin + ilSec
        If llTotalSec <= 120 Then
            ilHour = 0
            ilMin = 0
            ilSec = llTotalSec
            slSec = Trim$(str$(ilSec))
            If Len(slSec) = 1 Then
                slSec = "0" & slSec
            End If
        End If
    End If
    Select Case slFormat
        Case "1"    'hh:mm:ss
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                slLength = slHour & ":" & slMin & ":" & slSec
            ElseIf ilMin <> 0 Then
                slLength = slMin & ":" & slSec
            Else
                slLength = slSec
            End If
        Case "2"    'hh mm'ss"
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                If (ilMin = 0) And (ilSec = 0) Then
                    slLength = slHour
                ElseIf ilSec = 0 Then
                    slLength = slHour & " " & slMin & "'"
                Else
                    slLength = slHour & " " & slMin & "'" & slSec & """"
                End If
            ElseIf ilMin <> 0 Then
                If ilSec = 0 Then
                    slLength = slMin & "'"
                Else
                    slLength = slMin & "'" & slSec & """"
                End If
            Else
                slLength = slSec & """"
            End If
        Case "3"    'hhHmmMssS
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                If (ilMin = 0) And (ilSec = 0) Then
                    slLength = slHour & "h"
                ElseIf ilSec = 0 Then
                    slLength = slHour & "h" & slMin & "m"
                Else
                    slLength = slHour & "h" & slMin & "m" & slSec & "s"
                End If
            ElseIf ilMin <> 0 Then
                If ilSec = 0 Then
                    slLength = slMin & "m"
                Else
                    slLength = slMin & "m" & slSec & "s"
                End If
            Else
                slLength = slSec & "s"
            End If
        Case Else
            slLength = ""
    End Select
    slLength = Trim$(slLength)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackTime                     *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack time in btrieve format   *
'*                                                     *
'*******************************************************
Sub gUnpackTime(ilHsSec As Integer, ilMinHr As Integer, slStyle As String, slFormat As String, slTime As String)
'
'   gUnpackTime ilHsSec, ilMinHr, slStyle, slFormat, slTime
'   Where:
'       ilHsSec (I)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (I)- High order byte = minute; low order byte = hours
'       slStyle (I)- "A" = AM or PM style
'                    "M" = Military style (HHMM:SS)
'       slFormat (I)- "1" = hours, min, sec
'                     "2" = hours, minutes (no seconds)
'                     "3" = minutes, seconds (no hours)
'                     "4" = hours, min, sec if sec defined; hours, minutes if seconds = 0
'       slTime (O) - Time in AM/PM format
'

    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    Dim dlTimeSerial As Double

    If (ilHsSec = 1) And (ilMinHr = 0) Then
        slTime = ""
        Exit Sub
    End If
    If (ilHsSec = 0) And (ilMinHr = 0) Then
        If slStyle = "M" Then
            If slFormat = "1" Then
                slTime = "0000:00"
            ElseIf slFormat = "2" Then
                slTime = "0000"
            ElseIf slFormat = "3" Then
                slTime = "00:00"
            Else
                slTime = ""
            End If
        Else
            slTime = "12M"
        End If
        Exit Sub
    End If
    ilSec = ilHsSec \ 256  'Obtain seconds
    ilMin = ilMinHr And &HFF 'Obtain Minutes
    ilHour = ilMinHr \ 256  'Obtain month
    dlTimeSerial = TimeSerial(ilHour, ilMin, ilSec)
    If slStyle = "A" Then
        If (slFormat = "1") Or ((slFormat = "4") And (ilSec > 0)) Then
            slTime = Format$(dlTimeSerial, "h:nn:ssAM/PM")
        ElseIf (slFormat = "2") Or ((slFormat = "4") And (ilSec = 0)) Then
            slTime = Format$(dlTimeSerial, "h:nnAM/PM")
        ElseIf slFormat = "3" Then
            slTime = Format$(dlTimeSerial, "n:ssAM/PM")
        Else
            slTime = ""
        End If
        slTime = gUnformatTime(slTime)
    ElseIf slStyle = "M" Then
        If (slFormat = "1") Or ((slFormat = "4") And (ilSec > 0)) Then
            slTime = Format$(dlTimeSerial, "hhnn:ss")
        ElseIf (slFormat = "2") Or ((slFormat = "4") And (ilSec = 0)) Then
            slTime = Format$(dlTimeSerial, "hhnn")
        ElseIf slFormat = "3" Then
            slTime = Format$(dlTimeSerial, "nn:ss")
        Else
            slTime = ""
        End If
    Else
        slTime = ""
    End If
    slTime = Trim$(slTime)
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackTimeLong                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack time in btrieve format   *
'*                                                     *
'*******************************************************
Sub gUnpackTimeLong(ilHsSec As Integer, ilMinHr As Integer, ilChk12M As Integer, llTime As Long)
'
'   gUnpackTimeLong ilHsSec, ilMinHr, ilChk12M, llTime
'   Where:
'       ilHsSec (I)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (I)- High order byte = minute; low order byte = hours
'       ilChk12M(I)- True=If 12M (0) convert to 86400 (24*3600)- handle end time
'                    False=Leave 12m as (0)
'       llTime (O) - Time as long
'

    Dim llSec As Long    'Seconds
    Dim llMin As Long    'Minutes
    Dim llHour As Long   'Hours

    If (ilHsSec = 1) And (ilMinHr = 0) Then
        llTime = -1
        Exit Sub
    End If
    If (ilHsSec = 0) And (ilMinHr = 0) Then
        If ilChk12M Then
            llTime = 86400
        Else
            llTime = 0
        End If
        Exit Sub
    End If
    llSec = ilHsSec \ 256  'Obtain seconds
    llMin = ilMinHr And &HFF 'Obtain Minutes
    llHour = ilMinHr \ 256  'Obtain month
    llTime = llHour * 3600 + llMin * 60 + llSec
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gValidDate                      *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if date is valid           *
'*                                                     *
'*******************************************************
Function gValidDate(slInpDate As String) As Integer
'
'   ilRet = gValidDate (slDate)
'   Where:
'       slDate (I) - Date to be checked
'       ilRet (O)- Yes (or True) means date is OK
'                       No ( or False) means an error in date format
'

    Dim ilDate As Integer
    Dim slDate As String
    Dim slAnyDate As String
    Dim ilPosMonth As Integer
    Dim ilPosDay As Integer

    slAnyDate = Trim$(slInpDate)
    If Len(slAnyDate) <> 0 Then
        slDate = gAdjYear(slAnyDate)
        On Error GoTo gValidDateErr
        ilPosMonth = InStr(1, slDate, "/", vbTextCompare)
        If ilPosMonth <= 0 Then
            gValidDate = NO
            Exit Function
        End If
        ilPosDay = InStr(ilPosMonth + 1, slDate, "/", vbTextCompare)
        'ilDate = Day(slDate)
        If ilPosDay <= 0 Then
            If Val(Mid$(slDate, ilPosMonth + 1)) <> Day(slDate) Then
                gValidDate = NO
                Exit Function
            End If
        Else
            If Val(Mid$(slDate, ilPosMonth + 1, ilPosDay - ilPosMonth - 1)) <> Day(slDate) Then
                gValidDate = NO
                Exit Function
            End If
        End If
        'ilDate = Month(slDate)
        If Val(Left$(slDate, ilPosMonth - 1)) <> Month(slDate) Then
            gValidDate = NO
            Exit Function
        End If
        ilDate = Year(slDate)
        If (ilDate < 1970) Or (ilDate > 2069) Then
            gValidDate = NO
            Exit Function
        End If
    Else
        gValidDate = NO
        Exit Function
    End If
    On Error GoTo 0
    gValidDate = YES
    Exit Function
gValidDateErr:
    On Error GoTo 0
    gValidDate = NO
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gValidLength                    *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if time length is valid    *
'*                                                     *
'*******************************************************
Function gValidLength(slInpLength As String) As Integer
'
'   ilRet = gValidLength (slLength)
'   Where:
'       slLength (I) - Time length to be checked
'       ilRet (O)- Yes (or True) means Length is OK
'                       No ( or False) means an error in length format
'

    Dim ilPos As Integer
    Dim slLen As String
    Dim slHour As String
    Dim llHour As Long
    Dim slMin As String
    Dim ilMin As Integer
    Dim slSec As String
    Dim ilSec As Integer
    Dim ilFormat As Integer
    Dim slLength As String

    slLength = Trim$(slInpLength)
    If Len(slLength) = 0 Then
        gValidLength = YES
        Exit Function
    End If
    slHour = ""
    slMin = ""
    slSec = ""
    slLen = Trim$(slLength)
    ilPos = InStr(1, slLen, "::")
    If ilPos > 0 Then
        gValidLength = NO
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hm", 1)
    If ilPos > 0 Then
        gValidLength = NO
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hs", 1)
    If ilPos > 0 Then
        gValidLength = NO
        Exit Function
    End If
    ilPos = InStr(1, slLen, "ms", 1)
    If ilPos > 0 Then
        gValidLength = NO
        Exit Function
    End If
    ilPos = InStr(1, slLen, ":")
    If ilPos > 0 Then
        ilFormat = 1
    Else
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            ilFormat = 3
        Else
            ilPos = InStr(1, slLen, "m", 1)
            If ilPos > 0 Then
                ilFormat = 3
            Else
                ilPos = InStr(1, slLen, "s", 1)
                If ilPos > 0 Then
                    ilFormat = 3
                Else
                    ilFormat = 2
                End If
            End If
        End If
    End If
    If ilFormat = 2 Then 'hh mm'ss"
        ilPos = InStr(1, slLen, " ")
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "'")
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
                slLen = Mid$(slLen, ilPos + 1)
            End If
         Else
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
                slLen = Mid$(slLen, ilPos + 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                    slLen = ""
                End If
            End If
         End If
         If Len(slLen) > 0 Then
            gValidLength = NO
            Exit Function
         End If
    ElseIf ilFormat = 3 Then
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "m", 1)
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
                slLen = Mid$(slLen, ilPos + 1)
            End If
         Else
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
                slLen = Mid$(slLen, ilPos + 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                    slLen = ""
                End If
            End If
         End If
         If Len(slLen) > 0 Then
            gValidLength = NO
            Exit Function
         End If
    Else    'format hh:mm:ss
        ilPos = InStr(slLen, ":")
        If ilPos > 0 Then   'Might be hour/min/sec or min/sec only
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, ":")
            If ilPos > 0 Then
                slMin = Left$(slLen, ilPos - 1)
                slSec = Mid$(slLen, ilPos + 1)
            Else
                slMin = slLen
            End If
        Else
            slHour = slLen
        End If
    End If
    On Error GoTo gValidLengthErr
    If slHour <> "" Then
        If Len(slHour) > 2 Then
            gValidLength = NO
            Exit Function
        End If
        llHour = CLng(slHour) 'Val(slHour)
    Else
        llHour = 0
    End If
    If slMin <> "" Then
        If Len(slMin) > 2 Then
            gValidLength = NO
            Exit Function
        End If
        ilMin = CLng(slMin) 'Val(slMin)
    Else
        ilMin = 0
    End If
    If slSec <> "" Then
        If Len(slSec) > 3 Then
            gValidLength = NO
            Exit Function
        End If
        ilSec = CLng(slSec) 'Val(slSec)
    Else
        ilSec = 0
    End If
    If (llHour < 0) Or (llHour > 24) Then
        gValidLength = NO
        Exit Function
    End If
    If (llHour <> 0) Then
        If (ilMin < 0) Or (ilMin > 59) Then
            gValidLength = NO
            Exit Function
        End If
    End If
    If llHour <> 0 Then
'    If (ilHour <= 0) And (ilMin <= 0) Then
'        If (ilSec < 0) Or (ilSec > 120) Then
'            gValidLength = No
'            Exit Function
'        End If
'    Else
        If (ilSec < 0) Or (ilSec > 59) Then
            gValidLength = NO
            Exit Function
        End If
    End If
    If llHour * 3600& + CLng(ilMin) * 60& + CLng(ilSec) > 86400 Then
        gValidLength = NO
        Exit Function
    End If
    gValidLength = YES
    Exit Function
gValidLengthErr:
    On Error GoTo 0
    gValidLength = NO
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gValidTime                      *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Time if time input is valid     *
'*                                                     *
'*******************************************************
Function gValidTime(slInpTime As String) As Integer
'
'   ilRet = gValidTime (slTime)
'   Where:
'       slTime (I) - Time to be checked
'       ilRet (O)- Yes (or True) means time is OK
'                       No ( or False) means an error in time format
'

    Dim ilTime As Integer
    Dim slTime As String
    Dim slTimeChk As String

    slTimeChk = Trim$(slInpTime)
    slTime = slTimeChk
    If Len(slTime) <> 0 Then
        slTime = gConvertTime(slTime)
        On Error GoTo gValidTimeErr
        ilTime = Second(slTime)
        ilTime = Minute(slTime)
        ilTime = Hour(slTime)
        On Error GoTo 0
    End If
    gValidTime = YES
    Exit Function
gValidTimeErr:
    On Error GoTo 0
    gValidTime = NO
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gWeekDayLong                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Week day (0=Monday,      *
'*                      1=Tuesday,...,6=Sunday)        *
'*                                                     *
'*******************************************************
Function gWeekDayLong(llAnyDate As Long) As Integer
'
'   iRetDay = gWeekDayLong(lDate)
'   Where:
'       lDate (I)- Date to obtain week day for (serial)
'       iRetDay (O)- Week day (0=Mon, 1=Tue,..,6=Sun)
'

    Dim ilWeekDay As Integer

    ilWeekDay = Weekday(llAnyDate) - 2
    If ilWeekDay < 0 Then
        ilWeekDay = 6
    End If
    gWeekDayLong = ilWeekDay
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gWeekDayStr                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Week day (0=Monday,      *
'*                      1=Tuesday,...,6=Sunday)        *
'*                                                     *
'*******************************************************
Function gWeekDayStr(slInpDate As String) As Integer
'
'   iRetDay = gWeekDayStr(sDate)
'   Where:
'       sDate (I)- Date to obtain week day for
'       iRetDay (O)- Week day (0=Mon, 1=Tue,..6=Sun)
'

    Dim slDate As String
    Dim ilWeekDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    If slAnyDate = "" Then
        ilWeekDay = 0
        Exit Function
    End If
    slDate = gAdjYear(slAnyDate)
    ilWeekDay = Weekday(slDate) - 2
    If ilWeekDay < 0 Then
        ilWeekDay = 6
    End If
    gWeekDayStr = ilWeekDay
End Function


Public Function gLengthToLong(slLength As String) As Long
    ReDim ilLength(0 To 1) As Integer
    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    gPackLength Trim$(slLength), ilLength(0), ilLength(1)
    ilSec = ilLength(0) \ 256    'Obtain seconds
    ilMin = ilLength(1) And &HFF 'Obtain Minutes
    ilHour = ilLength(1) \ 256   'Obtain month
    gLengthToLong = 3600 * CLng(ilHour) + 60 * ilMin + ilSec
End Function

Public Function gIsDate(slInpDate As String) As Integer

    Dim ilDate As Integer
    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    If Len(slAnyDate) <> 0 Then
        If Asc(right$(slAnyDate, 1)) < KEY0 Or (Asc(right$(slAnyDate, 1)) > KEY9) Then
            gIsDate = False
            Exit Function
        End If
        slDate = gAdjYear(slAnyDate)
        On Error GoTo gValidDateErr
        ilDate = Day(slDate)
        ilDate = Month(slDate)
        ilDate = Year(slDate)
        If (ilDate < 1970) Or (ilDate > 2069) Then
            gIsDate = False
            Exit Function
        End If
    Else
        gIsDate = False
        Exit Function
    End If
    On Error GoTo 0
    gIsDate = True
    Exit Function
gValidDateErr:
    On Error GoTo 0
    gIsDate = False
    Exit Function

End Function

Public Function gDayStrToNames(slDayStr As String, slDefaultName As String, ilNoChars As Integer)
    Dim slStr As String
    slStr = slDayStr
    If slStr = "YYYYYYY" Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", 2)
        End If
    ElseIf (slStr = "YYYYYY ") Or (slStr = "YYYYYYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", 2)
        End If
    ElseIf (slStr = "YYYYY  ") Or (slStr = "YYYYYNN") Then
        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
    ElseIf (slStr = "YYYY   ") Or (slStr = "YYYYNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", 2)
        End If
    ElseIf (slStr = "YYY    ") Or (slStr = "YYYNNNN") Then
        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
    ElseIf (slStr = "YY     ") Or (slStr = "YYNNNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", 2)
        End If
    ElseIf (slStr = " YYYYYY") Or (slStr = "NYYYYYY") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Sunday", 2)
        End If
    ElseIf (slStr = " YYYYY ") Or (slStr = "NYYYYYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Saturday", 2)
        End If
    ElseIf (slStr = " YYYY  ") Or (slStr = "NYYYYNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Friday", ilNoChars)
        End If
    ElseIf (slStr = " YYY   ") Or (slStr = "NYYYNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Thursday", 2)
        End If
    ElseIf (slStr = " YY    ") Or (slStr = "NYYNNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Wednesday", ilNoChars)
        End If
    ElseIf (slStr = "  YYYYY") Or (slStr = "NNYYYYY") Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", 2)
        End If
    ElseIf (slStr = "  YYYY ") Or (slStr = "NNYYYYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", 2)
        End If
    ElseIf (slStr = "  YYY  ") Or (slStr = "NNYYYNN") Then
        slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
    ElseIf (slStr = "  YY   ") Or (slStr = "NNYYNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", 2)
        End If
    ElseIf (slStr = "   YYYY") Or (slStr = "NNNYYYY") Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Sunday", 2)
        End If
    ElseIf (slStr = "   YYY ") Or (slStr = "NNNYYYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Saturday", 2)
        End If
    ElseIf (slStr = "   YY  ") Or (slStr = "NNNYYNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Friday", ilNoChars)
        End If
    ElseIf (slStr = "     Y ") Or (slStr = "NNNNNYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Saturday", 2)
        End If
    ElseIf (slStr = "     YY") Or (slStr = "NNNNNYY") Then
        slStr = Left$("Saturday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
    ElseIf (slStr = "Y      ") Or (slStr = "YNNNNNN") Then
        slStr = Left$("Monday", ilNoChars)
    ElseIf (slStr = " Y     ") Or (slStr = "NYNNNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2)
        End If
    ElseIf (slStr = "  Y    ") Or (slStr = "NNYNNNN") Then
        slStr = Left$("Wednesday", ilNoChars)
    ElseIf (slStr = "   Y   ") Or (slStr = "NNNYNNN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2)
        End If
    ElseIf (slStr = "    Y  ") Or (slStr = "NNNNYNN") Then
        slStr = Left$("Friday", ilNoChars)
    ElseIf (slStr = "     Y ") Or (slStr = "NNNNNYN") Then
        If ilNoChars > 1 Then
            slStr = Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Ssturday", 2)
        End If
    ElseIf (slStr = "      Y") Or (slStr = "NNNNNNY") Then
        If ilNoChars > 1 Then
            slStr = Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Sunday", 2)
        End If
    Else
        slStr = slDefaultName
    End If
    gDayStrToNames = slStr
End Function
