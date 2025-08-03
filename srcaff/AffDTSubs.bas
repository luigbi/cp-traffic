Attribute VB_Name = "modDTSubs"
'******************************************************
'*  modDatesTimes - contains various date and time conversion/validation routines
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Public Function gCalcMoBehind(mStartMo As String, mEndMo As String) As Integer

    Dim slDateDiff As String
    
    slDateDiff = DateDiff("m", mEndMo, mStartMo)
    gCalcMoBehind = Val(slDateDiff) - Val(sgNumMoToRetain)
    Exit Function

End Function


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
    Dim ilYear As Integer
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
        slYear = Str$(Year(gNow()))
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
    
    ilYear = Val(slYear)
    If (ilYear >= 0) And (ilYear <= 69) Then
        ilYear = 2000 + ilYear
    ElseIf (ilYear >= 70) And (ilYear <= 99) Then
        ilYear = 1900 + ilYear
    End If
    gAdjYear = slMonthDay & "/" & Trim$(Str$(ilYear))
    Exit Function
gAdjYearErr:
    On Error GoTo 0
    gAdjYear = slAnyDate
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

    slDate = Trim$(sgNowDate)
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

'*********************************************************************
'TTP 10403 - Affiliate Spot MGMT report showing extra vehicles when run for a single vehicle
'               gRandomDateTime (Partially taken from Traffic Reports)
'               <Output> slDate - current date (xx/xx/xx)
'                        slTime - current time (xx:xx:xxa/p)
'                        Some routines may not use these return values
'                        slMonth - xx  (1-12)
'                        slDay - XX  (1-31)
'                        slYear - xxxx (19xx-20xx)
'               obtain random date and current time and return it
'               Random date is based on the low date of 1/1/1970 (value of 25569)
'                   Hi date of 12/31/2069 (value of 62093)
'                   Dates will be adjusted to lo value of 26000 and hi value of 62000)
'               in string format
'
'               Created:  10/26/20
'*********************************************************************
Sub gRandomDateTime(slDate As String, slTime As String, slMonth As String, slDay As String, slYear As String)
    Dim llRandomDate As Long
    Randomize
    llRandomDate = CLng(((62000 - 26000) + 1) * Rnd + 26000)
    slDate = Format$(llRandomDate, "ddddd")
    'gPackDate slDate, igNowDate(0), igNowDate(1)
    slTime = Format$(gNow(), "h:mm:ssAM/PM")
    'gPackTime slTime, igNowTime(0), igNowTime(1)
    'gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
End Sub


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
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    Dim slAnyDate As String

    slAnyDate = Format$(Trim$(slInpDate), "m/d/yyyy")
    slYear = ""
    slMonth = ""
    slDay = ""
    On Error GoTo gObtainYearMonthDayStrErr
    ilRet = gParseItem(slAnyDate, 1, "/", slMonth)
    If ilRet <> CSI_MSG_NONE Then
        Exit Sub
    End If
    ilRet = gParseItem(slAnyDate, 2, "/", slDay)
    If ilRet <> CSI_MSG_NONE Then
        Exit Sub
    End If
    ilRet = gParseItem(slAnyDate, 3, "/", slYear)
    If ilRet <> CSI_MSG_NONE Then
        slYear = Str$(Year(Now))
    End If
    ilYear = Val(slYear)
    If (ilYear >= 0) And (ilYear <= 69) Then
        ilYear = 2000 + ilYear
    ElseIf (ilYear >= 70) And (ilYear <= 99) Then
        ilYear = 1900 + ilYear
    End If
    slYear = Trim$(Str$(ilYear))
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
'*      Procedure Name:gObtainNextMonday               *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain first sunday from        *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainNextMonday(slInpDate As String) As String
'
'   sRetDate = gObtainNextMonday(sDate)
'   Where:
'       sDate (I)- Date to obtain previous monday from
'       sRetDate (O)- Previous monday including specified date
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Format$(Trim$(slInpDate), "m/d/yyyy")
    llDate = DateValue(gAdjYear(slAnyDate))
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbMonday
        llDate = llDate + 1
    Loop
    gObtainNextMonday = Format$(llDate, sgShowDateForm)

End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPrevMonday               *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain first sunday from        *
'*                     specified date, includes        *
'*                     specified date                  *
'*                                                     *
'*******************************************************
Function gObtainPrevMonday(slInpDate As String) As String
'
'   sRetDate = gObtainPrevMonday(sDate)
'   Where:
'       sDate (I)- Date to obtain previous monday from
'       sRetDate (O)- Previous monday including specified date
'

    Dim llDate As Long
    Dim slAnyDate As String

    slAnyDate = Format$(Trim$(slInpDate), "m/d/yyyy")
    llDate = DateValue(gAdjYear(slAnyDate))
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbMonday
        llDate = llDate - 1
    Loop
    gObtainPrevMonday = Format$(llDate, sgShowDateForm)

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
    Dim slAnyDate As String

    slAnyDate = Format$(Trim$(slInpDate), "m/d/yyyy")
    llDate = DateValue(gAdjYear(slAnyDate))
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbSunday
        llDate = llDate + 1
    Loop
    gObtainNextSunday = Format$(llDate, sgShowDateForm)

End Function
Function gObtainPrevSunday(slInpDate As String) As String
'   Dan M. 1/08/10 copied from traffic for new contacts form
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
'*      Procedure Name:gTimeToLong                     *
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
    'D.S. 07/12/01 changed slTime to slAnyTime below
    'ilPos = InStr(slTime, "-")
    ilPos = InStr(slAnyTime, "-")
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
    Dim ilMatchDay As Integer
    Dim slAnyDate As String

    slAnyDate = Trim$(slInpDate)
    slAnyDate = gObtainEndStd(slAnyDate)
    llDate = DateValue(gAdjYear(slAnyDate))
    slAnyDate = gObtainEndStd(Format$(llDate - 40, "m/d/yyyy"))
    llDate = DateValue(gAdjYear(slAnyDate)) + 1
    gObtainStartStd = Format$(llDate, sgShowDateForm)
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
    Dim ilMatchDay As Integer
    Dim slDate As String
    Dim slAnyDate As String

    slAnyDate = Format$(Trim$(slInpDate), "m/d/yyyy")
    slDate = gObtainNextSunday(slAnyDate)
    llDate = DateValue(gAdjYear(slDate))
    Do While Month(llDate) = Month(llDate + 7)
        llDate = llDate + 7
    Loop
    gObtainEndStd = Format$(llDate, sgShowDateForm)
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
    If (InStr(slFixedTime, "A") = 0) And (InStr(slFixedTime, "P") = 0) And (InStr(slFixedTime, "N") = 0) And (InStr(slFixedTime, "M") = 0) Then
        slFixedTime = Format$(slFixedTime, "hh:mm:ss am/pm")
    End If
    ilPos1 = InStr(slFixedTime, "N")
    If ilPos1 <> 0 Then
        slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "PM"
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    If (ilPos1 = 0) And (ilPos2 = 0) Then
        ilPos1 = InStr(slFixedTime, "M")
        If ilPos1 <> 0 Then
            slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "AM"
        End If
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    ilPos3 = InStr(slFixedTime, "M")
    If (ilPos3 = 0) And ((ilPos1 <> 0) Or (ilPos2 <> 0)) Then
        slFixedTime = slFixedTime & "M"
    End If
    If InStr(slFixedTime, ":") = 0 Then
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
'*      Procedure Name:gIsTime                      *
'*                                                     *
'*             Created:5/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Time if time input is valid     *
'*                                                     *
'*******************************************************
Function gIsTime(slInpTime As String) As Integer
'
'   ilRet = gIsTime (slTime)
'   Where:
'       slTime (I) - Time to be checked
'       ilRet (O)- Yes (or True) means time is OK
'                       No ( or False) means an error in time format
'

    Dim ilTime As Integer
    Dim slTime As String
    Dim slTimeChk As String
    Dim slLastPos As String * 1
        
    slTimeChk = Trim$(slInpTime)
    slTime = slTimeChk
    If Len(slTime) <> 0 Then
        'make sure that time ends in either AM, PM, M or N or its not a valid time
        slLastPos = UCase$(right$(slTime, 1))
        If (slLastPos <> "A") And (slLastPos <> "P") And (slLastPos <> "M") And (slLastPos <> "N") Then
            gIsTime = False
            Exit Function
        End If
        slTime = gConvertTime(slTime)
        On Error GoTo gIsTimeErr
        
        'The case ":00" causes a run-time error in the Second function below
        If Left$(slTime, 1) = ":" Then
            gIsTime = False
            Exit Function
        End If
        
        If InStr(slTime, ":P") <> 0 Or InStr(slTime, ":M") <> 0 Or InStr(slTime, ":A") <> 0 Then
            gIsTime = False
            Exit Function
        End If
        
        'make sure that seconds and minutes are from 0-59 and hours are from 0-23
        ilTime = Second(slTime)
        If (ilTime < 0) Or (ilTime > 59) Then
            gIsTime = False
            Exit Function
        End If
        ilTime = Minute(slTime)
        If (ilTime < 0) Or (ilTime > 59) Then
            gIsTime = False
            Exit Function
        End If
        ilTime = Hour(slTime)
        If (ilTime < 0) Or (ilTime > 23) Then
            gIsTime = False
            Exit Function
        End If
        
        On Error GoTo 0
    End If
    gIsTime = True
    Exit Function
gIsTimeErr:
    On Error GoTo 0
    gIsTime = False
    Exit Function
End Function



Public Function gLongToTime(llInTime As Long) As String
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilSec As Integer
    Dim llTime As Long
    Dim slTime As String
    
    llTime = llInTime
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
            slTime = Trim$(Str$(ilHour))
        Else
            slTime = Trim$(Str$(ilHour - 12))
        End If
    End If
    If ilMin <> 0 Then
        If ilMin < 10 Then
            slTime = slTime & ":0" & Trim$(Str$(ilMin))
        Else
            slTime = slTime & ":" & Trim$(Str$(ilMin))
        End If
    End If
    If ilSec <> 0 Then
        If ilMin = 0 Then
            slTime = slTime & ":00:"
        Else
            slTime = slTime & ":"
        End If
        If ilSec < 10 Then
            slTime = slTime & "0" & Trim$(Str$(ilSec))
        Else
            slTime = slTime & Trim$(Str$(ilSec))
        End If
    End If
    If ilHour < 12 Then
        slTime = slTime & "AM"
    Else
        slTime = slTime & "PM"
    End If
    gLongToTime = slTime
End Function

Public Function gValidTimeForm(sTime) As Integer


End Function

Public Function gValidDateForm(sDate) As Integer

    Dim ilRet As Integer
    
    gValidDateForm = False
    
    If sDate = "m/d/yy" Or sDate = "mm/dd/yy" Or sDate = "mm/dd/yyyy" Then
        gValidDateForm = True
    End If

    If sDate = "m-d-yy" Or sDate = "mm-dd-yy" Or sDate = "mm-dd-yyyy" Then
        gValidDateForm = True
    End If


End Function

Public Function gConvertMilitaryHourToRegTime(sTime As String) As String

    Dim slRetTime As String
    
    Select Case sTime
    Case 0
        slRetTime = "12:00a"
    Case 1
        slRetTime = "1:00a"
    Case 2
        slRetTime = "2:00a"
    Case 3
        slRetTime = "3:00a"
    Case 4
        slRetTime = "4:00a"
    Case 5
        slRetTime = "5:00a"
    Case 6
        slRetTime = "6:00a"
    Case 7
        slRetTime = "7:00a"
    Case 8
        slRetTime = "8:00a"
    Case 9
        slRetTime = "9:00a"
    Case 10
        slRetTime = "10:00a"
    Case 11
        slRetTime = "11:00a"
    Case 12
        slRetTime = "12:00p"
    Case 13
        slRetTime = "1:00p"
    Case 14
        slRetTime = "2:00p"
    Case 15
        slRetTime = "3:00p"
    Case 16
        slRetTime = "4:00p"
    Case 17
        slRetTime = "5:00p"
    Case 18
        slRetTime = "6:00p"
    Case 19
        slRetTime = "7:00p"
    Case 20
        slRetTime = "8:00p"
    Case 21
        slRetTime = "9:00p"
    Case 22
        slRetTime = "10:00p"
    Case 23
        slRetTime = "11:00p"
    End Select

    gConvertMilitaryHourToRegTime = slRetTime

End Function
Public Function gConvertRegHourToMilitaryHour(slTime As String) As String
    'for 10927
    '"12:30a"  looks at 12 and 'a' and returns 00
    
    Dim slRetTime As String
    Dim ilPos As Integer
    Dim slHour As String
    Dim slPosition As String
    
    ilPos = InStr(slTime, ":")
    If ilPos <> 0 Then
        slHour = Left$(slTime, ilPos - 1)
    End If
    slPosition = right$(slTime, 1)
    
    Select Case slHour
    Case 12
        If UCase(slPosition) = "A" Then
            slRetTime = "00"
        Else
            slRetTime = "12"
        End If
    Case 1
        If UCase(slPosition) = "A" Then
            slRetTime = "01"
        Else
            slRetTime = "13"
        End If
    Case 2
        If UCase(slPosition) = "A" Then
            slRetTime = "02"
        Else
            slRetTime = "14"
        End If
    Case 3
        If UCase(slPosition) = "A" Then
            slRetTime = "03"
        Else
            slRetTime = "15"
        End If
    Case 4
        If UCase(slPosition) = "A" Then
            slRetTime = "04"
        Else
            slRetTime = "16"
        End If
    Case 5
        If UCase(slPosition) = "A" Then
            slRetTime = "05"
        Else
            slRetTime = "17"
        End If
    Case 6
        If UCase(slPosition) = "A" Then
            slRetTime = "06"
        Else
            slRetTime = "18"
        End If
    Case 7
        If UCase(slPosition) = "A" Then
            slRetTime = "07"
        Else
            slRetTime = "19"
        End If
    Case 8
        If UCase(slPosition) = "A" Then
            slRetTime = "08"
        Else
            slRetTime = "20"
        End If
    Case 9
        If UCase(slPosition) = "A" Then
            slRetTime = "09"
        Else
            slRetTime = "21"
        End If
    Case 10
        If UCase(slPosition) = "A" Then
            slRetTime = "10"
        Else
            slRetTime = "22"
        End If
    Case 11
        If UCase(slPosition) = "A" Then
            slRetTime = "11"
        Else
            slRetTime = "23"
        End If
    End Select

    gConvertRegHourToMilitaryHour = slRetTime

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

'6/29/06: change gGetAstInfo to use API call
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
'6/29/06: End of change


'6/29/06: change gGetAstInfo to use API call
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
'6/29/06: End of Change

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
        If slFormat = "1" Then
            slTime = Format$(dlTimeSerial, "h:nn:ssAM/PM")
        ElseIf slFormat = "2" Then
            slTime = Format$(dlTimeSerial, "h:nnAM/PM")
        ElseIf slFormat = "3" Then
            slTime = Format$(dlTimeSerial, "n:ssAM/PM")
        Else
            slTime = ""
        End If
        slTime = gUnformatTime(slTime)
    ElseIf slStyle = "M" Then
        If slFormat = "1" Then
            slTime = Format$(dlTimeSerial, "hhnn:ss")
        ElseIf slFormat = "2" Then
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

Public Function gObtainYearStartDate(slInDate As String) As String
    Dim slDate As String
    Dim slYearNo As String
    
    slDate = gObtainEndStd(slInDate)
    slYearNo = Format(slDate, "yy")
    If Len(slYearNo) = 1 Then
        slYearNo = "0" & slYearNo
    End If
    slDate = "1/15/" & slYearNo
    gObtainYearStartDate = gObtainStartStd(slDate)

End Function

Public Function gTimeString(Seconds As Long, Optional Verbose As Boolean = False) As String

    'if verbose = false, returns
    'something like
    '02:22.08
    'if true, returns
    '2 hours, 22 minutes, and 8 seconds
    
    Dim llHrs As Long
    Dim llMinutes As Long
    Dim llSeconds As Long
    
    llSeconds = Seconds
    
    llHrs = Int(llSeconds / 3600)
    llMinutes = (Int(llSeconds / 60)) - (llHrs * 60)
    llSeconds = Int(llSeconds Mod 60)
    
    Dim sAns As String
    
    
    If llSeconds = 60 Then
        llMinutes = llMinutes + 1
        llSeconds = 0
    End If
    
    If llMinutes = 60 Then
        llMinutes = 0
        llHrs = llHrs + 1
    End If
    
    sAns = Format(CStr(llHrs), "#####0") & ":" & _
      Format(CStr(llMinutes), "00") & "." & _
      Format(CStr(llSeconds), "00")
    
    If Verbose Then sAns = TimeStringtoEnglish(sAns)
    gTimeString = sAns

End Function

Private Function TimeStringtoEnglish(sTimeString As String) As String

    Dim slAnswer As String
    Dim slHour As String
    Dim slMin As String
    Dim ilTemp As Integer
    Dim slTemp As String
    Dim ilPos As Integer
    
    ilPos = InStr(sTimeString, ":") - 1
    
    slHour = Left$(sTimeString, ilPos)
    If CLng(slHour) <> 0 Then
        slAnswer = CLng(slHour) & " hour"
        If CLng(slHour) > 1 Then slAnswer = slAnswer & "s"
        slAnswer = slAnswer & ", "
    End If
    
    slMin = Mid$(sTimeString, ilPos + 2, 2)
    
    ilTemp = slMin
    
    If slMin = "00" Then
       slAnswer = IIf(Len(slAnswer), slAnswer & "0 minutes, and ", "")
    Else
       slTemp = IIf(ilTemp = 1, " minute", " minutes")
       slTemp = IIf(Len(slAnswer), slTemp & ", and ", slTemp & " and ")
       slAnswer = slAnswer & Format$(ilTemp, "##") & slTemp
    End If
    
    ilTemp = Val(right$(sTimeString, 2))
    slMin = Format$(ilTemp, "#0")
    slAnswer = slAnswer & slMin & " second"
    If ilTemp <> 1 Then slAnswer = slAnswer & "s"
    
    TimeStringtoEnglish = slAnswer

End Function


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
    If (ilType = 0) Then   'Std
        slDate = gObtainEndStd(slAnyDate)
        slDate = gAdjYear(slDate)
'    ElseIf ilType = 4 Then  'Corp on Jan-Dec year
'        'slStartDate = gObtainStartCorp(slAnyDate, True)
'        'slEndDate = gObtainEndCorp(slAnyDate, True)
'        'slDate = Format$(gDateValue(slStartDate) + (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 2, "m/d/yy")
'        'slDate = gAdjYear(slDate)
'        llDate = gDateValue(slAnyDate)
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
'            For ilIndex = 1 To 12 Step 1
'                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
'                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
'                If (llDate >= llSDate) And (llDate <= llEDate) Then
'                    ilYear = tgMCof(ilLoop).iYear
'                    ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
'                    If ilMonth > 12 Then
'                        ilMonth = ilMonth - 12
'                    Else
'                        ilYear = ilYear - 1
'                    End If
'                    Exit Sub
'                End If
'            Next ilIndex
'        Next ilLoop
'        ilMonth = 0
'        ilYear = 0
'        Exit Sub
'    ElseIf ilType = 5 Then  'Corp on Fiscal (Oct-Sept)
'        'slStartDate = gObtainStartCorp(slAnyDate, True)
'        'slEndDate = gObtainEndCorp(slAnyDate, True)
'        'slDate = Format$(gDateValue(slStartDate) + (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 2, "m/d/yy")
'        'slDate = gAdjYear(slDate)
'        llDate = gDateValue(slAnyDate)
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
'            For ilIndex = 1 To 12 Step 1
'                gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilIndex), tgMCof(ilLoop).iStartDate(1, ilIndex), llSDate
'                gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilIndex), tgMCof(ilLoop).iEndDate(1, ilIndex), llEDate
'                If (llDate >= llSDate) And (llDate <= llEDate) Then
'                    ilYear = tgMCof(ilLoop).iYear
'                    ilMonth = tgMCof(ilLoop).iStartMnthNo + ilIndex - 1
'                    If ilMonth > 12 Then
'                        ilMonth = ilMonth - 12
'                    End If
'                    Exit Sub
'                End If
'            Next ilIndex
'        Next ilLoop
'        ilMonth = 0
'        ilYear = 0
'        Exit Sub
    Else
        slDate = gAdjYear(slAnyDate)
    End If
    ilMonth = Month(slDate)
    ilYear = Year(slDate)
End Sub

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
    gMonthYearFormat = slStr & Trim$(Str$(Year(slDate)))
    Exit Function
gMonthYearFormatErr:
    gMonthYearFormat = ""
    On Error GoTo 0
    Exit Function
End Function
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
    If slStr = "YYYYYYY" Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Sunday", 2)
        End If
    ElseIf slStr = "YYYYYY " Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Saturday", 2)
        End If
    ElseIf slStr = "YYYYY  " Then
        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
    ElseIf slStr = "YYYY   " Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Thursday", 2)
        End If
    ElseIf slStr = "YYY    " Then
        slStr = Left$("Monday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
    ElseIf slStr = "YY     " Then
        If ilNoChars > 1 Then
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", ilNoChars)
        Else
            slStr = Left$("Monday", ilNoChars) & "-" & Left$("Tuesday", 2)
        End If
    ElseIf slStr = " YYYYYY" Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Sunday", 2)
        End If
    ElseIf slStr = " YYYYY " Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Saturday", 2)
        End If
    ElseIf slStr = " YYYY  " Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Friday", ilNoChars)
        End If
    ElseIf slStr = " YYY   " Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Thursday", 2)
        End If
    ElseIf slStr = " YY    " Then
        If ilNoChars > 1 Then
            slStr = Left$("Tuesday", ilNoChars) & "-" & Left$("Wednesday", ilNoChars)
        Else
            slStr = Left$("Tuesday", 2) & "-" & Left$("Wednesday", ilNoChars)
        End If
    ElseIf slStr = "  YYYYY" Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Sunday", 2)
        End If
    ElseIf slStr = "  YYYY " Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Saturday", 2)
        End If
    ElseIf slStr = "  YYY  " Then
        slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
    ElseIf slStr = "  YY   " Then
        If ilNoChars > 1 Then
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", ilNoChars)
        Else
            slStr = Left$("Wednesday", ilNoChars) & "-" & Left$("Thursday", 2)
        End If
    ElseIf slStr = "   YYYY" Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Sunday", 2)
        End If
    ElseIf slStr = "   YYY " Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Saturday", 2)
        End If
    ElseIf slStr = "   YY  " Then
        If ilNoChars > 1 Then
            slStr = Left$("Thursday", ilNoChars) & "-" & Left$("Friday", ilNoChars)
        Else
            slStr = Left$("Thursday", 2) & "-" & Left$("Friday", ilNoChars)
        End If
    ElseIf slStr = "     Y " Then
        If ilNoChars > 1 Then
            slStr = Left$("Saturday", ilNoChars)
        Else
            slStr = Left$("Saturday", 2)
        End If
    ElseIf slStr = "     YY" Then
        slStr = Left$("Saturday", ilNoChars) & "-" & Left$("Sunday", ilNoChars)
    ElseIf slStr = "      Y" Then
        If ilNoChars > 1 Then
            slStr = Left$("Sunday", ilNoChars)
        Else
            slStr = Left$("Sunday", 2)
        End If
    Else
        slStr = slDayImage
    End If
    gDayNames = slStr
End Function
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
