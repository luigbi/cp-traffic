Attribute VB_Name = "CommSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CommSubs.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Common Subs code
Option Explicit
Option Compare Text
Public sgExePath As String
Public sgDBPath As String


'*******************************************************
'*                                                     *
'*      Procedure Name:gAddStr                         *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Sum two strings (accuracy to    *
'*                     same decimal place as the       *
'*                     larger decimal place of the     *
'*                     two numbers)                    *
'*                                                     *
'*******************************************************
Function gAddStr(slAddStr1 As String, slAddStr2 As String) As String
'
'   sRet = gAddStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- First string to be added
'       sStr2 (I)- Second string to be added
'       sRet (O)- sStr1 + sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilPos As Integer
    Dim ilNoDec As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slAddStr1)
    slInput2 = Trim$(slAddStr2)
    clNum1 = Val(slInput1)
    ilPos = InStr(slInput1, ".")
    If ilPos = 0 Then
    ilNoDec1 = 0
    Else
    ilNoDec1 = Len(slInput1) - ilPos
    End If
    clNum2 = Val(slInput2)
    ilPos = InStr(slInput2, ".")
    If ilPos = 0 Then
    ilNoDec2 = 0
    Else
    ilNoDec2 = Len(slInput2) - ilPos
    End If
    clNum = clNum1 + clNum2
    If (ilNoDec1 > 0) Or (ilNoDec2 > 0) Then
    clAdj = 0.5
    If ilNoDec1 > ilNoDec2 Then
        ilNoDec = ilNoDec1
    Else
        ilNoDec = ilNoDec2
    End If
    For ilLoop = 1 To ilNoDec Step 1
        clAdj = clAdj / 10@
    Next ilLoop
    If clNum >= 0 Then
        clNum = clNum + clAdj
    Else
        clNum = clNum - clAdj
    End If
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        slWhole = Left$(slStr, ilPos - 1)
        slDec = Right$(slStr, Len(slStr) - ilPos)
        If Len(slDec) > ilNoDec Then
        slDec = Left$(slDec, ilNoDec)
        End If
        Do While (Len(slDec) < ilNoDec)
        slDec = slDec & "0"
        Loop
        gAddStr = slWhole & "." & slDec
    Else
        slStr = slStr & "."
        For ilLoop = 1 To ilNoDec Step 1
        slStr = slStr & "0"
        Next ilLoop
        gAddStr = slStr
    End If
    Else
'        clNum = clNum + .5
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        gAddStr = Left$(slStr, ilPos - 1)
    Else
        gAddStr = slStr
    End If
    End If
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
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slYear As String
    Dim ilRet As Integer
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    Dim slAnyDate As String
    slAnyDate = Trim$(slInpDate)
    On Error GoTo gAdjYearErr
    ilRet = gParseItem(slAnyDate, 1, "/", slMonth)
    If ilRet <> CP_MSG_NONE Then
        gAdjYear = slAnyDate
        Exit Function
    End If
    ilRet = gParseItem(slAnyDate, 2, "/", slDay)
    If ilRet <> CP_MSG_NONE Then
        gAdjYear = slAnyDate
        Exit Function
    End If
    ilRet = gParseItem(slAnyDate, 3, "/", slYear)
    If ilRet <> CP_MSG_NONE Then
        gAdjYear = slAnyDate
        Exit Function
    End If
    ilYear = Val(slYear)
    If (ilYear >= 0) And (ilYear <= 69) Then
        ilYear = 2000 + ilYear
    ElseIf (ilYear >= 70) And (ilYear <= 99) Then
        ilYear = 1900 + ilYear
    End If
    slDate = Trim$(slMonth) & "/" & Trim$(slDay) & "/" & Trim$(Str$(ilYear))
    gAdjYear = slDate
    Exit Function
gAdjYearErr:
    On Error GoTo 0
    gAdjYear = slAnyDate
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterStdAlone                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form on screen           *
'*                                                     *
'*******************************************************
Sub gCenterStdAlone(Frm As Form)
    Frm.Move (Screen.Width - Frm.Width) \ 2, (Screen.Height - Frm.Height) \ 2 + 175 '+ Screen.Height \ 10
End Sub
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
    If InStr(slTime, ":") = 0 Then
        If ilPos1 <> 0 Then
            If Len(slFixedTime) <= 4 Then
                slFixedTime = Left$(slFixedTime, ilPos1 - 1) & ":00AM"
            Else
                If Len(slFixedTime) <= 6 Then
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & Right$(slFixedTime, 4)
                Else
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & Right$(slFixedTime, 6)
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & Right$(slFixedTime, 4)
                End If
            End If
        Else
            If ilPos2 <> 0 Then
                If Len(slFixedTime) <= 4 Then
                    slFixedTime = Left$(slFixedTime, ilPos2 - 1) & ":00PM"
                Else
                    If Len(slFixedTime) <= 6 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & Right$(slFixedTime, 4)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & Right$(slFixedTime, 6)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & Right$(slFixedTime, 4)
                    End If
                End If
            Else
                If Len(slFixedTime) <= 2 Then
                    slFixedTime = slFixedTime & ":00"
                Else
                    If Len(slFixedTime) <= 4 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & Right$(slFixedTime, 2)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & Right$(slFixedTime, 4)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & Right$(slFixedTime, 2)
                    End If
                End If
            End If
        End If
    End If
    gConvertTime = slFixedTime
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
    slAnyDate = Trim$(slInpDate)
    If Trim$(slAnyDate) = "" Then
        gDateValue = 0
        Exit Function
    End If
    If Not gValidDate(slAnyDate) Then
        gDateValue = 0
        Exit Function
    End If
    slDate = gAdjYear(slAnyDate)
    gDateValue = DateValue(slDate)
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gAddStr                         *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Divide two strings (accuracy to *
'*                     same decimal place as the       *
'*                     larger decimal place of the     *
'*                     two numbers)                    *
'*                                                     *
'*                                                     *
'*******************************************************
Function gDivStr(slTop As String, slBottom As String) As String
'
'   sRet = gDivStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- numerator string
'       sStr2 (I)- denominator string
'       sRet (O)- sStr1 / sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilPos As Integer
    Dim ilNoDec As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim ilSign1 As Integer
    Dim ilSign2 As Integer
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slTop)
    slInput2 = Trim$(slBottom)
    clNum1 = Val(slInput1)
    If clNum1 < 0 Then
    clNum1 = -clNum1
    ilSign1 = -1
    Else
    ilSign1 = 1
    End If
    ilPos = InStr(slInput1, ".")
    If ilPos = 0 Then
    ilNoDec1 = 0
    Else
    ilNoDec1 = Len(slInput1) - ilPos
    End If
    clNum2 = Val(slInput2)
    If clNum2 < 0 Then
    clNum2 = -clNum2
    ilSign2 = -1
    Else
    ilSign2 = 1
    End If
    ilPos = InStr(slInput2, ".")
    If ilPos = 0 Then
    ilNoDec2 = 0
    Else
    ilNoDec2 = Len(slInput2) - ilPos
    End If
    If clNum2 <> 0 Then
    clNum = clNum1 / clNum2
    Else
    clNum = 0
    End If
    If (ilNoDec1 > 0) Or (ilNoDec2 > 0) Then
    clAdj = 0.5
    If ilNoDec1 > ilNoDec2 Then
        ilNoDec = ilNoDec1
    Else
        ilNoDec = ilNoDec2
    End If
    For ilLoop = 1 To ilNoDec Step 1
        clAdj = clAdj / 10@
    Next ilLoop
    If clNum >= 0 Then
        clNum = clNum + clAdj
    Else
        clNum = clNum - clAdj
    End If
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        slWhole = Left$(slStr, ilPos - 1)
        slDec = Right$(slStr, Len(slStr) - ilPos)
        If Len(slDec) > ilNoDec Then
        slDec = Left$(slDec, ilNoDec)
        End If
        Do While (Len(slDec) < ilNoDec)
        slDec = slDec & "0"
        Loop
        If ilSign1 * ilSign2 > 0 Then
        gDivStr = slWhole & "." & slDec
        Else
        gDivStr = "-" & slWhole & "." & slDec
        End If
    Else
        slStr = slStr & "."
        For ilLoop = 1 To ilNoDec Step 1
        slStr = slStr & "0"
        Next ilLoop
        If ilSign1 * ilSign2 > 0 Then
        gDivStr = slStr
        Else
        gDivStr = "-" & slStr
        End If
    End If
    Else
    clNum = clNum + 0.5
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        If ilSign1 * ilSign2 > 0 Then
        gDivStr = Left$(slStr, ilPos - 1)
        Else
        gDivStr = "-" & Left$(slStr, ilPos - 1)
        End If
    Else
        If ilSign1 * ilSign2 > 0 Then
        gDivStr = slStr
        Else
        gDivStr = "-" & slStr
        End If
    End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gExtNoRec                       *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute max # records for       *
'*                     btrieve extend operation        *
'*                                                     *
'*            Formula: # Rec = 60000/(6+RecSize)       *
'*                     6= description size added to    *
'*                        each record extracted and    *
'*                        into the return buffer       *
'*                                                     *
'*******************************************************
Function gExtNoRec(ilRecSize As Integer) As Long
    gExtNoRec = 8000 \ (6 + ilRecSize)  'Change 60000 to 8000
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gLongToInt                     *
'*                                                     *
'*             Created:3/25/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert Long to unsigned integer*
'*                     This is required to handle      *
'*                     unsigned integers (subtract     *
'*                     65536 from the value) to DLL    *
'*                                                     *
'*******************************************************
Function gLongToUnsignInt(llNumber As Long) As Integer
    If llNumber <= &H7FFF Then  'Test if high order bit is on
        gLongToUnsignInt = CInt(llNumber)
    ElseIf llNumber <= &HFFFF Then  'Test for overflow (number to large)
        gLongToUnsignInt = CInt(llNumber - &H10000)
    Else    'Overflow
        gLongToUnsignInt = 0
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gMulStr                         *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Multiply two strings            *
'*                     Round result up (accuracy to    *
'*                     same decimal place as the       *
'*                     larger decimal place of the two *
'*                     numbers)                        *
'*                                                     *
'*******************************************************
Function gMulStr(slMultStr1 As String, slMultStr2 As String) As String
'
'   sRet = gMulStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- First string
'       sStr2 (I)- Second string
'       sRet (O)- sStr1 * sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilPos As Integer
    Dim ilNoDec As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim ilSign1 As Integer
    Dim ilSign2 As Integer
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slMultStr1)
    slInput2 = Trim$(slMultStr2)
    clNum1 = Val(slInput1)
    If clNum1 < 0 Then
    clNum1 = -clNum1
    ilSign1 = -1
    Else
    ilSign1 = 1
    End If
    ilPos = InStr(slInput1, ".")
    If ilPos = 0 Then
    ilNoDec1 = 0
    Else
    ilNoDec1 = Len(slInput1) - ilPos
    End If
    clNum2 = Val(slInput2)
    If clNum2 < 0 Then
    clNum2 = -clNum2
    ilSign2 = -1
    Else
    ilSign2 = 1
    End If
    ilPos = InStr(slInput2, ".")
    If ilPos = 0 Then
    ilNoDec2 = 0
    Else
    ilNoDec2 = Len(slInput2) - ilPos
    End If
    If (ilNoDec1 > 0) Or (ilNoDec2 > 0) Then
    clNum = (10 * clNum1) * clNum2  'multiple by 10 to get 5 places
    clAdj = 0.5
    If ilNoDec1 > ilNoDec2 Then
        ilNoDec = ilNoDec1
    Else
        ilNoDec = ilNoDec2
    End If
    For ilLoop = 1 To ilNoDec - 1 Step 1
        clAdj = clAdj / 10@
    Next ilLoop
    If clNum >= 0 Then
        clNum = clNum + clAdj
    Else
        clNum = clNum - clAdj
    End If
    clNum = clNum / 10
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        slWhole = Left$(slStr, ilPos - 1)
        slDec = Right$(slStr, Len(slStr) - ilPos)
        If Len(slDec) > ilNoDec Then
        slDec = Left$(slDec, ilNoDec)
        End If
        Do While (Len(slDec) < ilNoDec)
        slDec = slDec & "0"
        Loop
        If ilSign1 * ilSign2 > 0 Then
        gMulStr = slWhole & "." & slDec
        Else
        gMulStr = "-" & slWhole & "." & slDec
        End If
    Else
        slStr = slStr & "."
        For ilLoop = 1 To ilNoDec Step 1
        slStr = slStr & "0"
        Next ilLoop
        If ilSign1 * ilSign2 > 0 Then
        gMulStr = slStr
        Else
        gMulStr = "-" & slStr
        End If
    End If
    Else
    clNum = clNum1 * clNum2 '+ .5
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        If ilSign1 * ilSign2 > 0 Then
        gMulStr = Left$(slStr, ilPos - 1)
        Else
        gMulStr = "-" & Left$(slStr, ilPos - 1)
        End If
    Else
        If ilSign1 * ilSign2 > 0 Then
        gMulStr = slStr
        Else
        gMulStr = "-" & slStr
        End If
    End If
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
'*      Procedure Name:gParseItem                      *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Function gParseItem(ByVal slInputStr As String, ByVal ilItemNo As Integer, slDelimiter As String, slOutputStr As String) As Integer
'
'   iRet = gParseItem(slInputStr, ilItemNo, slDelimiter, slOutStr)
'   Where:
'       slInputStr (I)-string from which to obtain substring
'       ilItemNo (I)-substring number to obtain (first string is Item number 1)
'       slDelimiter (I)-delimiter string or character between strings
'       slOutStr (O)-substring
'       iRet =  TRUE if substring found, FALSE if substring not found
'
    Dim ilEndPos As Integer  'Enp position of substring within sInputStr
    Dim ilStartPos As Integer    'Start position of each substring
    Dim ilIndex As Integer   'For loop parameter
    Dim ilLen As Integer 'Length of string to be parsed
    Dim ilDelimiterLen As Integer    'Delimiter length
    ilLen = Len(slInputStr)   'Obtain length so start position will not exceed length
    ilDelimiterLen = Len(slDelimiter)
    ilStartPos = 1   'Initialize start position
    For ilIndex = 1 To ilItemNo - 1 Step 1    'Loop until at starting position of substring to be found
        ilStartPos = InStr(ilStartPos, slInputStr, slDelimiter, 1) + ilDelimiterLen
        If (ilStartPos = ilDelimiterLen) Or (ilStartPos > ilLen) Then
            gParseItem = CP_MSG_PARSE
            Exit Function
        End If
    Next ilIndex
    ilEndPos = InStr(ilStartPos, slInputStr, slDelimiter, 1)   'Position end to end of substring plus 1 (start of delimiter position)
    If (ilEndPos = 0) Then   'No end delimiter-copy reTrafficing string
        slOutputStr = Trim$(Mid$(slInputStr, ilStartPos))
        gParseItem = CP_MSG_NONE
        Exit Function
    End If
    slOutputStr = Trim$(Mid$(slInputStr, ilStartPos, ilEndPos - ilStartPos))
    gParseItem = CP_MSG_NONE
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPDNToStr                       *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Packed Decimal Number to String *
'*                                                     *
'*******************************************************
Sub gPDNToStr(slInPDN As String, ilNoDecPlaces As Integer, slOutValue As String)
'
'   gPDNToStr slPDN, ilDP, slStr
'   Where:
'       ilPDN (I)- Packed string to be converted
'       ilDP (I)- Number of decimal places
'       slStr (O) - Unpacked String
'
    Dim slNumber As String  'String to be converted (left to right)
    Dim ilByte As Integer   'Extracted byte from string to be converted
    Dim ilLoop As Integer
    Dim slPDN As String
    Dim slOut As String
    Dim slSign As String
    Dim slLeft As String
    Dim slRight As String
    Dim ilLen As Integer
    slPDN = Trim$(slInPDN)
    slNumber = slPDN
    slSign = ""
    slOut = ""
    For ilLoop = 1 To Len(slPDN) Step 1
        On Error GoTo gPDNToStrErr
        ilByte = Asc(slNumber)  'Obtain left most byte
        On Error GoTo 0
        slNumber = Mid$(slNumber, 2)    'Remove the left most byte
        If ilLoop = Len(slPDN) Then    'Last byte
            slOut = slOut & Trim$(Str$(ilByte \ 16))
            If (ilByte And 15) = 13 Then    'Negative number
                slSign = "-"
            End If
            Exit For
        Else
            slOut = slOut + Trim$(Str$(ilByte \ 16)) + Trim$(Str$(ilByte And 15))
        End If
    Next ilLoop
    If ilNoDecPlaces > 0 Then
        slLeft = Left$(slOut, Len(slOut) - ilNoDecPlaces)
        slRight = Right$(slOut, ilNoDecPlaces)
    Else
        slLeft = slOut
    End If
    ilLen = Len(slLeft) - 1
    For ilLoop = 1 To ilLen Step 1
        If Left$(slLeft, 1) <> "0" Then
            Exit For
        End If
        slLeft = Right$(slLeft, Len(slLeft) - 1)
    Next ilLoop
    If ilNoDecPlaces > 0 Then
        slOut = slLeft & "." & slRight
    Else
        slOut = slLeft
    End If
    slOutValue = Trim$(slSign & slOut)
    Exit Sub
gPDNToStrErr:
    On Error GoTo 0
    slOutValue = ""
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gStrToPDN                       *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:String to Packed Decimal Number *
'*                                                     *
'*******************************************************
Sub gStrToPDN(slInValue As String, ilNoDecPlaces As Integer, ilPDNLen As Integer, slOutPDN As String)
'
'   gStrToPDN slValue, iNoDP, ilLen, slStr
'   Where:
'       slValue (I)- String value to be packed (it can contain , $ - + .)
'       iNoDP (I)- Number of decimal places
'       ilLen (I)- Length of output string (number of bytes-two digits per byte, last byte contains one digit and sign)
'       slStr (O) - String to contain pack number
'
    Dim slNumExtra As String      'Number with extra characters
    Dim slNumStrip As String    'Number with characters removed
    Dim slLeftChar As String * 1
    Dim ilLoop As Integer
    Dim ilNoDp As Integer   'Number of decimal places found
    Dim ilCount As Integer  'Count decimal places
    Dim llNumStrip As Long
    Dim ilPosSign As Integer
    Dim ilLen As Integer
    slNumExtra = Trim$(slInValue)
    ilLen = Len(slNumExtra)
    'Remove all characters and count number of places after decimal
    ilNoDp = 0
    ilCount = False
    ilPosSign = True
    slNumStrip = ""
    For ilLoop = 1 To ilLen Step 1
        slLeftChar = Left$(slNumExtra, 1)
        slNumExtra = Mid$(slNumExtra, 2)
        If slLeftChar >= "0" And slLeftChar <= "9" Then
            slNumStrip = slNumStrip & slLeftChar
            If ilCount Then
                ilNoDp = ilNoDp + 1
            End If
        Else
            If slLeftChar = "." Then
                ilCount = True
            End If
            If slLeftChar = "-" Then
                ilPosSign = False
            End If
        End If
        If slNumExtra = "" Then
            Exit For
        End If
    Next ilLoop
    If ilNoDp < ilNoDecPlaces Then  'Add zeros
        For ilLoop = ilNoDp To ilNoDecPlaces - 1 Step 1
            slNumStrip = slNumStrip & "0"
        Next ilLoop
    End If
    If ilNoDp > ilNoDecPlaces Then  'Remove extra digits
        slNumStrip = Left$(slNumStrip, Len(slNumStrip) - (ilNoDp - ilNoDecPlaces))
    End If
    If Not ilPosSign Then
        slNumStrip = "-" & slNumStrip
    End If
    mMakePDN slNumStrip, ilPDNLen, slOutPDN
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gSubStr                         *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Subtract two strings            *
'*                     (accuracy to same decimal place *
'*                     as the larger decimal place     *
'*                     of the two numbers)             *
'*                                                     *
'*******************************************************
Function gSubStr(slSubStr1 As String, slSubStr2 As String) As String
'
'   sRet = gSubStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- String from which to subtract (minuend)
'       sStr2 (I)- String to subtract (subtrahend)
'       sRet (O)- sStr1 - sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilPos As Integer
    Dim ilNoDec As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slSubStr1)
    slInput2 = Trim$(slSubStr2)
    clNum1 = Val(slInput1)
    ilPos = InStr(slInput1, ".")
    If ilPos = 0 Then
    ilNoDec1 = 0
    Else
    ilNoDec1 = Len(slInput1) - ilPos
    End If
    clNum2 = Val(slInput2)
    ilPos = InStr(slInput2, ".")
    If ilPos = 0 Then
    ilNoDec2 = 0
    Else
    ilNoDec2 = Len(slInput2) - ilPos
    End If
    clNum = clNum1 - clNum2
    If (ilNoDec1 > 0) Or (ilNoDec2 > 0) Then
    clAdj = 0.5
    If ilNoDec1 > ilNoDec2 Then
        ilNoDec = ilNoDec1
    Else
        ilNoDec = ilNoDec2
    End If
    For ilLoop = 1 To ilNoDec Step 1
        clAdj = clAdj / 10@
    Next ilLoop
    If clNum >= 0 Then
        clNum = clNum + clAdj
    Else
        clNum = clNum - clAdj
    End If
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        slWhole = Left$(slStr, ilPos - 1)
        slDec = Right$(slStr, Len(slStr) - ilPos)
        If Len(slDec) > ilNoDec Then
        slDec = Left$(slDec, ilNoDec)
        End If
        Do While (Len(slDec) < ilNoDec)
        slDec = slDec & "0"
        Loop
        gSubStr = slWhole & "." & slDec
    Else
        slStr = slStr & "."
        For ilLoop = 1 To ilNoDec Step 1
        slStr = slStr & "0"
        Next ilLoop
        gSubStr = slStr
    End If
    Else
'        clNum = clNum + .5
    slStr = Trim$(Str$(clNum))
    ilPos = InStr(slStr, ".")
    If ilPos <> 0 Then
        gSubStr = Left$(slStr, ilPos - 1)
    Else
        gSubStr = slStr
    End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gAddStr                         *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Divide two strings (accuracy to *
'*                     same decimal place as the       *
'*                     larger decimal place of the     *
'*                     two numbers)                    *
'*                                                     *
'*                                                     *
'*******************************************************
Function gTDivStr(slTop As String, slBottom As String) As String
'
'   sRet = gTDivStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- numerator string
'       sStr2 (I)- denominator string
'       sRet (O)- sStr1 \ sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilPos As Integer
    Dim ilNoDec As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim ilSign1 As Integer
    Dim ilSign2 As Integer
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slTop)
    slInput2 = Trim$(slBottom)
    clNum1 = Val(slInput1)
    If clNum1 < 0 Then
        clNum1 = -clNum1
        ilSign1 = -1
    Else
        ilSign1 = 1
    End If
    ilPos = InStr(slInput1, ".")
    If ilPos = 0 Then
        ilNoDec1 = 0
    Else
        ilNoDec1 = Len(slInput1) - ilPos
    End If
    clNum2 = Val(slInput2)
    If clNum2 < 0 Then
        clNum2 = -clNum2
        ilSign2 = -1
    Else
        ilSign2 = 1
    End If
    ilPos = InStr(slInput2, ".")
    If ilPos = 0 Then
        ilNoDec2 = 0
    Else
        ilNoDec2 = Len(slInput2) - ilPos
    End If
    If clNum2 <> 0 Then
        clNum = clNum1 \ clNum2
    Else
        clNum = 0
    End If
    If (ilNoDec1 > 0) Or (ilNoDec2 > 0) Then
        clAdj = 0#  '.5
        If ilNoDec1 > ilNoDec2 Then
            ilNoDec = ilNoDec1
        Else
            ilNoDec = ilNoDec2
        End If
        For ilLoop = 1 To ilNoDec Step 1
            clAdj = clAdj / 10@
        Next ilLoop
        If clNum >= 0 Then
            clNum = clNum + clAdj
        Else
            clNum = clNum - clAdj
        End If
        slStr = Trim$(Str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = Right$(slStr, Len(slStr) - ilPos)
            If Len(slDec) > ilNoDec Then
                slDec = Left$(slDec, ilNoDec)
            End If
            Do While (Len(slDec) < ilNoDec)
                slDec = slDec & "0"
            Loop
            If ilSign1 * ilSign2 > 0 Then
                gTDivStr = slWhole & "." & slDec
            Else
                gTDivStr = "-" & slWhole & "." & slDec
            End If
        Else
            slStr = slStr & "."
            For ilLoop = 1 To ilNoDec Step 1
                slStr = slStr & "0"
            Next ilLoop
            If ilSign1 * ilSign2 > 0 Then
                gTDivStr = slStr
            Else
                gTDivStr = "-" & slStr
            End If
        End If
    Else
        clNum = clNum '+ .5
        slStr = Trim$(Str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            If ilSign1 * ilSign2 > 0 Then
                gTDivStr = Left$(slStr, ilPos - 1)
            Else
                gTDivStr = "-" & Left$(slStr, ilPos - 1)
            End If
        Else
            If ilSign1 * ilSign2 > 0 Then
                gTDivStr = slStr
            Else
                gTDivStr = "-" & slStr
            End If
        End If
    End If
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
        slUnTime = slUnTime & Right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00:00")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & Right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":0A")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & Right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":0P")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & Right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00A")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & Right$(slTime, 2)
        gUnformatTime = slUnTime
        Exit Function
    End If
    ilPos = InStr(slUnTime, ":00P")
    If ilPos <> 0 Then
        slUnTime = Left$(slTime, ilPos - 1)
        slUnTime = slUnTime & Right$(slTime, 2)
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
    dlDateSerial = DateSerial(ilYear, ilMn, ilDy)
    slDate = Trim$(Format$(dlDateSerial, "m/d/yy"))
'    slDate = Trim$(Str$(ilMn)) & "/" & Trim$(Str$(ilDy)) & "/" & Trim$(Str$(ilYear))
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
'*      Procedure Name:gUnsignIntToLong                *
'*                                                     *
'*             Created:3/25/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert Integer (unsigned) to   *
'*                     long. This is required to handle*
'*                     unsigned integers with DLL      *
'*                                                     *
'*******************************************************
Function gUnsignIntToLong(ilNumber As Integer) As Long
    gUnsignIntToLong = (CLng(ilNumber) And &HFFFF&) 'Remove high order 16 bits
End Function
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
    slAnyDate = Trim$(slInpDate)
    If Len(slAnyDate) <> 0 Then
        slDate = gAdjYear(slAnyDate)
        On Error GoTo gValidDateErr
        ilDate = Day(slDate)
        ilDate = Month(slDate)
        ilDate = Year(slDate)
        If (ilDate < 1970) Or (ilDate > 2069) Then
            gValidDate = False
            Exit Function
        End If
    Else
        gValidDate = False
        Exit Function
    End If
    On Error GoTo 0
    gValidDate = True
    Exit Function
gValidDateErr:
    On Error GoTo 0
    gValidDate = False
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
        gValidLength = True
        Exit Function
    End If
    slHour = ""
    slMin = ""
    slSec = ""
    slLen = Trim$(slLength)
    ilPos = InStr(1, slLen, "::")
    If ilPos > 0 Then
        gValidLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hm", 1)
    If ilPos > 0 Then
        gValidLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hs", 1)
    If ilPos > 0 Then
        gValidLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "ms", 1)
    If ilPos > 0 Then
        gValidLength = False
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
            gValidLength = False
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
            gValidLength = False
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
        llHour = CLng(slHour) 'Val(slHour)
    Else
        llHour = 0
    End If
    If slMin <> "" Then
        ilMin = CLng(slMin) 'Val(slMin)
    Else
        ilMin = 0
    End If
    If slSec <> "" Then
        ilSec = CLng(slSec) 'Val(slSec)
    Else
        ilSec = 0
    End If
    If (llHour < 0) Or (llHour > 24) Then
        gValidLength = False
        Exit Function
    End If
    If (llHour <> 0) Then
        If (ilMin < 0) Or (ilMin > 59) Then
            gValidLength = False
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
            gValidLength = False
            Exit Function
        End If
    End If
    If llHour * 3600& + CLng(ilMin) * 60& + CLng(ilSec) > 86400 Then
        gValidLength = False
        Exit Function
    End If
    gValidLength = True
    Exit Function
gValidLengthErr:
    On Error GoTo 0
    gValidLength = False
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
    gValidTime = True
    Exit Function
gValidTimeErr:
    On Error GoTo 0
    gValidTime = False
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
    Dim slDate As String
    Dim ilWeekDay As Integer
    
    ilWeekDay = Weekday(llAnyDate) - 2
    If ilWeekDay < 0 Then
        ilWeekDay = 6
    End If
    gWeekDayLong = ilWeekDay
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakePDN                        *
'*                                                     *
'*             Created:5/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:String to Packed Decimal Number *
'*                                                     *
'*******************************************************
Private Sub mMakePDN(slInStr As String, ilPDNLen As Integer, slOutPDN As String)
'
'   mMakeToPDN slInValue, ilLen, slOutStr
'   Where:
'       slInValue (I)- string value to be packed
'       ilLen (I)- Length of output string (number of bytes-two digits per byte, last byte contains one digit and sign)
'       slOutStr (O) - String to contain pack number
'
    Dim ilRem As Integer    'Remainder (value to be packed)
    Dim ilHRem As Integer   'High part of nibble to be packed
    Dim ilLRem As Integer   'Low part of nibble to be packed
    Dim ilNibble As Integer 'Nibble (1=High part, 2 = Low part)
    Dim slValue As String   'Value as string to be converted
    Dim ilPos As Integer    'Character position to be converted
    Dim ilZero As Integer   'ANSI code of zero: Asc("0")
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilIndex As Integer  'Location within string to store packed value
    Dim slRem As String 'Packed value for byte
    If Left$(slInStr, 1) = "-" Then
        slValue = Mid$(slInStr, 2)
        ilLRem = 13 '"D"=Negative number
    Else
        If Left$(slInStr, 1) = "+" Then
            slValue = Mid$(slInStr, 2)
        Else
            slValue = slInStr
        End If
        ilLRem = 15 '"F"=Positive number
    End If
    'Initialize string to all zero's
    slOutPDN = ""
    For ilLoop = 1 To ilPDNLen Step 1
        slOutPDN = slOutPDN + Chr$(0)
    Next ilLoop
    'Pack number
    ilNibble = 1
    ilIndex = ilPDNLen
    ilPos = Len(slValue)
    If ilPos <= 0 Then
        ilHRem = 0
        slRem = Chr$(ilHRem * 16 + ilLRem)
        Mid$(slOutPDN, ilIndex, 1) = slRem
        Exit Sub
    End If
    ilZero = Asc("0")
    For ilLoop = 1 To 2 * ilPDNLen - 1 Step 1
        ilRem = Asc(Mid$(slValue, ilPos, 1)) - ilZero
        ilPos = ilPos - 1
        If ilNibble = 1 Then
            ilHRem = ilRem
        Else
            ilLRem = ilRem
            ilHRem = 0
            ilNibble = 0
            ilIndex = ilIndex - 1
        End If
        slRem = Chr$(ilHRem * 16 + ilLRem)
        Mid$(slOutPDN, ilIndex, 1) = slRem
        ilNibble = ilNibble + 1
        If ilPos <= 0 Then
            Exit For
        End If
    Next ilLoop
End Sub
