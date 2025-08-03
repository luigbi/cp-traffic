Attribute VB_Name = "MATHSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of MATHSUBS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: MathSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Math subs and functions
Option Explicit
Option Compare Text

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

'slAns = Space$(30)
'ilLen = AddStr(slAddStr1, slAddStr2, slAns)
'gAddStr = Left$(slAns, ilLen)
'Exit Function
    'Code below moved to dll
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
        'For ilLoop = 1 To ilNoDec Step 1
        '    clAdj = clAdj / 10@
        'Next ilLoop
        'If clNum >= 0 Then
        '    clNum = clNum + clAdj
        'Else
        '    clNum = clNum - clAdj
        'End If
        slStr = Trim$(str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = right$(slStr, Len(slStr) - ilPos)
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
        slStr = Trim$(str$(clNum))
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
'*      Procedure Name:gCompAbsNumberStr               *
'*                                                     *
'*             Created:10/07/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compare two number strings.     *
'*                     The strings are converted to    *
'*                     numeric values, then compared.  *
'*                     Negative number are converted to*
'*                     positive number prior to compare*
'*                                                     *
'*******************************************************
Function gCompAbsNumberStr(slCompStr1 As String, slCompStr2 As String) As Integer
'
'   iRet = gCompAbsNumberStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- String to be compared
'       sStr2 (I)- String to compare
'       iRet (O)- 0 if sStr1=sStr2; -1 if sStr1 < sStr2; 1 If sStr1 > sStr2
'

    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim slInput1 As String
    Dim slInput2 As String

    slInput1 = Trim$(slCompStr1)
    slInput2 = Trim$(slCompStr2)
    clNum1 = Val(slInput1)
    If clNum1 < 0 Then
        clNum1 = -clNum1
    End If
    clNum2 = Val(slInput2)
    If clNum2 < 0 Then
        clNum2 = -clNum2
    End If
    If clNum1 = clNum2 Then
        gCompAbsNumberStr = 0
    ElseIf clNum1 > clNum2 Then
        gCompAbsNumberStr = 1
    Else
        gCompAbsNumberStr = -1
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCompNumberStr                  *
'*                                                     *
'*             Created:10/07/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compare two number strings.     *
'*                     The strings are converted to    *
'*                     numeric values, then compared.  *
'*                                                     *
'*******************************************************
Function gCompNumberStr(slCompStr1 As String, slCompStr2 As String) As Integer
'
'   iRet = gCompNumberStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- String to be compared
'       sStr2 (I)- String to compare
'       iRet (O)- 0 if sStr1=sStr2; -1 if sStr1 < sStr2; 1 If sStr1 > sStr2
'

    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim slInput1 As String
    Dim slInput2 As String

    slInput1 = Trim$(slCompStr1)
    slInput2 = Trim$(slCompStr2)
    clNum1 = Val(slInput1)
    'If clNum1 < 0 Then
    '    clNum1 = -clNum1
    'End If
    clNum2 = Val(slInput2)
    'If clNum2 < 0 Then
    '    clNum2 = -clNum2
    'End If
    If clNum1 = clNum2 Then
        gCompNumberStr = 0
    ElseIf clNum1 > clNum2 Then
        gCompNumberStr = 1
    Else
        gCompNumberStr = -1
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
Function gDivStr(slTop As String, slBottom As String) As String
'
'   sRet = gDivStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- numerator string
'       sStr2 (I)- denominator string
'       sRet (O)- sStr1 / sStr2
'

'slAns = Space$(30)
'ilLen = DivStr(slTop, slBottom, slAns)
'gDivStr = Left$(slAns, ilLen)
'Exit Function
    'Code below move to dll
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
        slStr = Trim$(str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = right$(slStr, Len(slStr) - ilPos)
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
        slStr = Trim$(str$(clNum))
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
'*      Procedure Name:gIntToRoman                     *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert arabic number (integer) *
'*                     to roman number                 *
'*                     I = 1; V=5; X=10; L=50; C=100   *
'*                     D=500; M=1000                   *
'*                                                     *
'*******************************************************
Function gIntToRoman(ilArabic As Integer) As String
'
'  slNumber = gIntToRoman(ilNumber)
'   Where:
'       ilNumber (I)- Number to be converted (max allowed is 1000)
'       slNumber (O)- Number as a roman number
'

'slAns = Space$(30)
'ilLen = IntToRoman(ilArabic, slAns)
'gIntToRoman = Left$(slAns, ilLen)
'Exit Function
    'Code below move to dll
    Dim ilNumber As Integer
    Dim ilRemainder As Integer
    Dim slRoman As String
    Dim sl1To3 As String
    Dim sl5 As String
    Dim sl10 As String
    sl1To3 = "I"
    sl5 = "V"
    sl10 = "X"
    slRoman = ""
    ilNumber = ilArabic
    Do While ilNumber > 0
        ilRemainder = ilNumber Mod 10
        Select Case ilRemainder
            Case 1
                slRoman = sl1To3 & slRoman
            Case 2
                slRoman = sl1To3 & sl1To3 & slRoman
            Case 3
                slRoman = sl1To3 & sl1To3 & sl1To3 & slRoman
            Case 4
                slRoman = sl1To3 & sl5 & slRoman
            Case 5
                slRoman = sl5 & slRoman
            Case 6
                slRoman = sl5 & sl1To3 & slRoman
            Case 7
                slRoman = sl5 & sl1To3 & sl1To3 & slRoman
            Case 8
                slRoman = sl5 & sl1To3 & sl1To3 & sl1To3 & slRoman
            Case 9
                slRoman = sl1To3 & sl10 & slRoman
        End Select
        ilNumber = ilNumber \ 10
        If sl5 = "V" Then
            sl1To3 = "X"
            sl5 = "L"
            sl10 = "C"
        ElseIf sl5 = "L" Then
            sl1To3 = "C"
            sl5 = "D"
            sl10 = "M"
        End If
    Loop
    gIntToRoman = slRoman
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

'slAns = Space$(30)
'ilLen = MulStr(slMultStr1, slMultStr2, slAns)
'gMulStr = Left$(slAns, ilLen)
'Exit Function
    'Code below moved to dll
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
        slStr = Trim$(str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = right$(slStr, Len(slStr) - ilPos)
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
        slStr = Trim$(str$(clNum))
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
'*      Procedure Name:gRoundStr                       *
'*                                                     *
'*             Created:7/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Round a string                  *
'*                                                     *
'*******************************************************
Function gRoundStr(slInputStr As String, slRoundStr As String, ilNoDecRd As Integer) As String
'
'   sRet = gRoundStr(sStr1, sRd, iNoDec)
'   Where:
'       sStr1 (I)- Value to be rounded
'       sRd (I)- Accuracy of value (to pennies .01, to dollar 1, to ten dollars 10)
'       iNoDec (I)- Number of decimal places to retain
'       sRet (O)- sStr1 rounded
'

'slAns = Space$(30)
'ilLen = RoundStr(slInputStr, slRoundStr, ilNoDecRd, slAns)
'gRoundStr = Left$(slAns, ilLen)
'Exit Function
    'Code below moved to DLL
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim clNum As Currency
    Dim clAdj As Currency
    Dim ilNoDec1 As Integer
    Dim ilNoDec2 As Integer
    Dim ilNoDec As Integer
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWhole As String
    Dim slDec As String
    Dim llCount As Long
    Dim ilSign1 As Integer
    Dim slInput As String
    Dim slRound As String
    Dim cl10000 As Currency
    Dim clTemp As Currency

    cl10000 = 10000@
    slInput = Trim$(slInputStr)
    slRound = Trim$(slRoundStr)
    clNum1 = Val(slInput)
    If clNum1 < 0 Then
        ilSign1 = -1
        clNum1 = -clNum1
    Else
        ilSign1 = 1
    End If
    clNum2 = Val(slRound)
    If clNum2 <> 0 Then
        ilPos = InStr(slInput, ".")
        If ilPos = 0 Then
            ilNoDec1 = 0
        Else
            ilNoDec1 = Len(slInput) - ilPos
        End If
        ilPos = InStr(slRound, ".")
        If ilPos = 0 Then
            ilNoDec2 = 0
        Else
            ilNoDec2 = Len(slRound) - ilPos
        End If
        If ilNoDec1 > ilNoDec2 Then
            ilNoDec = ilNoDec1
        Else
            ilNoDec = ilNoDec2
        End If
        clAdj = 1
        For ilLoop = 1 To ilNoDec Step 1
            clAdj = 10 * clAdj
        Next ilLoop
        ' A Mod B =>  A - B*(A/B) => in Currency A - B*(10000*((A/B)/10000))
        'clNum = ((clAdj * clNum1) Mod (clAdj * clNum2)) / clAdj
        clTemp = ((clAdj * clNum1) / (clAdj * clNum2)) / cl10000
        clNum = ((clAdj * clNum1) - ((cl10000 * (clTemp)) * (clAdj * clNum2))) / clAdj
        '4-17-03Remove:  if the number to be round was less then the rounding factor, then set the number to the rounding factor.
        'If clNum1 < clNum2 Then
        '    If clNum1 <> 0 Then
        '        clNum1 = clNum2
        '    End If
        'Else
            If clNum + clNum >= clNum2 Then
                clNum1 = clNum1 + clNum2 - clNum    'round up
            Else
                clNum1 = clNum1 - clNum 'Round down
            End If
        'End If
        llCount = 0
        'clNum = (clAdj * clNum1) Mod (clAdj * clNum2)
        clTemp = ((clAdj * clNum1) / (clAdj * clNum2)) / cl10000
        clNum = (clAdj * clNum1) - ((cl10000 * (clTemp)) * (clAdj * clNum2))
        Do While clNum <> 0
            llCount = llCount + 1
            If llCount > clAdj * clNum2 Then
                Exit Do
            End If
            If clNum > (clAdj * clNum2) / 2 Then
                clNum1 = clNum1 + 1 / clAdj
            Else
                clNum1 = clNum1 - 1 / clAdj
            End If
            'clNum = (clAdj * clNum1) Mod (clAdj * clNum2)
            clTemp = ((clAdj * clNum1) / (clAdj * clNum2)) / cl10000
            clNum = (clAdj * clNum1) - ((cl10000 * (clTemp)) * (clAdj * clNum2))
        Loop
    End If
    slStr = Trim$(str$(clNum1))
    If ilNoDecRd > 0 Then
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = right$(slStr, Len(slStr) - ilPos)
            If Len(slDec) > ilNoDecRd Then
                slDec = Left$(slDec, ilNoDecRd)
            End If
            Do While (Len(slDec) < ilNoDecRd)
                slDec = slDec & "0"
            Loop
            If ilSign1 < 0 Then
                gRoundStr = "-" & slWhole & "." & slDec
            Else
                gRoundStr = slWhole & "." & slDec
            End If
        Else
            slStr = slStr & "."
            For ilLoop = 1 To ilNoDecRd Step 1
                slStr = slStr & "0"
            Next ilLoop
            If ilSign1 < 0 Then
                gRoundStr = "-" & slStr
            Else
                gRoundStr = slStr
            End If
        End If
    Else
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            If ilSign1 < 0 Then
                gRoundStr = "-" & Left$(slStr, ilPos - 1)
            Else
                gRoundStr = Left$(slStr, ilPos - 1)
            End If
        Else
            If ilSign1 < 0 Then
                gRoundStr = "-" & slStr
            Else
                gRoundStr = slStr
            End If
        End If
    End If
End Function

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

'slAns = Space$(30)
'ilLen = SubStr(slSubStr1, slSubStr2, slAns)
'gSubStr = Left$(slAns, ilLen)
'Exit Function
    'Code below moved to DLL
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
        slStr = Trim$(str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            slWhole = Left$(slStr, ilPos - 1)
            slDec = right$(slStr, Len(slStr) - ilPos)
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
        slStr = Trim$(str$(clNum))
        ilPos = InStr(slStr, ".")
        If ilPos <> 0 Then
            gSubStr = Left$(slStr, ilPos - 1)
        Else
            gSubStr = slStr
        End If
    End If
End Function


