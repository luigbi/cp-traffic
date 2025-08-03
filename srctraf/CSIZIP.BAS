Attribute VB_Name = "CSIZIPSubs"
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

Public sgExePath As String
Public sgDBPath As String

Public fgListHTArial825 As Single
Public fgListHtSerif825 As Single
Public igAlignCharWidth As Integer
Public igScrollBarWidth As Integer
Public fgBoxInsetX As Single
Public fgBoxGridH As Single

'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Type ZIPNAME
    sKey As String * 60     'Date, Time, Name
    sName As String * 40
    sDate As String * 10
    sTime As String * 10
    lSize As Long
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

'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjShowLen                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust length of string to fit *
'*                 within specified area minus margins *
'*                                                     *
'*******************************************************
Sub gAdjShowLen(Ctrl As Control, slShow As String, flWidth As Single)
'
'   gAdjShowLen Ctrl, sShow, BoxWidth
'   Where:
'       Ctrl (I)- Control where string will be painted
'       sShow (I/O)- String to be shown
'       ilBoxWidth (I)- width of control area
'

    Dim flPicWidth
    flPicWidth = flWidth - 2 * fgBoxInsetX
    Do While (Ctrl.TextWidth(slShow) > flPicWidth)
        slShow = left$(slShow, Len(slShow) - 1)
    Loop
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gParseItemFields                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Sub gParseItemFields(ByVal slInputStr As String, slDelimiter As String, slFields() As String)
'
'   gParseItemToArray slInputStr, slDelimiter, slFields()
'   Where:
'       slInputStr (I)-string from which to obtain substring
'       slFields() (O)- fields parsed from comma delimited string
'

    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
        slFields(ilIndex) = ""
    Next ilIndex
    ilFieldNo = 1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slInputStr))
        slChar = Mid$(slInputStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = slDelimiter Then
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    Exit Sub
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                ilFieldType = 0 'Quote field
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" Then
                    If ilIndex = Len(Trim$(slInputStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slInputStr, ilIndex + 1, 1)
                        If slNextChar = slDelimiter Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> slDelimiter Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintArea                      *
'*                                                     *
'*             Created:7/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear area                      *
'*                                                     *
'*******************************************************
Sub gPaintArea(pbcCtrl As Control, flCurrX As Single, flCurrY As Single, flWidth As Single, flHeight As Single, llPaintColor As Long)
'
'   gPaintArea pbcCtrl, lX, lY, lW, lH, lC
'   Where:
'       pbcCtrl (I)- Picture box control containing area to be painted
'       lX (I)- Left position of area to be painted
'       lY (I)- Top position of area to be painted
'       lH (I)- Height of area to be painted
'       lW (I)- Width of area to be painted
'       lC (I)- Color of area to be painted
'

    pbcCtrl.CurrentX = flCurrX
    pbcCtrl.CurrentY = flCurrY
    pbcCtrl.Line Step(0, 0)-Step(flWidth, flHeight), llPaintColor, BF
End Sub

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
