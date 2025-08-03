Attribute VB_Name = "EngrGenSubs"
'
' Release: 1.0
'
' Description:
'   This file contains the General declarations
Option Explicit

'ttp 5218 Dan M.
Public ogEmailer As CEmail
Type EMPLOYEEEMAILS
    Name As String
    eMail As String
End Type

Sub gSetListBoxHeight(lbcCtrl As ListBox, llMaxHeight As Long)
'
'  flHeight = gListBoxHeight (ilNoRows, ilMaxRows)
'   Where:
'       ilNoRows (I) - current number of items within the list box
'       ilMaxRows (I) - max number of list box items to be displayed
'       flHeight (O) - height of list box in twips
'
    '+30 because of line above and below
    Dim llRowHeight As Long
    Dim slStr As String
    Dim ilMaxRow As Integer
    
    If lbcCtrl.ListCount > 0 Then
        'Determine standard height, set to small number so that only one row
        'height will be set (15 + 15 + RowHeight; 15 for size of boundaries)
        'Typical result is 300 (15 + 15 + 270)
'        lbcCtrl.Height = 10
        llRowHeight = 15 * SendMessageByString(lbcCtrl.hwnd, LB_GETITEMHEIGHT, 0, slStr)
        ilMaxRow = llMaxHeight / llRowHeight
        If lbcCtrl.ListCount <= ilMaxRow Then
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
            lbcCtrl.Height = (llRowHeight) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
        Else
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
            lbcCtrl.Height = (llRowHeight) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
        End If
    End If
End Sub

Sub gProcessArrowKey(ilShift As Integer, ilKeyCode As Integer, lbcCtrl As Control, ilRetainState As Integer)
'
'   gProcessArrowKey Shift, KeyCode, lbcCtrl, imLbcArrowSetting
'   Where:
'       Shift (I)- Shift key state
'       KeyCode (I)- Key code
'       lbcCtrl (I)- list box control
'       ilLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilState As Integer
    
    If (ilShift And ALTMASK) > 0 Then
        lbcCtrl.Visible = Not lbcCtrl.Visible
    ElseIf (ilShift And SHIFTMASK) > 0 Then
    Else
        ilState = lbcCtrl.Visible
        If ilKeyCode = KEYUP Then    'Up arrow
            If lbcCtrl.ListIndex > 0 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex - 1
                If ilRetainState Then
                    lbcCtrl.Visible = ilState
                End If
            End If
        Else
            If lbcCtrl.ListIndex < lbcCtrl.ListCount - 1 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex + 1
                If ilRetainState Then
                    lbcCtrl.Visible = ilState
                End If
            End If
        End If
    End If
End Sub



Public Sub gShellAndWait(Frm As Form, ByVal slFilePath As String, ilWinStyle As Integer)
    Dim llProcess As Long
    Dim llReturn As Long
    
    llProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(slFilePath, ilWinStyle))
    If llProcess <> 0 Then
        'Frm.WindowState = vbMinimized
        Frm.Enabled = False
        'Frm.Visible = False
        Do
            GetExitCodeProcess llProcess, llReturn
            Sleep 50
            DoEvents
        Loop While llReturn = STILL_ACTIVE
        Frm.Enabled = True
        'If Frm = "Traffic" Then
        '    FrmWindowState = vbMaximized
        'Else
        '    Frm.WindowState = vbNormal
        'End If
        'Frm.Visible = True
    Else
        MsgBox "Unable to Shell to " & slFilePath, vbOKOnly, "Shell Error"
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gFileDateTime                   *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain file time stamp          *
'*                                                     *
'*******************************************************
Function gFileDateTime(slPathFile As String) As String
    Dim ilRet As Integer
    
    ilRet = 0
    On Error GoTo gFileDateTimeErr
    gFileDateTime = FileDateTime(slPathFile)
    If ilRet <> 0 Then
        gFileDateTime = Format$(Now, "ddddd") & " " & Format$(Now, "ttttt")
    End If
    On Error GoTo 0
    Exit Function
gFileDateTimeErr:
    ilRet = Err.Number
    Resume Next
End Function

Public Function gFixQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = "'" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixQuote = sOutStr
End Function
Public Function gFixDoubleQuote(sInStr As String) As String

    Dim sOutStr As String
    Dim sChar As String
    Dim iLoop As Integer
    
    sOutStr = ""
    If IsNull(sInStr) <> True Then
        For iLoop = 1 To Len(sInStr) Step 1
            sChar = Mid$(sInStr, iLoop, 1)
            If sChar = """" Then
                sOutStr = sOutStr & "''"
            Else
                sOutStr = sOutStr & sChar
            End If
        Next iLoop
    End If
    gFixDoubleQuote = sOutStr
End Function








'*******************************************************
'*                                                     *
'*      Procedure Name:gParseCDFields                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Parse comma delimited fields    *
'*                     Note:including quotes that are  *
'*                     enclosed within quotes          *
'*                     ""xxxxxxxx"","xxxxx",           *
'*                                                     *
'*******************************************************
Sub gParseCDFields(slCDStr As String, ilLower As Integer, slFields() As String)
'
'   gParseCDFields slCDStr, ilLower, slFields()
'   Where:
'       slCDStr(I)- Comma delinited string
'       ilLower(I)- True=Convert string fields characters to lower case (preceding character is A-Z)
'       slFields() (O)- fields parsed from comma delimited string
'
    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

'    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
'        slFields(ilIndex) = ""
'    Next ilIndex
    ReDim slFields(1 To 1) As String
    slFields(UBound(slFields)) = ""
    ilFieldNo = 1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = "," Then    'Comma was followed by a comma-blank field
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                    slFields(UBound(slFields)) = ""
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
                    If ilIndex = Len(Trim$(slCDStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slCDStr, ilIndex + 1, 1)
                        If slNextChar = "," Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    If (slFields(ilFieldNo) <> "") And ilLower Then
                        ilAscChar = Asc(UCase(right$(slFields(ilFieldNo), 1)))
                        If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then
                            slChar = LCase$(slChar)
                        End If
                    End If
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> "," Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub


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

Public Function gDayMap(slInDays As String) As String
    Dim slDays As String
    Dim ilDay As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim slStr As String
    
    slDays = Trim$(slInDays)
    If (InStr(1, slDays, "Y", vbTextCompare) > 0) Or (InStr(1, slDays, "N", vbTextCompare) > 0) Then
        slStr = ""
        ilDay = 1
        Do
            If Mid(slDays, ilDay, 1) = "Y" Then
                ilStart = ilDay
                ilEnd = ilStart
                ilDay = ilDay + 1
                Do
                    If ilDay > 7 Then
                        Exit Do
                    End If
                    If Mid(slDays, ilDay, 1) = "N" Then
                        Exit Do
                    Else
                        ilEnd = ilDay
                    End If
                    ilDay = ilDay + 1
                Loop
                If slStr = "" Then
                    If ilStart = ilEnd Then
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                Else
                    If ilStart = ilEnd Then
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                    Else
                        slStr = slStr & "," & Switch(ilStart = 1, "M", ilStart = 2, "Tu", ilStart = 3, "W", ilStart = 4, "Th", ilStart = 5, "F", ilStart = 6, "Sa", ilStart = 7, "Su")
                        slStr = slStr & "-" & Switch(ilEnd = 1, "M", ilEnd = 2, "Tu", ilEnd = 3, "W", ilEnd = 4, "Th", ilEnd = 5, "F", ilEnd = 6, "Sa", ilEnd = 7, "Su")
                    End If
                End If
            End If
            ilDay = ilDay + 1
        Loop While ilDay <= 7
        slDays = slStr
    Else
        If (slDays = "MoTuWeThFrSaSu") Then
            slDays = "M-Su"
        ElseIf (slDays = "MoTuWeThFrSa") Then
            slDays = "M-Sa"
        ElseIf (slDays = "MoTuWeThFr") Then
            slDays = "M-F"
        ElseIf (slDays = "MoTuWeTh") Then
            slDays = "M-Th"
        ElseIf (slDays = "MoTuWe") Then
            slDays = "M-W"
        ElseIf slDays = ("MoTu") Then
            slDays = "M-Tu"
        ElseIf (slDays = "TuWeThFrSaSu") Then
            slDays = "Tu-Su"
        ElseIf (slDays = "TuWeThFrSa") Then
            slDays = "Tu-Sa"
        ElseIf (slDays = "TuWeThFr") Then
            slDays = "Tu-F"
        ElseIf (slDays = "TuWeTh") Then
            slDays = "Tu-Th"
        ElseIf (slDays = "TuWe") Then
            slDays = "Tu-W"
        ElseIf (slDays = "WeThFrSaSu") Then
            slDays = "W-Su"
        ElseIf (slDays = "WeThFrSa") Then
            slDays = "W-Sa"
        ElseIf (slDays = "WeThFr") Then
            slDays = "W-F"
        ElseIf (slDays = "WeTh") Then
            slDays = "W-Th"
        ElseIf (slDays = "ThFrSaSu") Then
            slDays = "Th-Su"
        ElseIf (slDays = "ThFrSa") Then
            slDays = "Th-Sa"
        ElseIf (slDays = "ThFr") Then
            slDays = "Th-F"
        ElseIf slDays = "FrSaSu" Then
            slDays = "F-Su"
        ElseIf slDays = "FrSa" Then
            slDays = "F-Sa"
        ElseIf slDays = "SaSu" Then
            slDays = "S-S"
        End If
    End If
    gDayMap = slDays
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
    If slTime = "24:00:00" Then
        llTime = 86400
        gTimeToLong = llTime
        Exit Function
    End If
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

'    slTime = Trim$(slInpTime)
'    If slTime = "" Then
'        slTime = "12:00AM"
'    End If
'    slFixedTime = UCase$(slTime)
'    If (InStr(slFixedTime, "A") = 0) And (InStr(slFixedTime, "P") = 0) And (InStr(slFixedTime, "N") = 0) And (InStr(slFixedTime, "M") = 0) Then
'        slFixedTime = Format$(slFixedTime, "hh:mm:ss am/pm")
'    End If
'    ilPos1 = InStr(slFixedTime, "N")
'    If ilPos1 <> 0 Then
'        slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "PM"
'    End If
'    ilPos1 = InStr(slFixedTime, "A")
'    ilPos2 = InStr(slFixedTime, "P")
'    If (ilPos1 = 0) And (ilPos2 = 0) Then
'        ilPos1 = InStr(slFixedTime, "M")
'        If ilPos1 <> 0 Then
'            slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "AM"
'        End If
'    End If
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
    llDate = gDateValue(slAnyDate)
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbSunday
        llDate = llDate + 1
    Loop
    gObtainNextSunday = Format$(llDate, sgShowDateForm)

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
    llDate = gDateValue(slDate)
    Do While Month(llDate) = Month(llDate + 7)
        llDate = llDate + 7
    Loop
    gObtainEndStd = Format$(llDate, sgShowDateForm)
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name: gCtrlGotFocus                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Text in the control is          *
'*                     highlighted                     *
'*                                                     *
'*******************************************************
Public Sub gCtrlGotFocus(Ctrl As Control)

'   Where:
'       Ctrl (I)- control for which text will be highlighted
'

    If TypeOf Ctrl Is TextBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.text)
    End If
End Sub
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
    llDate = gDateValue(slAnyDate)
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbMonday
        llDate = llDate - 1
    Loop
    gObtainPrevMonday = Format$(llDate, sgShowDateForm)

End Function


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
    llDate = gDateValue(slAnyDate)
    Do While Weekday(Format$(llDate, "m/d/yyyy")) <> vbMonday
        llDate = llDate + 1
    Loop
    gObtainNextMonday = Format$(llDate, sgShowDateForm)

End Function


'***************************************************************************************
'*
'* Procedure Name: gSetPathEndSlash
'*
'* Created: 10/02/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Install the final back slash if it does not already exist.
'*
'***************************************************************************************
Public Function gSetPathEndSlash(ByVal sPath As String) As String
    If right$(sPath, 1) <> "\" Then
        sPath = sPath + "\"
    End If
    gSetPathEndSlash = sPath
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
    Dim ilYear As Long
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
        If slYear = "" Then
            gAdjYear = slAnyDate
            Exit Function
        End If
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


Public Function gIsDate(slInpDate As String) As Integer

    Dim ilDate As Integer
    Dim slDate As String
    Dim slAnyDate As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slStr As String

    slAnyDate = Trim$(slInpDate)
    If Len(slAnyDate) <> 0 Then
        slDate = gAdjYear(slAnyDate)
        On Error GoTo gIsDateErr
        ilDate = Day(slDate)
        ilDate = Month(slDate)
        ilDate = Year(slDate)
        If (ilDate < 1970) Or (ilDate > 2069) Then
            gIsDate = False
            Exit Function
        End If
        ilPos1 = InStr(1, slDate, "/", vbTextCompare)
        slStr = Mid$(slDate, 1, ilPos1 - 1)
        If (Val(slStr) < 1) Or (Val(slStr) > 12) Then
            gIsDate = False
            Exit Function
        End If
        ilPos1 = ilPos1 + 1
        ilPos2 = InStr(ilPos1, slDate, "/", vbTextCompare)
        slStr = Mid$(slDate, ilPos1, ilPos2 - ilPos1)
        If (Val(slStr) < 1) Or (Val(slStr) > 31) Then
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
gIsDateErr:
    On Error GoTo 0
    gIsDate = False
    Exit Function

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
        slLastPos = UCase$(right$(sgShowTimeWSecForm, 1))
        If (slLastPos = "A") Or (slLastPos = "P") Or (slLastPos = "M") Or (slLastPos = "N") Then
        
            slLastPos = UCase$(right$(slTime, 1))
            If (slLastPos <> "A") And (slLastPos <> "P") And (slLastPos <> "M") And (slLastPos <> "N") Then
                gIsTime = False
                Exit Function
            End If
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
        If slTime = "24:00:00" Then
            gIsTime = True
            Exit Function
        End If
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

'*******************************************************
'*                                                     *
'* Procedure Name: gDateValue                          *
'*                                                     *
'*             Created:2/22/06       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:                                *
'*                                                     *
'*******************************************************
Function gDateValue(slDate As String) As Long
    If Trim$(slDate) = "" Then
        gDateValue = 0
        Exit Function
    End If
    gDateValue = DateValue(gAdjYear(slDate))
End Function

'***************************************************************************************
'*
'* Procedure Name: gLoadOption
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function loads a string value from the ini file.
'*           It relies on the global variable sgIniPathFileName to
'*           contain the path and name of the ini file to use.
'*
'***************************************************************************************
Public Function gLoadOption(Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    gLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, sgIniPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadOption = True
        End If
    End If
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function



Public Sub gClearControls(Frm As Form)
    Dim Ctrl As Control
    For Each Ctrl In Frm.Controls
        If TypeOf Ctrl Is TextBox Then
            Ctrl.text = ""
        ElseIf TypeOf Ctrl Is OptionButton Then
            Ctrl.Value = False
        End If
    Next Ctrl
End Sub




Public Function gStrTimeInTenthToLong(slTimeInTenths As String, ilChk12M As Integer) As Long
    Dim llTime As Long
    Dim ilPos As Integer
    Dim ilTPos As Integer
    Dim slAnyTime As String
    Dim llTenths As Long

    ilPos = InStr(slTimeInTenths, "-")
    ilTPos = InStr(slTimeInTenths, ".")
    If ilTPos > 0 Then
        llTenths = Val(Mid$(slTimeInTenths, ilTPos + 1))
        slAnyTime = Left$(slTimeInTenths, ilTPos - 1)
    Else
        llTenths = 0
        slAnyTime = slTimeInTenths
    End If
    slAnyTime = Format$(slAnyTime, "hh:mm:ss am/pm")
    llTime = gTimeToLong(slAnyTime, ilChk12M)
    If ilPos > 0 Then
        llTime = -1 * (llTime * 10 + llTenths)
    Else
        llTime = llTime * 10 + llTenths
    End If
    gStrTimeInTenthToLong = llTime
    Exit Function
gStrTimeInTenthToLongErr:
    On Error GoTo 0
    gStrTimeInTenthToLong = 0
    Exit Function
    
End Function

Public Function gLongToStrTimeInTenth(llInpTime As Long) As String
    Dim llTenths As Long
    Dim llTime As Long
    Dim slTenths As String
    Dim slTime As String
    
    llTenths = llInpTime Mod 10
    slTenths = Trim$(Str$(llTenths))
    llTime = llInpTime \ 10
    slTime = gLongToTime(llTime)
    gLongToStrTimeInTenth = Format(slTime, sgShowTimeWSecForm) & "." & slTenths
End Function

Public Function gIsTimeTenths(slTimeInTenths As String) As Integer
    Dim ilPos As Integer
    Dim ilTPos As Integer
    Dim slAnyTime As String
    Dim ilRet As Integer
    Dim slTenths As String
    Dim slTime As String
    
    slTime = Trim$(slTimeInTenths)
    If Len(Trim$(slTime)) <= 0 Then
        gIsTimeTenths = True
        Exit Function
    End If
    ilTPos = InStr(slTime, ".")
    If ilTPos > 0 Then
        ilPos = InStr(1, UCase$(slTime), "A", vbTextCompare)
        If ilPos = 0 Then
            ilPos = InStr(1, UCase$(slTime), "P", vbTextCompare)
            If ilPos = 0 Then
                ilPos = InStr(1, UCase$(slTime), "M", vbTextCompare)
                If ilPos = 0 Then
                    ilPos = InStr(1, UCase$(slTime), "N", vbTextCompare)
                    If ilPos = 0 Then
                    End If
                End If
            End If
        End If
        If ilPos = 0 Then
            slTenths = Trim$(Mid$(slTime, ilTPos + 1))
            If Len(slTenths) > 1 Then
                gIsTimeTenths = False
                Exit Function
            End If
            slAnyTime = Left$(slTime, ilTPos - 1)
        Else
            If ilPos <= ilTPos Then
                gIsTimeTenths = False
                Exit Function
            End If
            If ilPos = ilTPos + 1 Then
                gIsTimeTenths = False
                Exit Function
            End If
            slTenths = Trim$(Mid$(slTime, ilTPos + 1, ilPos - ilTPos - 1))
            If Len(slTenths) > 1 Then
                gIsTimeTenths = False
                Exit Function
            End If
            slAnyTime = Left$(slTime, ilTPos - 1) & Mid$(slTime, ilPos)
        End If
        If Len(slTenths) > 0 Then
            If (Asc(slTenths) < Asc("0")) Or (Asc(slTenths) > Asc("9")) Then
                gIsTimeTenths = False
                Exit Function
            End If
        End If
    Else
        slAnyTime = slTime
    End If
    slAnyTime = Format$(slAnyTime, "hh:mm:ss am/pm")
    ilRet = gIsTime(slAnyTime)
    gIsTimeTenths = ilRet
End Function


Public Function gGetListBoxRow(lbcCtrl As ListBox, y As Single) As Long
    Dim flRowHeight As Single  'Standard text height
    Dim llRow As Long
    Dim slStr As String
    Dim llRowHeight As Long
    
    
    lbcCtrl.ToolTipText = ""
    If lbcCtrl.ListCount <= 0 Then
        gGetListBoxRow = -1
        Exit Function
    End If
    llRowHeight = 15 * SendMessageByString(lbcCtrl.hwnd, LB_GETITEMHEIGHT, 0, slStr)
    'flRowHeight = (lbcCtrl.Height - 30) / llRowHeight
    llRow = (y - 15) \ llRowHeight + lbcCtrl.TopIndex
    If llRow >= lbcCtrl.ListCount Then
        gGetListBoxRow = -1
        Exit Function
    End If
    gGetListBoxRow = llRow
End Function

Public Function gListBoxFind(lbcCtrl As ListBox, slFindString As String, Optional blExactOnly As Boolean = False) As Long
    Dim llStartPoint As Long
    Dim llRow As Long
    
    llStartPoint = 0
    If Left$(slFindString, 1) <> "[" Then
        If lbcCtrl.ListCount >= 0 Then
            If Left$(lbcCtrl.List(0), 1) = "[" Then
                llStartPoint = llStartPoint + 1
            End If
        End If
        If lbcCtrl.ListCount > 0 Then
            If Left$(lbcCtrl.List(1), 1) = "[" Then
                llStartPoint = llStartPoint + 1
            End If
        End If
    End If
    llRow = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRINGEXACT, llStartPoint, slFindString)
    If (llRow < 0) And (blExactOnly = False) Then
        llRow = SendMessageByString(lbcCtrl.hwnd, LB_FINDSTRING, llStartPoint, slFindString)
    End If
    gListBoxFind = llRow
End Function

Public Function gFormatTime(slInpTime As String) As String
    Dim slTime As String
    If Not gIsTime(slInpTime) Then
        gFormatTime = ""
        Exit Function
    End If
    slTime = gConvertTime(slInpTime)
    gFormatTime = Format(slTime, sgShowTimeWSecForm)
End Function

Public Function gFormatTimeTenths(slInpTime As String) As String
    Dim slTime As String
    Dim ilTPos As Integer
    Dim slTenths As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    
    If Not gIsTimeTenths(slInpTime) Then
        gFormatTimeTenths = ""
        Exit Function
    End If
    ilTPos = InStr(slInpTime, ".")
    If ilTPos > 0 Then
        slTime = Left$(slInpTime, ilTPos - 1)
        slTenths = Mid$(slInpTime, ilTPos)
        ilPos1 = InStr(slInpTime, ":")
        ilPos2 = InStr(ilPos1 + 1, slInpTime, ":")
        If ilPos2 <= 0 Then
            slTime = "00:" & slTime
        End If
    Else
        slTime = slInpTime
        slTenths = ".0"
    End If
    slTime = gConvertTime(slTime)
    gFormatTimeTenths = Format(slTime, sgShowTimeWSecForm) & slTenths
End Function

Public Function gLongToLength(llInTime As Long, ilWithHours As Integer) As String
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilSec As Integer
    Dim llTime As Long
    Dim slTime As String
    Dim ilPos As Integer
    
    llTime = llInTime
    ilHour = llTime \ 3600
    Do While ilHour > 23
        ilHour = ilHour - 24
    Loop
    llTime = llTime Mod 3600
    ilMin = llTime \ 60
    ilSec = llTime Mod 60
    slTime = ""
'    If (ilHour = 0) Or (ilHour = 12) Then
'        slTime = "12"
'    Else
'        If ilHour < 12 Then
'            slTime = Trim$(Str$(ilHour))
'        Else
'            slTime = Trim$(Str$(ilHour - 12))
'        End If
'    End If
    slTime = Trim$(Str$(ilHour))
    If Len(slTime) = 1 Then
        slTime = "0" & slTime
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
    Else
        If ilMin = 0 Then
            slTime = slTime & ":00:00"
        Else
            slTime = slTime & ":00"
        End If
    End If
    If Not ilWithHours Then
        ilPos = InStr(1, slTime, ":", vbTextCompare)
        If ilPos > 0 Then
            slTime = Mid$(slTime, ilPos + 1)
        End If
    End If
    gLongToLength = slTime
End Function

Public Function gStrLengthInTenthToLong(slLengthInTenths As String) As Long
'
'   llRetLength = gStrLengthInTenthToLong(slLength)
'   Where:
'       slLength (I)- Length as string to be converted to currency
'       llRetLength (O)- length as long
'
    Dim slLength As String
    Dim llLength As Long
    Dim ilPos As Integer
    Dim ilLoc As Integer
    Dim ilLoc2 As Integer
    Dim slAnyLength As String
    Dim ilTPos As Integer
    Dim llTenths As Long

    On Error GoTo gStrLengthInTenthToLongErr
    ilTPos = InStr(slLengthInTenths, ".")
    If ilTPos > 0 Then
        llTenths = Val(Mid$(slLengthInTenths, ilTPos + 1))
        slAnyLength = Left$(slLengthInTenths, ilTPos - 1)
    Else
        llTenths = 0
        slAnyLength = slLengthInTenths
    End If
    ilPos = InStr(slAnyLength, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gStrLengthInTenthToLong = 0
            Exit Function
        End If
        slLength = UCase$(Mid$(slAnyLength, 2))
    Else
        slLength = UCase$(slAnyLength)
    End If
    llLength = 0
    ilLoc = InStr(1, slLength, "H", vbTextCompare)
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, ":", 1)
        ilLoc2 = InStr(ilLoc + 1, slLength, ":", vbTextCompare)
        If ilLoc2 = 0 Then
            ilLoc = 0
        End If
    End If
    If ilLoc <> 0 Then
        llLength = llLength + 3600 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "M", 1)
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, ":", 1)
    End If
    If ilLoc <> 0 Then
        llLength = llLength + 60 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "S", 1)
    If ilLoc <> 0 Then
        llLength = llLength + Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    Else
        llLength = llLength + Val(slLength)
    End If
    If ilPos = 0 Then
        gStrLengthInTenthToLong = llLength * 10 + llTenths
    Else
        gStrLengthInTenthToLong = -1 * (llLength * 10 + llTenths)
    End If
    Exit Function
gStrLengthInTenthToLongErr:
    On Error GoTo 0
    gStrLengthInTenthToLong = 0
    Exit Function
    
End Function

Function gIsLength(slInpLength As String) As Integer
'
'   ilRet = gIsLength (slLength)
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
        gIsLength = True
        Exit Function
    End If
    slHour = ""
    slMin = ""
    slSec = ""
    slLen = Trim$(slLength)
    ilPos = InStr(1, slLen, "::")
    If ilPos > 0 Then
        gIsLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hm", 1)
    If ilPos > 0 Then
        gIsLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "hs", 1)
    If ilPos > 0 Then
        gIsLength = False
        Exit Function
    End If
    ilPos = InStr(1, slLen, "ms", 1)
    If ilPos > 0 Then
        gIsLength = False
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
                    ilPos = InStr(1, slLen, "'", 1)
                    If ilPos > 0 Then
                        ilFormat = 2
                    Else
                        ilPos = InStr(1, slLen, """", 1)
                        If ilPos > 0 Then
                            ilFormat = 2
                        Else
                            ilFormat = 1
                        End If
                    End If
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
            gIsLength = False
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
            gIsLength = False
            Exit Function
         End If
    Else    'format hh:mm:ss
        ilPos = InStr(slLen, ":")
        If ilPos > 0 Then   'Might be hour/min/sec or min/sec only
            If InStr(ilPos + 1, slLen, ":", vbTextCompare) > 0 Then
                slHour = Left$(slLen, ilPos - 1)
                slLen = Mid$(slLen, ilPos + 1)
                ilPos = InStr(1, slLen, ":")
            Else
                slHour = "0"
            End If
            If ilPos > 0 Then
                slMin = Left$(slLen, ilPos - 1)
                slSec = Mid$(slLen, ilPos + 1)
            Else
                slMin = slLen
            End If
        Else
            slSec = slLen
        End If
    End If
    On Error GoTo gIsLengthErr
    If slHour <> "" Then
        If Len(slHour) > 2 Then
            gIsLength = False
            Exit Function
        End If
        llHour = CLng(slHour) 'Val(slHour)
    Else
        llHour = 0
    End If
    If slMin <> "" Then
        If Len(slMin) > 2 Then
            gIsLength = False
            Exit Function
        End If
        ilMin = CLng(slMin) 'Val(slMin)
    Else
        ilMin = 0
    End If
    If slSec <> "" Then
        If Len(slSec) > 3 Then
            gIsLength = False
            Exit Function
        End If
        ilSec = CLng(slSec) 'Val(slSec)
        If ilSec > 59 Then
            ilMin = ilSec \ 60
            ilSec = ilSec Mod 60
        End If
    Else
        ilSec = 0
    End If
    If (llHour < 0) Or (llHour > 24) Then
        gIsLength = False
        Exit Function
    End If
    If (llHour <> 0) Then
        If (ilMin < 0) Or (ilMin > 59) Then
            gIsLength = False
            Exit Function
        End If
    End If
    If llHour <> 0 Then
        If (ilSec < 0) Or (ilSec > 59) Then
            gIsLength = False
            Exit Function
        End If
    End If
    If llHour * 3600& + CLng(ilMin) * 60& + CLng(ilSec) > 86400 Then
        gIsLength = False
        Exit Function
    End If
    gIsLength = True
    Exit Function
gIsLengthErr:
    On Error GoTo 0
    gIsLength = False
    Exit Function
End Function

Function gLengthToLong(slInpLength As String) As Long
'
'   llRetLength = gLengthToLong(slLength)
'   Where:
'       slLength (I)- Length as string to be converted to currency
'       llRetLength (O)- length as long
'
    Dim slLength As String
    Dim llLength As Long
    Dim ilPos As Integer
    Dim ilLoc As Integer
    Dim ilLoc2 As Integer
    Dim slAnyLength As String

    slAnyLength = Trim$(slInpLength)
    On Error GoTo gLengthToLongErr
    ilPos = InStr(slAnyLength, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gLengthToLong = 0
            Exit Function
        End If
        slLength = UCase$(Mid$(slAnyLength, 2))
    Else
        slLength = UCase$(slAnyLength)
    End If
    llLength = 0
    ilLoc = InStr(1, slLength, "H", vbTextCompare)
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, ":", 1)
        ilLoc2 = InStr(ilLoc + 1, slLength, ":", vbTextCompare)
        If ilLoc2 = 0 Then
            ilLoc = 0
        End If
    End If
    If ilLoc <> 0 Then
        llLength = llLength + 3600 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "M", 1)
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, ":", 1)
    End If
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, "'", 1)
    End If
    If ilLoc <> 0 Then
        llLength = llLength + 60 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "S", 1)
    If ilLoc = 0 Then
        ilLoc = InStr(1, slLength, """", 1)
    End If
    If ilLoc <> 0 Then
        llLength = llLength + Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    Else
        llLength = llLength + Val(slLength)
    End If
    If ilPos = 0 Then
        gLengthToLong = llLength
    Else
        gLengthToLong = -llLength
    End If
    Exit Function
gLengthToLongErr:
    On Error GoTo 0
    gLengthToLong = 0
    Exit Function
End Function

Public Function gLongToStrLengthInTenth(llInpTime As Long, ilWithHours As Integer) As String
    Dim llTenths As Long
    Dim llTime As Long
    Dim slTenths As String
    Dim slTime As String
    
    llTenths = llInpTime Mod 10
    slTenths = Trim$(Str$(llTenths))
    llTime = llInpTime \ 10
    slTime = gLongToLength(llTime, ilWithHours)
    gLongToStrLengthInTenth = slTime & "." & slTenths
End Function

Public Function gIsLengthTenths(slLengthInTenths As String) As Integer
    Dim ilPos As Integer
    Dim ilTPos As Integer
    Dim slAnyLength As String
    Dim ilRet As Integer
    Dim slTenths As String
    Dim slLength As String
    
    slLength = Trim$(slLengthInTenths)
    If Len(Trim$(slLength)) <= 0 Then
        gIsLengthTenths = True
        Exit Function
    End If
    ilTPos = InStr(slLength, ".")
    If ilTPos > 0 Then
        slTenths = Trim$(Mid$(slLength, ilTPos + 1))
        If Len(slTenths) > 1 Then
            gIsLengthTenths = False
            Exit Function
        End If
        slAnyLength = Left$(slLength, ilTPos - 1)
        If Len(slTenths) > 0 Then
            If (Asc(slTenths) < Asc("0")) Or (Asc(slTenths) > Asc("9")) Then
                gIsLengthTenths = False
                Exit Function
            End If
        End If
    Else
        slAnyLength = slLength
    End If
    ilRet = gIsLength(slAnyLength)
    gIsLengthTenths = ilRet
End Function

Public Function gHourMap(slHours As String) As String
    'slHours(I)- string of 24 N's and Y's
    Dim ilHour As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim slStr As String
    
    slStr = ""
    ilHour = 1
    Do
        If Mid(slHours, ilHour, 1) = "Y" Then
            ilStart = ilHour
            ilEnd = ilStart
            ilHour = ilHour + 1
            Do
                If ilHour > 24 Then
                    Exit Do
                End If
                If Mid(slHours, ilHour, 1) = "N" Then
                    Exit Do
                Else
                    ilEnd = ilHour
                End If
                ilHour = ilHour + 1
            Loop
            If slStr = "" Then
                If ilStart = ilEnd Then
                    slStr = Trim$(Str$(ilStart - 1))
                Else
                    slStr = Trim$(Str$(ilStart - 1)) & "-" & Trim$(Str$(ilEnd - 1))
                End If
            Else
                If ilStart = ilEnd Then
                    slStr = slStr & "," & Trim$(Str$(ilStart - 1))
                Else
                    slStr = slStr & "," & Trim$(Str$(ilStart - 1)) & "-" & Trim$(Str$(ilEnd - 1))
                End If
            End If
        End If
        ilHour = ilHour + 1
    Loop While ilHour <= 24
    gHourMap = slStr
End Function

Public Function gCheckSum(slInStr As String) As Integer
    Dim ilResult As Integer
    Dim ilLoop As Integer
    
    ilResult = 0
    For ilLoop = 1 To Len(slInStr) Step 1
        ilResult = ilResult Xor Asc(Mid(slInStr, ilLoop, 1))
    Next ilLoop
    gCheckSum = ilResult
End Function

Public Function gIntToHex(ilInValue As Integer, ilOutputLength As Integer) As String
    Dim ilValue As Integer
    Dim ilMod As Integer
    Dim slHex As String
    
    slHex = ""
    ilValue = ilInValue
    Do While ilValue > 0
        ilMod = ilValue Mod 16
        slHex = Switch(ilMod = 0, "0", ilMod = 1, "1", ilMod = 2, "2", ilMod = 3, "3", ilMod = 4, "4", _
                       ilMod = 5, "5", ilMod = 6, "6", ilMod = 7, "7", ilMod = 8, "8", ilMod = 9, "9", _
                       ilMod = 10, "A", ilMod = 11, "B", ilMod = 12, "C", ilMod = 13, "D", ilMod = 14, "E", _
                       ilMod = 15, "F") & slHex
        ilValue = ilValue \ 16
    Loop
    Do While Len(slHex) < ilOutputLength
        slHex = "0" & slHex
    Loop
    gIntToHex = slHex
End Function

Public Function gLongToHex(llInValue As Long, ilOutputLength As Integer) As String
    Dim llValue As Long
    Dim ilMod As Integer
    Dim slHex As String
    
    slHex = ""
    llValue = llInValue
    Do While llValue > 0
        ilMod = llValue Mod 16
        slHex = Switch(ilMod = 0, "0", ilMod = 1, "1", ilMod = 2, "2", ilMod = 3, "3", ilMod = 4, "4", _
                       ilMod = 5, "5", ilMod = 6, "6", ilMod = 7, "7", ilMod = 8, "8", ilMod = 9, "9", _
                       ilMod = 10, "A", ilMod = 11, "B", ilMod = 12, "C", ilMod = 13, "D", ilMod = 14, "E", _
                       ilMod = 15, "F") & slHex
        llValue = llValue \ 16
    Loop
    Do While Len(slHex) < ilOutputLength
        slHex = "0" & slHex
    Loop
    gLongToHex = slHex
End Function


Public Sub gLogMsg(sMsg As String, sFileName As String, iKill As Integer)

    'D.S. 4/04
    'Purpose: A general file routine that shows: Date and Time followed by a message
    'so we can try to stop adding a separate file routine to every single module
    
    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'iKill = True then delete the file first, iKill = False then append to the file
    
    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slToFile As String
    
    slToFile = sgMsgDirectory & sFileName
    On Error GoTo Error

    If iKill = True Then
        ilRet = 0
        slDateTime = FileDateTime(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
    End If
    
    hlLogFile = FreeFile
    Open slToFile For Append As hlLogFile
    slFullMsg = Format$(Now, "mm/dd/yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " " & sMsg
    Print #hlLogFile, slFullMsg
    Close hlLogFile
    
    slFullMsg = UCase(slFullMsg)
    If InStr(1, slFullMsg, "ERROR", vbTextCompare) > 0 Then
        gSaveStackTrace slToFile
    End If

    Exit Sub
    
Error:
    ilRet = 1
    Resume Next
    
End Sub

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
    'Remove single quotes '
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
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "'"
            ilFound = True
        End If
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

Public Sub gMsgBox(slErrMsg As String, llMsgBoxStyle As Long, slTitle As String)
    If igBkgdProg = 0 Then
        If llMsgBoxStyle <> -1 Then
            If slTitle <> "" Then
                MsgBox slErrMsg, llMsgBoxStyle, slTitle
            Else
                MsgBox slErrMsg, llMsgBoxStyle
            End If
        Else
            gLogMsg slErrMsg, "EngrErrors.Txt", False
        End If
    ElseIf igBkgdProg = 1 Then
        gLogMsg slErrMsg, "Bkgd_Schd.Txt", False
    ElseIf igBkgdProg = 2 Then
        gLogMsg slErrMsg, "Set_Credit.Txt", False
    ElseIf igBkgdProg = 3 Then
        gLogMsg slErrMsg, "ExptSQL.Txt", False
    ElseIf igBkgdProg = 10 Then
        gLogMsg slErrMsg, "EngrService.Txt", False
    Else
        gLogMsg slErrMsg, "EngrErrors.Txt", False
    End If
End Sub

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
    slDate = Trim$(Format$(dlDateSerial, "m/d/yy"))  '"ddddd"))
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

Public Function gGetComputerName() As String

'D.S. 2/9/05 Returns the name of the users computer

   Dim strBuffer As String * 255

   If GetComputerName(strBuffer, 255&) <> 0 Then
      ' Name exist
      gGetComputerName = Left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
   Else
      ' Name doesn't exist
      gGetComputerName = "N/A"
   End If
   
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFieldOffset                    *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the offset to a field *
'*                     within a record                 *
'*                                                     *
'*******************************************************
Function gFieldOffset(slFile As String, slField As String) As Integer
'
'   sFile = "ADF"
'   sField = "slfFirstName"
'   ilOffset = gFieldOffset(sFile, sField)
'   Where:
'       sFile (I)- Name of the file
'       sField (I)- Name of the field as in the DDF
'       ilOffset (O)- The offset of the start of the field from 0 (-1 if not found)
'

    Dim slMsg As String
    Dim ilOffset As Integer

    slMsg = ""
    ilOffset = csiGetOffset(UCase$(slFile), UCase$(slField))
    If ilOffset >= 0 Then
        gFieldOffset = ilOffset
    Else
        slMsg = "Offset to field " & slField & " missing from file " & slFile
    End If
    If slMsg = "" Then
        Exit Function
    End If
    gMsgBox slMsg, vbOKOnly + vbCritical, "Offset Error"
    gFieldOffset = -1
End Function

