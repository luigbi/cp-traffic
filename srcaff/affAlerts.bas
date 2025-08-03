Attribute VB_Name = "modAlerts"
'******************************************************
'* Copyright 1993 Counterpoint Software®. All rights reserved.
'* Proprietary Software, Do not copy
'*
'*  modAlerts - Alert support routines
'*
'******************************************************
Option Explicit
Option Compare Text

'Alert Menu Item
Public Const MF_BITMAP = &H4&
Public Const MF_BYPOSITION = &H400&
Public Const MF_ENABLED = &H0
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
' Bitmap Header Definition
Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long


Private rstAlertClear As ADODB.Recordset
Private rstAlertAdd As ADODB.Recordset
Private rstAlertClearFinal As ADODB.Recordset







'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertCheck                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if Alerts should be   *
'*                     turn on or off                  *
'*                                                     *
'*******************************************************
Public Function gAlertCheck() As Integer
    Dim ilRet As Integer
    Dim slSQL_AlertCheck As String
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then
        gAlertCheck = True
        Exit Function
    End If
    If igAlertInterval <> 0 Then
        If ((igAlertInterval <> 0) And (igAlertInterval <= igAlertTimer)) Then
            'Check if Alert exist
            If (sgExptSpotAlert <> "N") Then
                slSQL_AlertCheck = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufStatus = 'R'"
                Set rstAlert = gSQLSelectCall(slSQL_AlertCheck)
                If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
                    frmMain!tmcCheckAlert.Enabled = False
                    igAlertFlash = -1
                    igAlertTimer = 0
                    frmMain!tmcFlashAlert.Interval = 2000   'Every 2 second
                    frmMain!tmcFlashAlert.Enabled = True
                    DoEvents
                    gAlertCheck = True
                    Exit Function
                End If
            End If
            If (sgExptISCIAlert <> "N") Then
                slSQL_AlertCheck = "SELECT * FROM AUF_ALERT_USER WHERE (aufType = 'F' or aufType = 'R') AND aufStatus = 'R'"
                Set rstAlert = gSQLSelectCall(slSQL_AlertCheck)
                If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
                    frmMain!tmcCheckAlert.Enabled = False
                    igAlertFlash = -1
                    igAlertTimer = 0
                    frmMain!tmcFlashAlert.Interval = 2000   'Every 2 second
                    frmMain!tmcFlashAlert.Enabled = True
                    DoEvents
                    gAlertCheck = True
                    Exit Function
                End If
            End If
            '8273
            If gAllowVendorAlerts(True) Then
                '8226
                slSQL_AlertCheck = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'V' AND aufStatus = 'R'"
                Set rstAlert = gSQLSelectCall(slSQL_AlertCheck)
                If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
                    frmMain!tmcCheckAlert.Enabled = False
                    igAlertFlash = -1
                    igAlertTimer = 0
                    frmMain!tmcFlashAlert.Interval = 2000   'Every 2 second
                    frmMain!tmcFlashAlert.Enabled = True
                    DoEvents
                    gAlertCheck = True
                    Exit Function
                End If
            End If
            slSQL_AlertCheck = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'P' AND aufStatus = 'R'"
            Set rstAlert = gSQLSelectCall(slSQL_AlertCheck)
            If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
                frmMain!tmcCheckAlert.Enabled = False
                igAlertFlash = -1
                igAlertTimer = 0
                frmMain!tmcFlashAlert.Interval = 2000   'Every 2 second
                frmMain!tmcFlashAlert.Enabled = True
                DoEvents
                gAlertCheck = True
                Exit Function
            End If
            'Unfound Pool
            slSQL_AlertCheck = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'U' AND aufStatus = 'R'"
            Set rstAlert = gSQLSelectCall(slSQL_AlertCheck)
            If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
                frmMain!tmcCheckAlert.Enabled = False
                igAlertFlash = -1
                igAlertTimer = 0
                frmMain!tmcFlashAlert.Interval = 2000   'Every 2 second
                frmMain!tmcFlashAlert.Enabled = True
                DoEvents
                gAlertCheck = True
                Exit Function
            End If
            frmMain!tmcFlashAlert.Enabled = False
            frmMain!tmcCheckAlert.Enabled = True
            frmMain!mnuAlert.Visible = False
            igAlertTimer = 1
        Else
            igAlertTimer = igAlertTimer + 1
        End If
    Else
        frmMain!tmcCheckAlert.Enabled = True
        frmMain!mnuAlert.Visible = False
    End If
    gAlertCheck = False
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertCheck"
    gAlertCheck = False
    Exit Function
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertAdd                       *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Alert if not previously     *
'*                     added                           *
'*                                                     *
'*******************************************************
Public Function gAlertAdd(slType As String, slSubType As String, ilVefCode As Integer, slInWeekDate As String) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim slMoWeekDate As String
    Dim ilLoop As Integer
    Dim slUlfCode As String
    Dim llUlfCode As Long
    Dim slCefCode As String
    Dim llCefCode As Long
    Dim slCountdown As String
    Dim ilCountdown As Integer
    Dim slSQL_AlertAdd As String
    
    On Error GoTo ErrHand
    
    If (slType <> "N") And (slType <> "B") And (slType <> "U") Then
        slMoWeekDate = slInWeekDate
        Do While Weekday(slMoWeekDate, vbSunday) <> vbMonday
            slMoWeekDate = DateAdd("d", -1, slMoWeekDate)
        Loop
        llMoWeekDate = DateValue(gAdjYear(slMoWeekDate))
    ElseIf (slType = "U") Then
        llMoWeekDate = DateValue(gAdjYear(slInWeekDate))
    Else
        llMoWeekDate = 0
    End If
    ilFound = False
    If slType <> "N" Then
        slSQL_AlertAdd = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
        Set rstAlertAdd = gSQLSelectCall(slSQL_AlertAdd)
        Do While (Not rstAlertAdd.EOF) And (Not rstAlertAdd.BOF)
            tgAuf.lChfCode = rstAlertAdd!aufChfCode
            tgAuf.sStatus = rstAlertAdd!aufStatus
            tgAuf.sSubType = rstAlertAdd!aufSubType
            tgAuf.iVefCode = rstAlertAdd!aufVefCode
            If IsNull(rstAlertAdd!aufMoWeekDate) Then
                tgAuf.lMoWeekDate = 0
            ElseIf Not gIsDate(rstAlertAdd!aufMoWeekDate) Then
                tgAuf.lMoWeekDate = 0
            Else
                tgAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlertAdd!aufMoWeekDate, sgShowDateForm)))
            End If
            If slType = "F" Then    'Affiliate Export-Final
                If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                    If tgAuf.lMoWeekDate = llMoWeekDate Then
                        ilFound = True
                        Exit Do
                    End If
                End If
            ElseIf slType = "R" Then    'Affiliate Export-Reprint
                If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                    If tgAuf.lMoWeekDate = llMoWeekDate Then
                        ilFound = True
                        Exit Do
                    End If
                End If
            ElseIf slType = "U" Then    'Unfound Pool
                If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                    If tgAuf.lMoWeekDate = llMoWeekDate Then
                        ilFound = True
                        Exit Do
                    End If
                End If
            ElseIf slType = "B" Then
                ilFound = True
                ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
                ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
                ilRet = gParseItem(slInWeekDate, 3, "|", slCountdown)
                ilCountdown = Val(slCountdown)
                slSQL_AlertAdd = "UPDATE AUF_ALERT_USER SET "
                If ilCountdown <= 0 Then
                    slSQL_AlertAdd = slSQL_AlertAdd & "aufStatus = 'C'" & ", "
                    slSQL_AlertAdd = slSQL_AlertAdd & "aufClearMethod = 'M'" & ", "
                    slSQL_AlertAdd = slSQL_AlertAdd & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    slSQL_AlertAdd = slSQL_AlertAdd & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                    slSQL_AlertAdd = slSQL_AlertAdd & "aufClearUstCode = " & igUstCode & ", "
                End If
                slSQL_AlertAdd = slSQL_AlertAdd & "aufUlfCode = " & slUlfCode & ", "
                slSQL_AlertAdd = slSQL_AlertAdd & "aufCefCode = " & slCefCode & ", "
                slSQL_AlertAdd = slSQL_AlertAdd & "aufCountdown = " & slCountdown & ", "
                slSQL_AlertAdd = slSQL_AlertAdd & "aufEnteredDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                slSQL_AlertAdd = slSQL_AlertAdd & "aufEnteredTime = '" & Format$(gNow(), sgSQLTimeForm) & "' "
                slSQL_AlertAdd = slSQL_AlertAdd & "WHERE aufCode = " & rstAlertAdd!aufCode
                'cnn.Execute slSQL_AlertAdd, rdExecDirect
                If gSQLWaitNoMsgBox(slSQL_AlertAdd, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "modAlerts-gAlertAdd"
                    gAlertAdd = False
                    On Error Resume Next
                    rstAlertAdd.Close
                    Exit Function
                End If
                Exit Do
            End If
            rstAlertAdd.MoveNext
        Loop
    End If
    If Not ilFound Then
        If slSubType = "I" Then 'Check if provider defined, if not then don't add
            For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If tgVehicleInfo(ilLoop).iCode = ilVefCode Then
                    If (tgVehicleInfo(ilLoop).iCommProvArfCode <= 0) Then
                        gAlertAdd = True
                        Exit Function
                    End If
                End If
            Next ilLoop
        End If
        If slType = "B" Then    'Block
            ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
            ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
            ilRet = gParseItem(slInWeekDate, 3, "|", slCountdown)
        ElseIf slType = "N" Then    'Notification
            ilRet = gParseItem(slInWeekDate, 1, "|", slUlfCode)
            ilRet = gParseItem(slInWeekDate, 2, "|", slCefCode)
            slCountdown = "0"
        Else
            slUlfCode = "0"
            slCefCode = "0"
            slCountdown = "0"
        End If
        slSQL_AlertAdd = "INSERT INTO AUF_ALERT_USER (aufEnteredDate, aufEnteredTime, aufStatus, "
        slSQL_AlertAdd = slSQL_AlertAdd & "aufType, aufSubType, aufChfCode, "
        slSQL_AlertAdd = slSQL_AlertAdd & "aufVefCode, aufMoWeekDate, aufCreateUrfCode, "
        slSQL_AlertAdd = slSQL_AlertAdd & "aufCreateUstCode, aufClearUrfCode, aufClearUstCode, "
        slSQL_AlertAdd = slSQL_AlertAdd & "aufClearMethod, aufClearDate, aufClearTime, "
        slSQL_AlertAdd = slSQL_AlertAdd & "aufUlfCode, aufCefCode, aufCountdown)"
        slSQL_AlertAdd = slSQL_AlertAdd & "VALUES ('" & Format$(gNow(), sgSQLDateForm) & "', '" & Format$(gNow(), sgSQLTimeForm) & "', '" & "R" & "', "
        slSQL_AlertAdd = slSQL_AlertAdd & "'" & slType & "', '" & slSubType & "', " & 0 & ", "
        slSQL_AlertAdd = slSQL_AlertAdd & ilVefCode & ", "
        If llMoWeekDate = 0 Then
            slSQL_AlertAdd = slSQL_AlertAdd & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
        Else
            slSQL_AlertAdd = slSQL_AlertAdd & "'" & Format$(llMoWeekDate, sgSQLDateForm) & "', "
        End If
        slSQL_AlertAdd = slSQL_AlertAdd & "0, "
        slSQL_AlertAdd = slSQL_AlertAdd & igUstCode & ", 0, 0, "
        slSQL_AlertAdd = slSQL_AlertAdd & "'', '" & Format$("1/1/1970", sgSQLDateForm) & "', '" & Format$("1/1/1970", sgSQLDateForm) & "', "
        slSQL_AlertAdd = slSQL_AlertAdd & slUlfCode & ", " & slCefCode & ", " & slCountdown & ")"
        'cnn.Execute slSQL_AlertAdd, rdExecDirect
        If gSQLWaitNoMsgBox(slSQL_AlertAdd, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "modAlerts-gAlertAdd"
            gAlertAdd = False
            On Error Resume Next
            rstAlertAdd.Close
            Exit Function
        End If
        gAlertAdd = True
        ilRet = gAlertForceCheck()
    Else
        gAlertAdd = True
    End If
    'Dan causing error if never set 90909
    If Not rstAlertAdd Is Nothing Then
        rstAlertAdd.Close
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertAdd"
    gAlertAdd = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertClear                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Alert previously added    *
'*                                                     *
'*******************************************************
Public Function gAlertClear(slMethod As String, slType As String, slSubType As String, ilVefCode As Integer, slInWeekDate As String)
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim slMoWeekDate As String
    Dim llAufCode As Long
    Dim slSQL_AlertClear As String
    
    On Error GoTo ErrHand
    
    If (slType <> "B") And (slType <> "N") Then
        slMoWeekDate = slInWeekDate
        Do While Weekday(slMoWeekDate, vbSunday) <> vbMonday
            slMoWeekDate = DateAdd("d", -1, slMoWeekDate)
        Loop
        llMoWeekDate = DateValue(gAdjYear(slMoWeekDate))
    Else
        llMoWeekDate = 0
    End If
    
    If slType = "N" Then
        llAufCode = Val(slInWeekDate)
    Else
        llAufCode = 0
    End If
    
    
    
    ilFound = False
    slSQL_AlertClear = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
    Set rstAlertClear = gSQLSelectCall(slSQL_AlertClear)
    Do While (Not rstAlertClear.EOF) And (Not rstAlertClear.BOF)
        DoEvents
        tgAuf.lCode = rstAlertClear!aufCode
        tgAuf.sStatus = rstAlertClear!aufStatus
        tgAuf.sSubType = rstAlertClear!aufSubType
        tgAuf.iVefCode = rstAlertClear!aufVefCode
        If IsNull(rstAlertClear!aufMoWeekDate) Then
            tgAuf.lMoWeekDate = 0
        ElseIf Not gIsDate(rstAlertClear!aufMoWeekDate) Then
            tgAuf.lMoWeekDate = 0
        Else
            tgAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlertClear!aufMoWeekDate, sgShowDateForm)))
        End If
        If slType = "F" Then
            If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                If tgAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "R" Then
            If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                If tgAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "P" Then
            If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                If tgAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "B" Then
            ilFound = True
            Exit Do
        ElseIf slType = "N" Then
            If tgAuf.lCode = llAufCode Then
                ilFound = True
                Exit Do
            End If
        End If
        rstAlertClear.MoveNext
    Loop
    
    If (ilFound = False) And ((slType = "F") Or (slType = "R")) Then
        'D.S. 12/09/04 - If we get here we failed to find the alert, probably because someone put in a
        'new AUF file.  If we don't have the alert set and then set the exported flag we won't know that
        'it's been exported in the past which will cause problems in future exports
        ilRet = gAlertAdd(slType, slSubType, ilVefCode, slInWeekDate)
        If ilRet = True Then
            slSQL_AlertClear = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
            Set rstAlertClear = gSQLSelectCall(slSQL_AlertClear)
            Do While (Not rstAlertClear.EOF) And (Not rstAlertClear.BOF)
                DoEvents
                tgAuf.lCode = rstAlertClear!aufCode
                tgAuf.sStatus = rstAlertClear!aufStatus
                tgAuf.sSubType = rstAlertClear!aufSubType
                tgAuf.iVefCode = rstAlertClear!aufVefCode
                If IsNull(rstAlertClear!aufMoWeekDate) Then
                    tgAuf.lMoWeekDate = 0
                ElseIf Not gIsDate(rstAlertClear!aufMoWeekDate) Then
                    tgAuf.lMoWeekDate = 0
                Else
                    tgAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlertClear!aufMoWeekDate, sgShowDateForm)))
                End If
                If slType = "F" Then
                    If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                        If tgAuf.lMoWeekDate = llMoWeekDate Then
                            ilFound = True
                            Exit Do
                        End If
                    End If
                ElseIf slType = "R" Then
                    If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                        If tgAuf.lMoWeekDate = llMoWeekDate Then
                            ilFound = True
                            Exit Do
                        End If
                    End If
                End If
                rstAlertClear.MoveNext
            Loop
        End If
    End If
    
    If ilFound Then
        DoEvents
        slSQL_AlertClear = "UPDATE AUF_ALERT_USER SET "
        slSQL_AlertClear = slSQL_AlertClear & "aufStatus = 'C'" & ", "
        slSQL_AlertClear = slSQL_AlertClear & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        slSQL_AlertClear = slSQL_AlertClear & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
        slSQL_AlertClear = slSQL_AlertClear & "aufClearUstCode = " & igUstCode & " "
        slSQL_AlertClear = slSQL_AlertClear & "WHERE aufCode = " & tgAuf.lCode
        'cnn.Execute slSQL_AlertClear, rdExecDirect
        If gSQLWaitNoMsgBox(slSQL_AlertClear, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "modAlerts-gAlertClear"
            gAlertClear = False
            On Error Resume Next
            rstAlertClear.Close
            Exit Function
        End If
        gAlertClear = True
    Else
        gAlertClear = False
    End If
    DoEvents
    ilRet = gAlertForceCheck()
    rstAlertClear.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertClear"
    rstAlertClear.Close
    gAlertClear = False
    Exit Function
End Function

Public Function gAlertForceCheck() As Integer
    Dim ilRet As Integer
    
    If igExportSource = 2 Then
        gAlertForceCheck = True
        Exit Function
    End If
    If igAlertInterval <> 0 Then
        'Removed test because of Program change Alerts
        'If (sgExptSpotAlert <> "N") Or (sgExptISCIAlert <> "N") Then
            frmMain!tmcFlashAlert.Enabled = False
            frmMain!tmcCheckAlert.Enabled = False
            If igAlertInterval > 0 Then
                igAlertTimer = igAlertInterval
            End If
            gAlertForceCheck = gAlertCheck()
            igAlertFlash = -1
            Exit Function
        'End If
    End If
    frmMain!mnuAlert.Visible = False
    gAlertForceCheck = False
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gAlertFound                     *
'*                                                     *
'*             Created:5/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Test if Alert exist       *
'*                                                     *
'*******************************************************
Public Function gAlertFound(slType As String, slSubType As String, ilVefCode As Integer, slInWeekDate As String) As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim tlAuf As AUF
    Dim slMoWeekDate As String
    Dim slSQL_AlertFound As String
    
    On Error GoTo ErrHand
    
    If (slType <> "N") And (slType <> "B") Then
        slMoWeekDate = slInWeekDate
        Do While Weekday(slMoWeekDate, vbSunday) <> vbMonday
            slMoWeekDate = DateAdd("d", -1, slMoWeekDate)
        Loop
        llMoWeekDate = DateValue(gAdjYear(slMoWeekDate))
    Else
        llMoWeekDate = 0
    End If
    
    ilFound = False
    slSQL_AlertFound = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
    Set rstAlert = gSQLSelectCall(slSQL_AlertFound)
    Do While (Not rstAlert.EOF) And (Not rstAlert.BOF)
        tlAuf.sStatus = rstAlert!aufStatus
        tlAuf.sSubType = rstAlert!aufSubType
        tlAuf.iVefCode = rstAlert!aufVefCode
        If IsNull(rstAlert!aufMoWeekDate) Then
            tlAuf.lMoWeekDate = 0
        ElseIf Not gIsDate(rstAlert!aufMoWeekDate) Then
            tlAuf.lMoWeekDate = 0
        Else
            tlAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlert!aufMoWeekDate, sgShowDateForm)))
        End If
        If slType = "F" Then
            If (ilVefCode = tlAuf.iVefCode) And (slSubType = tlAuf.sSubType) Then
                If tlAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "R" Then
            If (ilVefCode = tlAuf.iVefCode) And (slSubType = tlAuf.sSubType) Then
                If tlAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "P" Then
            If (ilVefCode = tlAuf.iVefCode) And (slSubType = tlAuf.sSubType) Then
                If tlAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        End If
        rstAlert.MoveNext
    Loop
    If ilFound Then
        gAlertFound = True
    Else
        gAlertFound = False
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertFound"
    gAlertFound = False
    Exit Function
End Function

Public Function gAlertClearFinalAndReprint(slMethod As String, slType As String, slSubType As String, ilVefCode As Integer, slInWeekDate As String)
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim slTime As String
    Dim llMoWeekDate As Long
    Dim slMoWeekDate As String
    Dim slSQL_AlertClearFinalAndRep As String
    
    On Error GoTo ErrHand
    
    If (slType <> "N") And (slType <> "B") Then
        slMoWeekDate = slInWeekDate
        Do While Weekday(slMoWeekDate, vbSunday) <> vbMonday
            slMoWeekDate = DateAdd("d", -1, slMoWeekDate)
        Loop
        llMoWeekDate = DateValue(gAdjYear(slMoWeekDate))
    Else
        llMoWeekDate = 0
    End If
    
    ilFound = False
    
    If ilVefCode = 105 Then
        ilRet = ilRet
    End If
    'Old Statement
    'slSQL_AlertClearFinalAndRep = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
    'New with OR and Vehicle Code Statement
    slSQL_AlertClearFinalAndRep = "SELECT * FROM AUF_ALERT_USER WHERE aufStatus = 'R' And aufVefCode = " & ilVefCode & " AND (aufType = 'F' OR aufType = 'R')"
    Set rstAlertClearFinal = gSQLSelectCall(slSQL_AlertClearFinalAndRep)
    Do While (Not rstAlertClearFinal.EOF) And (Not rstAlertClearFinal.BOF)
        DoEvents
        tgAuf.lCode = rstAlertClearFinal!aufCode
        tgAuf.sStatus = rstAlertClearFinal!aufStatus
        tgAuf.sSubType = rstAlertClearFinal!aufSubType
        tgAuf.iVefCode = rstAlertClearFinal!aufVefCode
        If IsNull(rstAlertClearFinal!aufMoWeekDate) Then
            tgAuf.lMoWeekDate = 0
        ElseIf Not gIsDate(rstAlertClearFinal!aufMoWeekDate) Then
            tgAuf.lMoWeekDate = 0
        Else
            tgAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlertClearFinal!aufMoWeekDate, sgShowDateForm)))
        End If
        If slType = "F" Then
            If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                If tgAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        ElseIf slType = "R" Then
            If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                If tgAuf.lMoWeekDate = llMoWeekDate Then
                    ilFound = True
                    Exit Do
                End If
            End If
        End If
        rstAlertClearFinal.MoveNext
    Loop
    
    If ilFound = False Then
        'D.S. 12/09/04 - If we get here we failed to find the alert, probably because someone put in a
        'new AUF file.  If we don't have the alert set and then set the exported flag we won't know that
        'it's been exported in the past which will cause problems in future exports
        'ilRet = gAlertAdd(slType, slSubType, ilVefCode, slInWeekDate)
        ilRet = gAlertAdd("F", slSubType, ilVefCode, slInWeekDate)
        If ilRet = True Then
            'Old Statement
            'slSQL_AlertClearFinalAndRep = "SELECT * FROM AUF_ALERT_USER WHERE aufType = '" & Trim$(slType) & "' AND aufStatus = 'R'"
            'New with OR and Vehicle Code Statement
            slSQL_AlertClearFinalAndRep = "SELECT * FROM AUF_ALERT_USER WHERE aufStatus = 'R' And aufVefCode = " & ilVefCode & " AND (aufType = 'F' OR aufType = 'R')"
            
            Set rstAlertClearFinal = gSQLSelectCall(slSQL_AlertClearFinalAndRep)
            Do While (Not rstAlertClearFinal.EOF) And (Not rstAlertClearFinal.BOF)
                DoEvents
                tgAuf.lCode = rstAlertClearFinal!aufCode
                tgAuf.sStatus = rstAlertClearFinal!aufStatus
                tgAuf.sSubType = rstAlertClearFinal!aufSubType
                tgAuf.iVefCode = rstAlertClearFinal!aufVefCode
                If IsNull(rstAlertClearFinal!aufMoWeekDate) Then
                    tgAuf.lMoWeekDate = 0
                ElseIf Not gIsDate(rstAlertClearFinal!aufMoWeekDate) Then
                    tgAuf.lMoWeekDate = 0
                Else
                    tgAuf.lMoWeekDate = DateValue(gAdjYear(Format$(rstAlertClearFinal!aufMoWeekDate, sgShowDateForm)))
                End If
                If slType = "F" Then
                    If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                        If tgAuf.lMoWeekDate = llMoWeekDate Then
                            ilFound = True
                            Exit Do
                        End If
                    End If
                ElseIf slType = "R" Then
                    If (ilVefCode = tgAuf.iVefCode) And (slSubType = tgAuf.sSubType) Then
                        If tgAuf.lMoWeekDate = llMoWeekDate Then
                            ilFound = True
                            Exit Do
                        End If
                    End If
                End If
                rstAlertClearFinal.MoveNext
            Loop
        End If
    End If
    
    If ilFound Then
        DoEvents
        slSQL_AlertClearFinalAndRep = "UPDATE AUF_ALERT_USER SET "
        slSQL_AlertClearFinalAndRep = slSQL_AlertClearFinalAndRep & "aufStatus = 'C'" & ", "
        slSQL_AlertClearFinalAndRep = slSQL_AlertClearFinalAndRep & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        slSQL_AlertClearFinalAndRep = slSQL_AlertClearFinalAndRep & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
        slSQL_AlertClearFinalAndRep = slSQL_AlertClearFinalAndRep & "aufClearUstCode = " & igUstCode & " "
        slSQL_AlertClearFinalAndRep = slSQL_AlertClearFinalAndRep & "WHERE aufCode = " & tgAuf.lCode
        'cnn.Execute slSQL_AlertClearFinalAndRep, rdExecDirect
        If gSQLWaitNoMsgBox(slSQL_AlertClearFinalAndRep, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "modAlerts-gAlertClearFinalAndReprint"
            gAlertClearFinalAndReprint = False
            On Error Resume Next
            rstAlertClearFinal.Close
            Exit Function
        End If
        gAlertClearFinalAndReprint = True
    Else
        gAlertClearFinalAndReprint = False
    End If
    DoEvents
''''    ilRet = gAlertForceCheck()
    rstAlertClearFinal.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertClearFinalAndReprint"
    rstAlertClearFinal.Close
    gAlertClearFinalAndReprint = False
    Exit Function
End Function

Public Function gAlertCheckBlock(tlAuf As AUF) As Integer
    
    Dim ilRet As Integer
    Dim slSQL_AlertCheckBlock As String
    
    On Error GoTo ErrHand
    
    slSQL_AlertCheckBlock = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'B' AND aufStatus = 'R'"
    Set rstAlert = gSQLSelectCall(slSQL_AlertCheckBlock)
    If (Not rstAlert.EOF) And (Not rstAlert.BOF) Then
        If (rstAlert!aufSubType = "A") Or (rstAlert!aufSubType = "B") Then
            tlAuf.iCountdown = rstAlert!aufcountdown
            tlAuf.lCefCode = rstAlert!aufcefcode
            tlAuf.lEnteredTime = gTimeToLong(Format$(rstAlert!aufEnteredTime, sgShowTimeWOSecForm), False)
            tlAuf.iCreateUstCode = rstAlert!aufCreateUstCode
            If (tlAuf.iCreateUstCode <> igUstCode) Then
                gAlertCheckBlock = 1    'Block user
                Exit Function
            Else
                gAlertCheckBlock = 2    'User that initiate block
                Exit Function
            End If
        End If
    End If
    gAlertCheckBlock = 0    'No blocks
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertCheckBlock"
    gAlertCheckBlock = 0
    Exit Function
End Function

Public Function gAlertCheckNotice(llUlfCode As Long, llCefCode As Long) As Long
    
    Dim ilRet As Integer
    Dim slSQL_AlertCheckNotice As String
    
    slSQL_AlertCheckNotice = "SELECT * FROM AUF_ALERT_USER WHERE aufType = 'N' AND aufStatus = 'R'"
    Set rstAlert = gSQLSelectCall(slSQL_AlertCheckNotice)
    Do While (Not rstAlert.EOF) And (Not rstAlert.BOF)
        llCefCode = rstAlert!aufcefcode
        If rstAlert!aufulfcode = llUlfCode Then
            gAlertCheckNotice = rstAlert!aufCode
            Exit Function
        End If
        SQLAlertULF = "SELECT * FROM ULF_USER_LOG WHERE ulfCode = " & rstAlert!aufulfcode
        Set rstAlertUlf = gSQLSelectCall(SQLAlertULF)
        If (Not rstAlertUlf.EOF) And (Not rstAlertUlf.BOF) Then
            If igUstCode = rstAlertUlf!ulfUstCode Then
                gAlertCheckNotice = rstAlert!aufCode
                Exit Function
            End If
        End If
        rstAlert.MoveNext
    Loop
    gAlertCheckNotice = 0
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAllerts-gAlertCheckNotice"
    gAlertCheckNotice = 0
    Exit Function
End Function
