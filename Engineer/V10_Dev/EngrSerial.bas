Attribute VB_Name = "EngrSerial"
'
' Release: 1.0
'
' Description:
'   This file contains the SerialPort Communication declarations
Option Explicit
Private smEndCodes() As String

Public Function gOpenPort(Ctrl As MSComm, tlITE As ITE) As Integer
    Dim slStr As String
    Dim slMsg As String
    
    On Error GoTo gOpenPortErr:
    Ctrl.CommPort = 1
    slStr = Trim$(Str$(tlITE.iBaud))
    slStr = slStr & "," & tlITE.sParity
    slStr = slStr & "," & Trim$(Str$(tlITE.iDataBits))
    slStr = slStr & "," & tlITE.sStopBit
    Ctrl.Settings = slStr
    Ctrl.InputLen = 0
    Ctrl.PortOpen = True
    gOpenPort = True
    Exit Function
gOpenPortErr:
    If (Err.Number <> 0) Then
        slMsg = "A general error has occured in Serial Port- gOpenPort: " & Err.Description & "; Error #" & Err.Number
        gShowPortErrorMsg slMsg
    End If
    gOpenPort = False
End Function

Public Function gClosePort(Ctrl As MSComm) As Integer
    Dim slMsg As String
    
    On Error GoTo gClosePortErr:
    Ctrl.PortOpen = False
    gClosePort = True
    Exit Function
gClosePortErr:
    If (Err.Number <> 0) Then
        slMsg = "A general error has occured in Serial Port- gClosePort: " & Err.Description & "; Error #" & Err.Number
        gShowPortErrorMsg slMsg
    End If
    gClosePort = False
    Exit Function
End Function

Public Function gPutPort(Ctrl As MSComm, slInMsg As String) As Integer
    Dim slSendMsg As String
    Dim slMsg As String
    
    On Error GoTo gPutPortErr:
    slSendMsg = slInMsg
    Ctrl.Output = slSendMsg
    gPutPort = True
    Exit Function
gPutPortErr:
    If (Err.Number <> 0) Then
        slMsg = "A general error has occured in Serial Port- gPutPort: " & Err.Description & "; Error #" & Err.Number
        gShowPortErrorMsg slMsg
    End If
    gPutPort = False
    Exit Function
End Function

Public Function gGetPort(Ctrl As MSComm, slEndStr As String, slGetPattern As String, slErrChk As String, slCheckSum As String, ilShowErrorMsg As Integer, slReturnedMsg As String) As Integer
    Dim slBuffer As String
    Dim slRetCheckSum As String
    Dim slTestCheckSum As String
    Dim slMsg As String
    Dim llTime As Long
    
    On Error GoTo gGetPortErr
    llTime = gTimeToLong(Time$(), False) + gTimeToLong("00:02:00", False)
    Do
        slBuffer = ""
        Do
            DoEvents
            If llTime < gTimeToLong(Time$(), False) Then
                gGetPort = False
                If ilShowErrorMsg Then
                    slMsg = "Serial Port- gGetPort: Time out- No Response for 2 minutes"
                    gShowPortErrorMsg slMsg
                End If
                slReturnedMsg = "Failed"
                Exit Function
            End If
            slBuffer = slBuffer & Ctrl.Input
        Loop Until InStr(1, slBuffer, slEndStr, vbTextCompare) >= 1
    Loop While (InStr(1, slBuffer, slGetPattern, vbTextCompare) <> 1) And (InStr(1, slBuffer, slErrChk, vbTextCompare) <> 1)
    If slCheckSum <> "N" Then
        slRetCheckSum = Mid$(slBuffer, Len(slBuffer) - Len(slEndStr) - 1, 2)
        slBuffer = Left$(slBuffer, Len(slBuffer) - Len(slEndStr) - 2)
        slTestCheckSum = gIntToHex(gCheckSum(slBuffer), 2)
        If StrComp(slRetCheckSum, slTestCheckSum, vbTextCompare) <> 0 Then
            gGetPort = False
        Else
            If InStr(1, slBuffer, slErrChk, vbTextCompare) <> 1 Then
                gGetPort = True
            Else
                gGetPort = False
                slBuffer = "Failed- " & Mid$(slBuffer, Len(slErrChk) + 1)
            End If
        End If
    Else
        slBuffer = Left$(slBuffer, Len(slBuffer) - Len(slEndStr))
        If InStr(1, slBuffer, slErrChk, vbTextCompare) <> 1 Then
            gGetPort = True
        Else
            gGetPort = False
            slBuffer = "Failed- " & Mid$(slBuffer, Len(slErrChk) + 1)
        End If
    End If
    slReturnedMsg = Trim$(slBuffer)
    Exit Function
gGetPortErr:
    If (Err.Number <> 0) Then
        slMsg = "A general error has occured in Serial Port- gGetPort: " & Err.Description & "; Error #" & Err.Number
        gShowPortErrorMsg slMsg
    End If
    slReturnedMsg = "Failed"
    gGetPort = False
    Exit Function
End Function

Public Sub gErrorMsgPort(Ctrl As MSComm)
    Dim slMsg As String
    
    Select Case Ctrl.CommEvent
        'Event Errors
        Case comEventBreak
            slMsg = "Communication Port Error: " & "A Break signal was received, Error 1001"
        Case comEventCDTO
            slMsg = "Communication Port Error: " & "Carrier Detect Timeout, Error 1007"
        Case comEventCTSTO
            slMsg = "Communication Port Error: " & "Clear to Send Timeout, Error 1002"
        Case comEventDSRTO
            slMsg = "Communication Port Error: " & "Data Set Ready Timeout, Error 1003"
        Case comEventFrame
            slMsg = "Communication Port Error: " & "Frame Error, Error 1004"
        Case comEventOverrun
            slMsg = "Communication Port Error: " & "Port Overrun, Error 1006"
        Case comEventRxOver
            slMsg = "Communication Port Error: " & "Receive Buffer Overrun, Error 1008"
        Case comEventRxParity
            slMsg = "Communication Port Error: " & "Parity Error, Error 1009"
        Case comEventTxFull
            slMsg = "Communication Port Error: " & "Transmit Buffer Full, Error 1010"
        Case comEventDCB
            slMsg = "Communication Port Error: " & "Unexpected error retrieving Device Control Block, Error 1011"
        Case Else
            Exit Sub
    End Select
        
    gShowPortErrorMsg slMsg
        
End Sub

Private Sub mBuildItemIDQuery(slItemID As String, tlITE As ITE, slEndStr As String, slTitleQuery As String, slLengthQuery As String, slReplyChk As String, slErrChk As String)
    Dim ilQuery As String
    Dim slMessageNo As String
    Dim slMaxNo As String
    Dim ilMessageID As Integer
    Dim ilRet As Integer
    Dim slQuery As String
    
    slQuery = ""
    slQuery = slQuery & Trim$(tlITE.sStartCode)
    slQuery = slQuery & Trim$(tlITE.sMachineID)
    ilRet = gGetItemMsgID(tlITE.iCode, "Get Item ID next Message ID", ilMessageID)
    tlITE.iCurrMgsID = ilMessageID
    slMessageNo = Trim$(Str$(ilMessageID))
    slMaxNo = Trim$(Str$(tlITE.iMaxMgsID))
    Do While Len(slMessageNo) < Len(slMaxNo)
        slMessageNo = "0" & slMessageNo
    Loop
    slQuery = slQuery & slMessageNo
    slErrChk = Trim$(tlITE.sReplyCode) & Mid$(slQuery, Len(Trim$(tlITE.sStartCode)) + 1) & Trim$(tlITE.sMgsErrType)
    slQuery = slQuery & Trim$(tlITE.sMgsType)
    slReplyChk = Trim$(tlITE.sReplyCode) & Mid$(slQuery, Len(Trim$(tlITE.sStartCode)) + 1)

    slQuery = slQuery & slItemID
    If Trim$(tlITE.sTitleID) <> "-" Then
        slTitleQuery = slQuery & Trim$(tlITE.sTitleID)
        If tlITE.sCheckSum <> "N" Then
            slTitleQuery = slTitleQuery & gIntToHex(gCheckSum(slTitleQuery), 2)
        End If
    Else
        slTitleQuery = ""
    End If
    If Trim$(tlITE.sLengthID) <> "-" Then
        slLengthQuery = slQuery & Trim$(tlITE.sLengthID)
        If tlITE.sCheckSum <> "N" Then
            slLengthQuery = slLengthQuery & gIntToHex(gCheckSum(slLengthQuery), 2)
        End If
    Else
        slLengthQuery = ""
    End If
    If slTitleQuery <> "" Then
        slTitleQuery = slTitleQuery & slEndStr
    End If
    If slLengthQuery <> "" Then
        slLengthQuery = slLengthQuery & slEndStr
    End If
End Sub

Public Function gTestItemID(spcItemID As MSComm, tlITE As ITE, slItemID As String, ilDoConnectTest As Integer, slTitle As String, slLength As String) As Integer
    Dim ilRet As Integer
    Dim ilTRet As Integer
    Dim slReturnStr As String
    Dim slStr As String
    Dim slEndStr As String
    Dim slMsg As String
    Dim ilLoop As Integer
    Dim slTitleQuery As String
    Dim slLengthQuery As String
    Dim slConnectSeq As String
    Dim slReplyChk As String
    Dim slErrChk As String
    Dim slConnectSeqReplyChk As String
    Dim slConnectSeqErrChk As String
    
    slTitle = ""
    slLength = ""
    slStr = tlITE.sMgsEndCode
    If Trim$(slStr) <> "" Then
        gParseCDFields slStr, False, smEndCodes()
        slEndStr = ""
        For ilLoop = LBound(smEndCodes) To UBound(smEndCodes) Step 1
            slStr = Trim$(smEndCodes(ilLoop))
            If StrComp(slStr, "cr", vbTextCompare) = 0 Then
                slEndStr = slEndStr & Chr$(13)
            Else
                If StrComp(slStr, "lf", vbTextCompare) = 0 Then
                    slEndStr = slEndStr & Chr$(10)
                Else
                    If StrComp(slStr, "Tab", vbTextCompare) = 0 Then
                        slEndStr = slEndStr & Chr$(9)
                    Else
                        If (Asc(slStr) >= Asc("0")) And (Asc(slStr) <= Asc("9")) Then
                            slEndStr = slEndStr & Chr$(slStr)
                        Else
                            slEndStr = slEndStr & slStr
                        End If
                    End If
                End If
            End If
        Next ilLoop
    End If
    mBuildItemIDQuery slItemID, tlITE, slEndStr, slTitleQuery, slLengthQuery, slReplyChk, slErrChk
    gTestItemID = False
    ilRet = gOpenPort(spcItemID, tlITE)
    If ilRet Then
        If (Trim$(tlITE.sConnectSeq) <> "") And (ilDoConnectTest) Then
            slConnectSeq = ""
            slConnectSeq = slConnectSeq & Trim$(tlITE.sStartCode)
            slConnectSeq = slConnectSeq & Trim$(tlITE.sMachineID)
            slConnectSeq = slConnectSeq & Trim$(tlITE.sConnectSeq)
            slConnectSeqReplyChk = Trim$(tlITE.sReplyCode) & Mid$(slConnectSeq, Len(Trim$(tlITE.sStartCode)) + 1)
            slConnectSeqErrChk = Trim$(tlITE.sReplyCode) & Trim$(tlITE.sMachineID) & Left$(Trim$(tlITE.sConnectSeq), Len(Trim$(tlITE.sConnectSeq)) - 1) & Trim$(tlITE.sMgsErrType)
            slConnectSeq = slConnectSeq & gIntToHex(gCheckSum(slConnectSeq), 2) & slEndStr
            ilRet = gPutPort(spcItemID, slConnectSeq)
            If ilRet Then
                ilRet = gGetPort(spcItemID, slEndStr, slConnectSeqReplyChk, slConnectSeqErrChk, tlITE.sCheckSum, True, slReturnStr)
                If (Not ilRet) Or (slReturnStr = "Failed") Then
                    gTestItemID = False
                    slTitle = "Failed"
                    ilRet = gClosePort(spcItemID)
                    Exit Function
                End If
            Else
                gTestItemID = False
                slTitle = "Failed"
                ilRet = gClosePort(spcItemID)
                Exit Function
            End If
        End If
        ilTRet = True
        If slTitleQuery <> "" Then
            ilTRet = gPutPort(spcItemID, slTitleQuery)
            If ilTRet Then
                If (Trim$(tlITE.sConnectSeq) <> "") Then
                    ilTRet = gGetPort(spcItemID, slEndStr, slReplyChk, slErrChk, tlITE.sCheckSum, False, slReturnStr)
                Else
                    ilTRet = gGetPort(spcItemID, slEndStr, slReplyChk, slErrChk, tlITE.sCheckSum, True, slReturnStr)
                End If
                If ilTRet Then
                    slTitle = Mid$(slReturnStr, Len(slReplyChk) + 1)
                    gTestItemID = True
                Else
                    slTitle = slReturnStr
                End If
            Else
                slTitle = "Failed"
            End If
        End If
        If slLengthQuery <> "" Then
            ilRet = gPutPort(spcItemID, slLengthQuery)
            If ilRet Then
                If (Trim$(tlITE.sConnectSeq) <> "") Then
                    ilRet = gGetPort(spcItemID, slEndStr, slReplyChk, slErrChk, tlITE.sCheckSum, False, slReturnStr)
                Else
                    ilRet = gGetPort(spcItemID, slEndStr, slReplyChk, slErrChk, tlITE.sCheckSum, True, slReturnStr)
                End If
                slLength = slReturnStr
                If ilRet Then
                    slLength = Mid$(slReturnStr, Len(slReplyChk) + 1)
                    If ilTRet Then
                        gTestItemID = True
                    End If
                Else
                    If ilTRet Then
                        gTestItemID = False
                    End If
                End If
            Else
                slLength = "Failed"
                If ilTRet Then
                    gTestItemID = False
                End If
            End If
        End If
        ilRet = gClosePort(spcItemID)
    Else
        slTitle = "Failed"
        slLength = "Failed"
    End If

End Function

Public Sub gShowPortErrorMsg(slMsg As String)
    If igOperationMode = 1 Then
        gLogMsg slMsg, "CommPort_Server.txt", False
    Else
        gLogMsg slMsg, "CommPort_Client.txt", False
        MsgBox slMsg, vbCritical
    End If

End Sub
