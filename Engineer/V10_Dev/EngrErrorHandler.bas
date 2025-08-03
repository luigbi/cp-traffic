Attribute VB_Name = "EngrErrorHandler"

Option Explicit

Public sgCallStack(0 To 9) As String


'Differences between Traffic ErrorHandle and EngrErrorHandle:
'  1.  Replaced sgExePath with sgExeDirectory
'  2.  Removed reference to tgUrf(0)


Public Sub gDbg_HandleError(ModuleAndFunction As String)
    Dim slMsg As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilTo As Integer
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    Dim slAppName As String
    Dim ilPos As Integer
    Dim ilRet As Integer

    ' Get the error information now to preserve it.
    'It must be prior to Resume Next as that clears the values
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description

    On Error Resume Next

    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"

    slDateTime = Format$(Now(), "ddddd h:mm:ssAM/PM")
    slMsg = slDateTime & vbCrLf & vbCrLf & "Module: " & ModuleAndFunction & vbCrLf & _
          "Line No: " & ilLine & vbCrLf & _
          "Error: " & Str(ilErrNo) & vbCrLf & _
          "Desc: " & slDesc & vbCrLf & vbCrLf & _
          slAppName & ": " & Format$(FileDateTime(sgExeDirectory & slAppName), "ddddd") & " at " & Format$(FileDateTime(sgExeDirectory & slAppName), "ttttt") & vbCrLf & vbCrLf & _
          "The System will now shut down."

    'ilRet = gDeleteLockRec_ByUser(tgUrf(0).iCode)

    If igBkgdProg = 0 Then
        MsgBox slMsg, vbCritical, "Application Error"
    End If

    ' Reformat and Log the error message as well
    slMsg = slDateTime & _
          ", Module: " & ModuleAndFunction & _
          ", Line No: " & ilLine & _
          ", Error: " & Str(ilErrNo) & _
          ", Desc: " & slDesc & _
          ", " & slAppName & ": " & Format$(FileDateTime(sgExeDirectory & slAppName), "ddddd") & " at " & Format$(FileDateTime(sgExeDirectory & slAppName), "ttttt")

'    If igBkgdProg = 0 Then
''        slToFile = sgDBPath & "Messages\TrafficErrors.Txt"
''        ilTo = FreeFile
''        Open slToFile For Append As ilTo
''        Print #ilTo, slMsg
''        Close #ilTo
'        gLogMsg slMsg, "TrafficErrors.Txt", False
'    ElseIf igBkgdProg = 1 Then
'        gLogMsg slMsg, "Bkgd_Schd.Txt", False
'    ElseIf igBkgdProg = 2 Then
'        gLogMsg slMsg, "Set_Credit.Txt", False
'    Else
'        gLogMsg slMsg, "TrafficErrors.Txt", False
'    End If
    gMsgBox slMsg, -1, "Error Handle"

'    slAppName = App.EXEName
'    If InStr(1, slAppName, ".", 1) > 0 Then
'        slAppName = Left$(slAppName, ilPos - 1)
'    End If

    'Unload Traffic
    btrStopAppl
    End
End Sub



Public Sub gErrorApplStop(ModuleAndFunction As String)
    Dim slMsg As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilTo As Integer
    Dim slAppName As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    On Error Resume Next

    slAppName = App.EXEName
    ilPos = InStr(1, slAppName, ".", 1)
    If ilPos > 0 Then
        slAppName = Left$(slAppName, ilPos - 1)
    End If
    slAppName = slAppName & ".exe"

    slDateTime = Format$(Now(), "ddddd h:mm:ssAM/PM")
    slMsg = slDateTime & vbCrLf & vbCrLf & "Module: " & ModuleAndFunction & vbCrLf & _
          slAppName & ": " & Format$(FileDateTime(sgExeDirectory & slAppName), "ddddd") & " at " & Format$(FileDateTime(sgExeDirectory & slAppName), "ttttt") & vbCrLf & vbCrLf & _
          "The System will now shut down."
    
    'ilRet = gDeleteLockRec_ByUser(tgUrf(0).iCode)

    If igBkgdProg = 0 Then
        MsgBox slMsg, vbCritical, "Application Error"
    End If
    
    ' Reformat and Log the error message as well
    slMsg = slDateTime & _
          ", Module: " & ModuleAndFunction & _
          ", " & slAppName & ": " & Format$(FileDateTime(sgExeDirectory & slAppName), "ddddd") & " at " & Format$(FileDateTime(sgExeDirectory & slAppName), "ttttt")

'    If igBkgdProg = 0 Then
''        slToFile = sgDBPath & "Messages\TrafficErrors.Txt"
''        ilTo = FreeFile
''        Open slToFile For Append As ilTo
''        Print #ilTo, slMsg
''        Close #ilTo
'        gLogMsg slMsg, "TrafficErrors.Txt", False
'    ElseIf igBkgdProg = 1 Then
'        gLogMsg slMsg, "Bkgd_Schd.Txt", False
'    ElseIf igBkgdProg = 2 Then
'        gLogMsg slMsg, "Set_Credit.Txt", False
'    Else
'        gLogMsg slMsg, "TrafficErrors.Txt", False
'    End If
    gMsgBox slMsg, -1, "Error Handle"
    btrStopAppl
    End

End Sub

Public Sub gAddCallToStack(slCallToAdd As String)
    Dim ilLoop As Integer
    
    For ilLoop = UBound(sgCallStack) To LBound(sgCallStack) + 1 Step -1
        sgCallStack(ilLoop) = sgCallStack(ilLoop - 1)
    Next ilLoop
    sgCallStack(LBound(sgCallStack)) = slCallToAdd
End Sub

Public Sub gRemoveCallFromStack()
    Dim ilLoop As Integer
    
    For ilLoop = LBound(sgCallStack) To UBound(sgCallStack) - 1 Step 1
        sgCallStack(ilLoop) = sgCallStack(ilLoop + 1)
    Next ilLoop
    sgCallStack(UBound(sgCallStack)) = ""
End Sub

Public Sub gSaveStackTrace(slLogFileName As String)
    Dim ilTo As Integer
    Dim ilLoop As Integer
    Dim slMethodName As String
    
    ilTo = FreeFile
    Open slLogFileName For Append As ilTo
    Print #ilTo, "Call Stack Trace"
    Print #ilTo, "----------------"
    For ilLoop = LBound(sgCallStack) To UBound(sgCallStack)
        slMethodName = sgCallStack(ilLoop)
        If Len(slMethodName) > 0 Then
            Print #ilTo, slMethodName
        End If
    Next ilLoop
    Print #ilTo, "----------------"
    Close #ilTo
End Sub

Public Sub gUserActivityLog(slFunction As String, slInName As String)
    'Added so that the latest CodeProtect can be used
    'User activity will not be mantained within the Engineering project at this time
End Sub
