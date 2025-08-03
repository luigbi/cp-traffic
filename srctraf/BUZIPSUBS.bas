Attribute VB_Name = "BUZIPSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of BUZIPSUBS.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Constants (Marked)                                                              *
'*  NO_ACTION                                                                             *
'******************************************************************************************

Option Explicit
Public sgFileAttachment As String
Public sgFileAttachmentName As String
Public sgContDBPath As String
Public igBUerror As Integer
Private hmBUMsg As Integer
Public igZipCancel As Integer
Global Const NO_ACTION = 0 'VBC NR

'*********************************************************************************
'
'*********************************************************************************
Public Function gContDBPathResolved() As Integer

    'If the database is on a network then ServerDatabase must be defined
    'in the Traffic.ini file.  If not then no backups
    If sgServerDatabase = "" And btrIsANetworkPath(sgDBPath) = True Then
        MsgBox "ServerDatabase is not defined in Traffic.ini file." & Chr(10) & Chr(13) _
               & "The database path must be exactly as if sitting at the server." _
               & Chr(10) & Chr(13) & "Example: ServerDatabase = C:\Csi\Prod\Data"
        gContDBPathResolved = False
    End If

    'We have a standard network configuration
    If sgServerDatabase <> "" And btrIsANetworkPath(sgDBPath) = True Then
        sgContDBPath = sgServerDatabase
        gContDBPathResolved = True
    End If

    'We have a stand alone machine; does not matter if sgServerDatabase is defined or not
    If btrIsANetworkPath(sgDBPath) = False Then
        sgContDBPath = sgDBPath
        gContDBPathResolved = True
    End If

End Function

Public Sub gOpenBUMsgFile(sErrStr As String)

    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer

    'On Error GoTo gOpenBUMsgFileErr:
    slToFile = sgDBPath & "Messages\" & "Backup.Txt"
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        'If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo gOpenBUMsgFileErr:
            'hmBUMsg = FreeFile
            'Open slToFile For Append As hmBUMsg
            ilRet = gFileOpen(slToFile, "Append", hmBUMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
        'End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo gOpenBUMsgFileErr:
        'hmBUMsg = FreeFile
        'Open slToFile For Output As hmBUMsg
        ilRet = gFileOpen(slToFile, "Output", hmBUMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Print #hmBUMsg, "** " & sErrStr & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
    Close hmBUMsg
    Exit Sub
'gOpenBUMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Sub

'*********************************************************************************
'
'*********************************************************************************
Public Sub gCheckForContFiles()
    Dim ilRet As Integer
    Dim slMsg As String
    Dim imSalesperson As Boolean
    Dim slLastBackupDateTime As String
    Dim slCurDateTime As String
    Dim ilLoop As Integer
    Dim llTotalHours As Long
    Dim ilValue As Integer
    Dim SvrRsp_FilesStuckInCntMode As CSISvr_Rsp_Answer

    On Error GoTo ErrHand
    If (Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName) Then
        Exit Sub
    End If
    ' New Code
    ' Remove all code that lets a person know they are the backup person.
    ' Also Get rid of all backup related code.
    imSalesperson = False
    If (tgUrf(0).iSlfCode > 0) Then
        For ilLoop = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgMSlf(ilLoop).iCode = tgUrf(0).iSlfCode Then
                If StrComp(tgMSlf(ilLoop).sJobTitle, "S", 1) = 0 Then
                    imSalesperson = True
                End If
                Exit For
            End If
        Next ilLoop
    End If
    If imSalesperson Then
        Exit Sub
    End If

    ilRet = csiCheckForFilesStuckInCntMode(sgDBPath, SvrRsp_FilesStuckInCntMode)
    If SvrRsp_FilesStuckInCntMode.iAnswer = 1 Then
        slMsg = ""
        slMsg = slMsg & "<<< WARNING >>>" & vbCrLf
        slMsg = slMsg & "<<< Your last database back failed. >>>" & vbCrLf & vbCrLf
        slMsg = slMsg & "Adding to or editing information while in this condition could result in data loss and/or data corruption." & vbCrLf & vbCrLf
        slMsg = slMsg & "Although you may continue to view information, it is imperative that you call Counterpoint or email Counterpoint at service@counterpoint.net ASAP to remedy this condition." & vbCrLf
        gMsgBox slMsg, vbCritical, "Backup Failure"
        gLogMsg "User was warned that files are stuck in continuous mode.", "TrafficErrors.Txt", False
        Exit Sub
    End If
    ilValue = Asc(tgSpf.sUsingFeatures7)
    If (ilValue And CSIBACKUP) <> CSIBACKUP Then
        ' CSI Backups are not turned on.
        Exit Sub
    End If

    slLastBackupDateTime = gGetLastBackupDateTime()
    slCurDateTime = gNow()
    llTotalHours = DateDiff("h", slLastBackupDateTime, slCurDateTime)
    If llTotalHours > 24 Then
        slMsg = ""
        slMsg = slMsg & "<<< WARNING >>>" & vbCrLf
        slMsg = slMsg & "<<< A database backup has not occurred in over 24 hours.  >>>" & vbCrLf & vbCrLf
        gMsgBox slMsg, vbExclamation, "Backup Notice"
        gLogMsg "User was warned that backup has not been performed within 24 hours.", "TrafficErrors.Txt", False
    End If
    Exit Sub

ErrHand:
    gLogMsg "A general error occured in gCheckForContFiles", "TrafficErrors.Txt", False
    Resume Next
End Sub

'*********************************************************************************
'
'*********************************************************************************
Public Function gGetLastBackupDateTime() As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  ErrHand                                                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slDateTime As String
    Dim SvrRsp_GetLastBackupDate As CSISvr_Rsp_GetLastBackupDate

    gGetLastBackupDateTime = ""
    ilRet = csiGetLastBackupDate(sgDBPath, SvrRsp_GetLastBackupDate)
    slDateTime = SvrRsp_GetLastBackupDate.sLastBackupDateTime
    gGetLastBackupDateTime = gRemoveIllegalChars(slDateTime)
    Exit Function

ErrHand: 'VBC NR
    MsgBox "A general error has occured in gGetLastBackupDateTime."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gGetLastCopyDateTime() As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  ErrHand                                                                               *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slDateTime As String
    Dim SvrRsp_GetLastCopyDate As CSISvr_Rsp_GetLastBackupDate

    gGetLastCopyDateTime = ""
    ilRet = csiGetLastCopyDate(sgDBPath, SvrRsp_GetLastCopyDate)
    slDateTime = SvrRsp_GetLastCopyDate.sLastBackupDateTime
    gGetLastCopyDateTime = gRemoveIllegalChars(slDateTime)
    Exit Function

ErrHand: 'VBC NR
    MsgBox "A general error has occured in gGetLastCopyDateTime."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gIsBackupRunning() As Boolean
    Dim ilRet As Integer
    Dim SvrRsp_IsBackupRunning As CSISvr_Rsp_Answer

    On Error GoTo ErrHand

    gIsBackupRunning = False
    ilRet = csiIsBackupRunning(sgDBPath, SvrRsp_IsBackupRunning)
    If ilRet <> 0 Then
        Exit Function
    End If
    If SvrRsp_IsBackupRunning.iAnswer = 1 Then
        gIsBackupRunning = True
    End If
    Exit Function

ErrHand:
    MsgBox "A general error has occured in IsBackupRunning."
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gLoadINIValue(sPathFileName As String, Section As String, Key As String, sValue As String) As Boolean
    On Error GoTo ErrHand
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128

    gLoadINIValue = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, sPathFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            gLoadINIValue = True
        End If
    End If
    Exit Function

ErrHand:
    ' return now if an error occurs
End Function

'*********************************************************************************
'
'*********************************************************************************
Public Function gSaveINIValue(sPathFileName As String, Section As String, Key As String, sValue As String) As Boolean
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  BytesCopied                                                                           *
'******************************************************************************************

    On Error GoTo ErrHand

    gSaveINIValue = False
    If WritePrivateProfileString(Section, Key, sValue, sPathFileName) Then
        gSaveINIValue = True
    End If
    Exit Function

ErrHand:
    ' return now if an error occurs
End Function



