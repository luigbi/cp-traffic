Attribute VB_Name = "Comm"
'******************************************************
'*
'*  Created March,2014 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private tmCsiFtpInfo As CSIFTPINFO
Private tmCsiFtpStatus As CSIFTPSTATUS
Private tmCsiFtpErrorInfo As CSIFTPERRORINFO
Private imFTPEvents As Boolean
Private imFtpInProgress As Boolean
Public gFtpArray() As String
Private FTPIsOn As Integer
Private smIniPath As String
Private smCurDir As String
Private smTemp As String
Private imError As Integer
Private imBackGround As Integer


Public Function gTestFTP() As Boolean

    Dim ilLoop As Integer
    
    ReDim gFtpArray(0 To 7)
    For ilLoop = 0 To 7 Step 1
        gFtpArray(ilLoop) = "FtpFileTest_" & ilLoop + 1 & ".txt"
    Next ilLoop
    ilLoop = gFTPMain(False)

End Function

Public Function gFTPMain(iBackGround As Integer) As Boolean

    Dim slFTPIsOn As String
    Dim ilRet As Integer
    Dim ilIdx As Integer

    gFTPMain = False
    'If iBackGround = True then don't show error messages to the user. Program is running unattended
    imBackGround = iBackGround
    ilRet = mReadIniFile("[FTP_INFO]")
    If ilRet Then
        Call mLoadOption("FTP_INFO", "FTPIsOn", slFTPIsOn, smIniPath)
        If slFTPIsOn = "1" Then
            ilRet = gInitFTP()
            If ilRet Then
                gLogMsg "**** Starting FTP Process", "FTPLog.txt", False
                For ilIdx = 0 To UBound(gFtpArray) Step 1
                    imFtpInProgress = True
                    ilRet = gFTPSendFiles(Trim$(gFtpArray(ilIdx)))
                    imError = False
                    While imFtpInProgress And Not imError
                        Sleep (250)
                        DoEvents
                        ilRet = mCheckFTPStatus(gFtpArray(ilIdx))
                    Wend
                    If imError Then
                        Exit For
                    End If
                    
                Next ilIdx
            End If
        End If
        If imError Then
            mDeleteUselessFile tmCsiFtpInfo.sLogPathName
            gLogMsg "**** Ending FTP Process ****", "FTPLog.txt", False
            gLogMsg "", "FTPLog.txt", False
            Exit Function
        End If
        If slFTPIsOn = "1" And ilRet Then
            mDeleteUselessFile tmCsiFtpInfo.sLogPathName
            gLogMsg "**** Ending FTP Process", "FTPLog.txt", False
            gLogMsg "", "FTPLog.txt", False
        End If
    End If
    gFTPMain = True
End Function

Public Function gInitFTP() As Boolean

    Dim slTemp As String
    Dim ilRet As Integer
    
    gInitFTP = False
    'Port Number
    If Not mLoadOption("FTP_INFO", "FTPPort", slTemp, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: FTPPort is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: FTPPort is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    tmCsiFtpInfo.nPort = CInt(slTemp)
    tgCsiFtpFileListing.nPort = CInt(slTemp)
    
    'FTP Address
    If Not mLoadOption("FTP_INFO", "FTPAddress", tmCsiFtpInfo.sIPAddress, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: FTPAddress is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: FTPAddress is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    ilRet = mLoadOption("FTP_INFO", "FTPAddress", tgCsiFtpFileListing.sIPAddress, smIniPath)
    
    'FTP User ID
    If Not mLoadOption("FTP_INFO", "FTPUID", tmCsiFtpInfo.sUID, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: FTPUID is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: FTPUID is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    ilRet = mLoadOption("FTP_INFO", "FTPUID", tgCsiFtpFileListing.sUID, smIniPath)
    
    'FTP PAssword
    If Not mLoadOption("FTP_INFO", "FTPPWD", tmCsiFtpInfo.sPWD, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: FTPPWD is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: FTPPWD is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    ilRet = mLoadOption("FTP_INFO", "FTPPWD", tgCsiFtpFileListing.sPWD, smIniPath)
    
    'Local folder where files to FTP
    If Not mLoadOption("FTP_INFO", "LocalExports", tmCsiFtpInfo.sSendFolder, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: LocalExports is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: LocalExports is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    
    'Location of import folder on FTP site
    If Not mLoadOption("FTP_INFO", "FTPImportDir", tmCsiFtpInfo.sServerDstFolder, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [FTPINFO] the value: FTPImportDir is not defined"
        End If
        Exit Function
        gLogMsg "In the Traffic.ini file under [FTPINFO] the value: FTPImportDir is not defined", "FTPLog.txt.Txt", False
    End If
    ilRet = mLoadOption("FTP_INFO", "FTPImportDir", tgCsiFtpFileListing.sPathFileMask, smIniPath)
    
    'Database location on local machine
    If Not mLoadOption("Locations", "Database", tmCsiFtpInfo.sLogPathName, smIniPath) Then
        If Not imBackGround Then
            MsgBox "In the Traffic.ini file under [Locations] the value: Database is not defined"
        End If
        gLogMsg "In the Traffic.ini file under [Locations] the value: Database is not defined", "FTPLog.txt.Txt", False
        Exit Function
    End If
    ilRet = mLoadOption("Locations", "Database", tgCsiFtpFileListing.sLogPathName, smIniPath)
    
    'Work around for auto generated file we don't need
    tmCsiFtpInfo.sLogPathName = Trim$(tmCsiFtpInfo.sLogPathName) & "\" & "Messages\Relic.txt"
    ilRet = csiFTPInit(tmCsiFtpInfo)

    gInitFTP = True
    
    Exit Function
End Function

Public Function gFTPSendFiles(mFileName As String) As Boolean
    
    Dim ilRet As Integer
    Dim ilExportSource As Integer
    
    ilRet = csiFTPFileToServer(Trim$(mFileName))
 
End Function

Public Function gFTPRecvFiles() As Boolean
    
    Dim ilRet As Integer
    Dim ilExportSource As Integer
    
    ' Receive the following files from the server.
    ilRet = csiFTPFileFromServer("FTPTestFile_1.txt")
    ilRet = csiFTPGetStatus(tmCsiFtpStatus)
    While tmCsiFtpStatus.iState = 1
        If ilExportSource = 2 Then DoEvents
        Sleep (200)
        ilRet = csiFTPGetStatus(tmCsiFtpStatus)
    Wend
    If tmCsiFtpStatus.iStatus <> 0 Then
        ' Errors occured.
        ilRet = csiFTPGetError(tmCsiFtpErrorInfo)
        If Not imBackGround Then
            MsgBox "FTP Failed. " & tmCsiFtpErrorInfo.sInfo
            MsgBox "The file name was " & tmCsiFtpErrorInfo.sFileThatFailed
        End If
        gLogMsg "FTP Failed. " & tmCsiFtpErrorInfo.sInfo, "FTPLog.txt.Txt", False
        gLogMsg "   The file name was " & tmCsiFtpErrorInfo.sFileThatFailed, "FTPLog.txt.Txt", False
        Exit Function
    End If
    If Not imBackGround Then
        MsgBox "Recv Complete"
    End If
    gLogMsg "Recv Complete", "FTPLog.txt.Txt", False
End Function

Public Function gTestFTPFileExists(sFileName As String) As Integer

    Dim ilRet As Integer
    Dim slTemp As String
    Dim llCount As Long
    Dim slMsg As String

    gTestFTPFileExists = 0
    tgCsiFtpFileListing.nTotalFiles = 0
    slTemp = Trim$(tgCsiFtpFileListing.sPathFileMask) & "\" & sFileName
    tgCsiFtpFileListing.sPathFileMask = slTemp
    tgCsiFtpFileListing.sSavePathFileName = tgCsiFtpFileListing.sLogPathName
    tgCsiFtpFileListing.sLogPathName = Trim$(tgCsiFtpFileListing.sLogPathName) & "\" & "Messages\Relic.txt"
    tgCsiFtpFileListing.sSavePathFileName = tgCsiFtpFileListing.sLogPathName
    ilRet = csiFTPGetFileListing(tgCsiFtpFileListing)
    llCount = tgCsiFtpFileListing.nTotalFiles
    If ilRet = 0 Then
        If Not imBackGround Then
            MsgBox "csiFTPGetFileListing failed. Please notify Counterpoint."
        End If
        gLogMsg "csiFTPGetFileListing failed. Please notify Counterpoint.", "FTPLog.txt.Txt", False
        gTestFTPFileExists = -1
        Exit Function
    End If
            
    'Debug
    'MsgBox tgCsiFtpFileListing.nTotalFiles & " Files were found and written to " & tgCsiFtpFileListing.sSavePathFileName
    gTestFTPFileExists = llCount
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    slMsg = ""
    If (Err.Number <> 0) And (slMsg = "") Then
        slMsg = "A general error has occured in modWebSubs - gTestFTPFileExists: "
        If Not imBackGround Then
            gMsgBox slMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        gLogMsg slMsg & Err.Description & " Error #" & Err.Number, "FTPLog.txt.Txt", False
        gLogMsg slMsg & Err.Description & " Error #" & Err.Number, "FTPLog.txt", False
    End If
    Exit Function

End Function

Public Function mCheckFTPStatus(mFileName As String) As Boolean

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slTemp As String
    
    On Error GoTo ErrHand
    mCheckFTPStatus = False
    imFtpInProgress = True
    
    ilRet = csiFTPGetStatus(tmCsiFtpStatus)
    '1 = Busy, 0 = Not Busy
    If tmCsiFtpStatus.iState = 1 Then
        Exit Function
    Else
        If tmCsiFtpStatus.iStatus <> 0 Then
            ' Errors occured.
            ilRet = csiFTPGetError(tmCsiFtpErrorInfo)
            slTemp = Trim(tmCsiFtpErrorInfo.sInfo)
            If Not imBackGround Then
                MsgBox "Please Check FTPLog.txt" & vbCrLf & vbCrLf & "FTP Failed. " & slTemp
            End If
            gLogMsg "Error: " & "FAILED to FTP " & tmCsiFtpErrorInfo.sFileThatFailed, "FTPLog.txt", False
            imError = True
            Exit Function
        Else
            ilRet = gTestFTPFileExists(Trim$(mFileName))
            If ilRet = 1 Then
                gLogMsg "  Success, FTP - " & mFileName, "FTPLog.txt", False
                imFtpInProgress = False
                mCheckFTPStatus = True
            End If
        End If
    End If
    Exit Function
    Chr (10)
ErrHand:
    Screen.MousePointer = vbDefault
    'debug
    'Resume Next
    Exit Function
    
End Function


Public Function mLoadOption(Section As String, Key As String, sValue As String, sFileName As String) As Boolean
    
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    Dim slFileName As String

    mLoadOption = False
    BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, sFileName)
    If BytesCopied > 0 Then
        If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
            sValue = Left(sBuffer, BytesCopied)
            mLoadOption = True
        End If
    End If
    Exit Function

ERR_mLoadOption:
    ' return now if an error occurs
End Function

Private Function mReadIniFile(sTargetStr As String) As Boolean

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim slTemp As String
    Dim ilPos As Integer
    
    mReadIniFile = False
    smCurDir = CurDir$
    smIniPath = sgCurDir & "\Traffic.Ini"
    
    If fs.FileExists(smIniPath) Then
        Set tlTxtStream = fs.OpenTextFile(smIniPath, ForReading, False)
    End If
    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        If InStr(Trim$(slRetString), sTargetStr) Then
            ilPos = InStr(Trim$(slRetString), sTargetStr)
            If ilPos > 1 Then
                tlTxtStream.Close
                Exit Function
            End If
            tlTxtStream.Close
            mReadIniFile = True
            Exit Function
        End If
    Loop
    tlTxtStream.Close
End Function

Private Sub mDeleteUselessFile(sLocation As String)

    'In CSI_Utils it writes a file of it's FTP listing.  Is has no value here so delete it to avoid confusion

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject

    If fs.FileExists(sLocation) Then
        fs.DeleteFile sLocation
    End If
End Sub


