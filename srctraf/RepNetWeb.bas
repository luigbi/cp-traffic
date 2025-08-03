Attribute VB_Name = "RepNetWeb"
Option Explicit
Option Compare Text

Public gMsg As String
Public gWebAccessTestedOk As Boolean
Public sgServerDateTime As String
Public Const INTERNET_SERVICE_FTP = 1
Public gURLHasAccess() As Integer

Public Declare Function InternetOpen _
   Lib "wininet.dll" Alias "InternetOpenA" ( _
   ByVal sAgent As String, _
   ByVal nAccessType As Long, _
   ByVal sProxyName As String, _
   ByVal sProxyBypass As String, _
   ByVal nFlags As Long) As Long
   
Public Declare Function InternetConnect _
   Lib "wininet.dll" Alias "InternetConnectA" ( _
   ByVal hInternetSession As Long, _
   ByVal sServerName As String, _
   ByVal nServerPort As Integer, _
   ByVal sUserName As String, _
   ByVal sPassword As String, _
   ByVal nService As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Public Declare Function FtpGetFile _
   Lib "wininet.dll" Alias "FtpGetFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszRemoteFile As String, _
   ByVal lpszNewFile As String, _
   ByVal fFailIfExists As Boolean, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile _
   Lib "wininet.dll" Alias "FtpPutFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszLocalFile As String, _
   ByVal lpszRemoteFile As String, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Boolean

Public Declare Function FtpDeleteFile _
   Lib "wininet.dll" Alias "FtpDeleteFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszFileName As String) As Boolean

Public Declare Function FtpRenameFile _
   Lib "wininet.dll" Alias "FtpRenameFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszExisting As String, _
   ByVal lpszNewName As String) As Boolean

Public Declare Function FtpFindFirstFile _
   Lib "wininet.dll" Alias "FtpFindFirstFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszSearchFile As String, _
   ByRef lpFindFileData As WIN32_FIND_DATA, _
   ByVal dwFlags As Long, _
   ByVal dwContent As Long) As Long
   
Private Declare Function InternetFindNextFile _
   Lib "wininet.dll" Alias "InternetFindNextFileA" ( _
   ByVal hFind As Long, _
   ByRef lpvFindData As WIN32_FIND_DATA) As Long
   
Public Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * 260
   cAlternate As String * 14
End Type




' This function attempts to take control of the WebSession table by inserting an entry in it.
' If it can insert an entry using UKey = 1 then is gains control.
' If there is already an entry in the table it checks to make sure the PCName is not ours. Otherwise
' it means we have it locked already. If another PC has an entry in it, then the date and time
' is checked. If the date and time is greater than the amount specified, this function will take
' control of the web session anyway.
' This function returns one of two things.
' Either 0 meaning we now have control or a number indicating the total number of minutes to
' wait for the other PC.

Public Function gStartWebSession(PType As String, PSubType As String, LogFileName As String) As Integer
    
    On Error GoTo ErrHandler
    
    Dim slDateTime As String
    Dim slServersDateTime As String
    Dim slPCName As String
    Dim slExportLockoutMinutes As String
    Dim slComputerName As String
    Dim alDataArray() As String
    Dim slRootURL As String
    Dim SQLQuery As String
    Dim ilTotalTime As Integer
    Dim llTotalRecords As Long
    Dim llRet As Long
    Dim ilRet As Integer
    Dim slRegSection As String
    Const lcMaxTime = 30
    
    gStartWebSession = 0
    
    slRegSection = Trim$(tgServerNRF(igNrfIndex).sIISRegSection)
    
    llRet = gExecWebSQL(alDataArray, "Select GetDate() as ServerDateTime", True)
    slDateTime = gGetDataNoQuotes(alDataArray(0))
    slDateTime = Format(slDateTime, "YYYY-MM-DD HH:MM:SS")
    
    SQLQuery = "Insert into LPT_Lock_Process (UKey, lptRepDBID, lptNetDBID, lptProcessType, lptProcessSubType, lptDateTimeEntered) "
    SQLQuery = SQLQuery & "Values(1, '" & sgRepDBID & "', '" & sgNetDBID & "', '" & PType & "', '" & PSubType & "', "
    SQLQuery = SQLQuery & "'" & slDateTime & "')"
    llRet = gExecWebSQLWithRowsEffected(SQLQuery)
    
    If llRet = -1 Then
        'The insert failed so somebody has it locked???
        SQLQuery = "Select * From LPT_Lock_Process Where Ukey = 1"
        llRet = gExecWebSQL(alDataArray, SQLQuery, True)
        If CInt(Trim(alDataArray(1))) = CInt(sgRepDBID) Then
            'the correct user already has the lock
            gStartWebSession = 0
            Exit Function
        Else
            'somebody else has the lock, see how long they have had it
            ilTotalTime = DateDiff("n", alDataArray(5), slDateTime)
            If ilTotalTime >= lcMaxTime Then
                'it's been locked too long, end the lock and start a new lock for the new user
                ilRet = gEndWebSession(LogFileName)
                If ilRet Then
                    llRet = gExecWebSQL(alDataArray, SQLQuery, True)
                    If llRet = 0 Then
                        'Insert was successful, we have control
                        gStartWebSession = 0
                        Exit Function
                    Else
                        'Insert was NOT successful, we don't have control
                        llRet = llRet
                    End If
                Else
                    'Error, the web lock failed to clear
                    ilRet = ilRet
                End If
                
            Else
                'we need to let the lock finish, get out and try again later
                gStartWebSession = 1
                Exit Function
            End If
        End If
    Else
        'Insert was successful, we have control
        gStartWebSession = 0
        Exit Function
    End If

    Exit Function

ErrHandler:
    gMsg = "A general error has occured in modRepNetWeb-gStartWebSession: "
    gLogMsg gMsg & err.Description & " Error #" & err.Number, LogFileName, False
    gMsgBox gMsg & err.Description & "; Error #" & err.Number, vbCritical + vbOKOnly, "General Traffic Error"
End Function


Public Function gEndWebSession(LogFileName As String) As Boolean
    
    Dim alDataArray() As String
    Dim llRet As Long
    
    On Error GoTo ErrHandler
    
    gEndWebSession = False
    llRet = gExecWebSQL(alDataArray, "Delete From LPT_Lock_Process Where UKey = 1", False)
    If llRet <> -1 Then
        gLogMsg "User: " & sgUserName & " " & "Web Lock Ended Successfully", LogFileName, False
    Else
        gLogMsg "User: " & sgUserName & " " & "Web Lock Failed to End", LogFileName, False
    End If
    gEndWebSession = True
    Exit Function

ErrHandler:
    gMsg = "A general error has occured in modRepNetWeb-gEndWebSession: "
    gLogMsg gMsg & err.Description & " Error #" & err.Number, LogFileName, False
    gMsgBox gMsg & err.Description & "; Error #" & err.Number, vbCritical + vbOKOnly, "General Traffic Error"
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

Public Function gExecWebSQL(aDataArray() As String, sSQL As String, iWantData) As Long
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim alRecordsArray() As String
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim WebCmds As New WebCommands

    gExecWebSQL = -1    ' -1 is an error condition.

    slRootURL = Trim$(tgServerNRF(igNrfIndex).sIISRootURL)
    slRootURL = gSetPathEndSlash(slRootURL)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    
    slRegSection = Trim$(tgServerNRF(igNrfIndex).sIISRegSection)
    If slRegSection <> "" Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To 5
        llErrorCode = 0
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(sSQL)
            If Not Left(slResponse, 5) = "ERROR" Then
                llReturn = 200
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            objXMLHTTP.Send
            llReturn = objXMLHTTP.Status
            slResponse = objXMLHTTP.responseText
            Set objXMLHTTP = Nothing
        End If
    
        If llReturn = 200 Then
            If Not iWantData Then
                ' Caller does not want any data returned.
                gExecWebSQL = 0
                Exit Function
            End If
    
            ' Parse out the response we got.
            '
            slResponse = Replace(slResponse, """", "")
            alRecordsArray = Split(slResponse, vbCrLf)
            If Not IsArray(alRecordsArray) Then
                Exit Function
            End If
            ' We have to have back at least two records. The first one is the column headers.
            ' The rest of the entries are the data itself.
            If UBound(alRecordsArray) < 2 Then
                ' If the table is empty, we will get back at least one record containing the column
                ' definitions of the table itself, but no data records.
                gExecWebSQL = 0
                Exit Function
            End If
        
            ' Each record we get back is a comma delimited string. In this case were only interested
            ' in the first record.
            aDataArray = Split(alRecordsArray(1), ",")
            If Not IsArray(aDataArray) Then
                Exit Function
            End If
            gExecWebSQL = UBound(aDataArray)
            Exit Function
        Else
            gLogMsg "gExecWebSQL is retrying" & "  User: " & sgUserName, "RepNetLink.txt", False
            Call gSleep(2)
        End If
    Next ilRetries
    Exit Function
    
ErrHandler:
    llErrorCode = err.Number
    gMsg = "A general error has occured in modRepNetWeb-gExecWebSQL: "
    gLogMsg gMsg & err.Description & " Error #" & err.Number, "RepNetLog.txt", False
    gMsgBox gMsg & err.Description & "; Error #" & err.Number, vbCritical + vbOKOnly, "General Traffic Error"
    Resume Next
End Function

'Public Sub gSleep(ilTotalSeconds As Long)
'    Dim ilLoop As Integer
'
'    For ilLoop = 0 To ilTotalSeconds
'        DoEvents
'        Sleep (1000)   ' Wait 1 second
'    Next
'End Sub

'***************************************************************************************
'*
'* Procedure Name: gSendCmdToWebServer
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function instructs the web server to execute a command.
'*           For example, by retrieving the Import.asp page from the server.
'*           The Import.Asp page, once loaded, will start the DTS package on SQL
'*           Server. When complete, it returns a status code in the page itself.
'*
'***************************************************************************************
Public Function gSendCmdToWebServer(sWebPageToAccess As String, sFileName As String) As Boolean
    On Error GoTo ERR_gSendCmdToWebServer
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slWebLogLevel As String
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim slDateTime As String
    Dim alDataArray() As String
    Dim ilCount As Integer
    Dim llRet As Long
    Dim WebCmds As New WebCommands

    gSendCmdToWebServer = False
    
    If slWebLogLevel = "" Then
        slWebLogLevel = "0"
    End If

    slRootURL = tgServerNRF(igNrfIndex).sIISRootURL
    slRootURL = Trim$(slRootURL)
    slRootURL = slRootURL & "/" & "IISIOMngr.dll?"
    slRegSection = Trim$(tgServerNRF(igNrfIndex).sIISRegSection)
    
    llRet = gExecWebSQL(alDataArray, "Select GetDate() as ServerDateTime", True)
    slDateTime = gGetDataNoQuotes(alDataArray(0))
    slDateTime = Format(slDateTime, "YYYY-MM-DD HH:MM:SS")
    
    slISAPIExtensionDLL = slRootURL & sWebPageToAccess & "&RK=" & Trim(slRegSection) & "&FN=" & sFileName & "&DT=" & slDateTime
           
    ilCount = 0
    For ilRetries = 0 To 14
        llErrorCode = 0
        llReturn = 1
        If bgUsingSockets Then
            If sWebPageToAccess = "ImportSiteOptions.dll" Then
                slResponse = WebCmds.ImportSiteOptions(sFileName)
                If Not Left(slResponse, 5) = "ERROR" Then
                    llReturn = 200
                End If
            End If
            
            If sWebPageToAccess = "ExportWebL.dll" Then
                slResponse = WebCmds.ExportWebL(sFileName)
                If Not Left(slResponse, 5) = "ERROR" Then
                    llReturn = 200
                End If
            End If
            
            If sWebPageToAccess = "ExportHeaders.dll" Then
                slResponse = WebCmds.ExportHeaders(sFileName)
                If Not Left(slResponse, 5) = "ERROR" Then
                    llReturn = 200
                End If
            End If
            
            If sWebPageToAccess = "ExportCommit.dll" Then
                slResponse = WebCmds.ExportCommit()
                If Not Left(slResponse, 5) = "ERROR" Then
                    llReturn = 200
                End If
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            objXMLHTTP.Send
            llReturn = objXMLHTTP.Status
            slResponse = objXMLHTTP.responseText
            gLogMsg "slResponse = " & slResponse, "WebServerCmnds.html", False
            Set objXMLHTTP = Nothing
        End If
        If llReturn = 200 Then
            gSendCmdToWebServer = True
            Exit Function
        Else
            ilCount = ilCount + 1
            'gLogMsg "gSendCmdToWebServer is retrying" & "  User: " & sgUserName, "RepNetLink.txt", False
            Call gSleep(2)
        End If
    Next ilRetries

    gLogMsg "gSendCmdToWebServer Failed - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]" & "  User: " & sgUserName & " Retried" & ilCount & " times", "RepNetLink.txt", False
    'gMsgBox "gSendCmdToWebServer Failed - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", vbCritical + vbOkOnly, "Traffic.Ini Error"
    Exit Function

ERR_gSendCmdToWebServer:
    llErrorCode = err.Number
    gMsg = "A general error has occured in modRepNetWeb-gSendCmdToWebServer: "
    gLogMsg gMsg & err.Description & " Error #" & err.Number, "RepNetLog.txt", False
    Resume Next
End Function


Public Function gTestAccessToWebServer(NRFNetIndex As Integer) As String
    On Error GoTo ERR_gTestAccessToWebServer
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slWebPage As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slWebLogLevel As String
    Dim WebCmds As New WebCommands

    gTestAccessToWebServer = ""
    gWebAccessTestedOk = False
    
    slRootURL = tgServerNRF(NRFNetIndex).sIISRootURL
    slRootURL = Trim$(slRootURL)
    slRootURL = gSetPathEndSlash(slRootURL)  ' Make sure the path has the final slash on it.
    
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
'    If Not gLoadOption("WebServer", "RegSection", slRegSection) Then
'        Exit Function
'    End If
    ' Make a request to view the headers. If this operation succeeds, then all is ok.
    ' NOTE: The VH.ASP page will not return any data from the header table because the password here
    ' is not correct due to the date and time being added. But this is ok because it will return the
    ' text that the password is invalid and access to the web page itself returns success.
    
    
    slWebPage = slRootURL & "Main.htm?" & Now()
    If bgUsingSockets Then
        slResponse = WebCmds.ExecSQL("select top 1 * from dbversion")
        If InStr(slResponse, "6.5") Then
            llReturn = 200
        End If
    Else
        Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
        objXMLHTTP.Open "GET", slWebPage, False
        objXMLHTTP.Send
        llReturn = objXMLHTTP.Status
        slResponse = objXMLHTTP.responseText
        Set objXMLHTTP = Nothing
    End If
    
    ' Very the return code here. Anything but 200 is an error.
    If llReturn <> 200 Then
        Exit Function
    End If
    
    gTestAccessToWebServer = slRootURL
    gWebAccessTestedOk = True
    Exit Function

ERR_gTestAccessToWebServer:
    ' Exit the function if any errors occur
End Function

'*****************************************************************************************************
' Returns the number of rows deleted.
'
'*****************************************************************************************************
Public Function gExecWebSQLWithRowsEffected(sSQL As String) As Long
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim ilIdx As Integer
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim WebCmds As New WebCommands
    
    gExecWebSQLWithRowsEffected = -1    ' -1 is an error condition.

    slRootURL = Trim$(tgServerNRF(igNrfIndex).sIISRootURL)
    slRootURL = gSetPathEndSlash(slRootURL)  ' Make sure the path has the final slash on it.
    
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    
    slRegSection = Trim$(tgServerNRF(igNrfIndex).sIISRegSection)
    If slRegSection <> "" Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    For ilRetries = 0 To 14
        llErrorCode = 0
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(sSQL)
            If Not Left(slResponse, 5) = "ERROR" Then
                llReturn = 200
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            objXMLHTTP.Send
            llReturn = objXMLHTTP.Status
            slResponse = Trim(objXMLHTTP.responseText)
            Set objXMLHTTP = Nothing
        End If
        
        If Left(slResponse, 5) = "ERROR" Then
            Exit Function
        End If
        
        If llReturn = 200 Then
            ilIdx = 1
            llReturn = 0
            While ilIdx < Len(slResponse) - 1 And Mid(slResponse, ilIdx, 1) <> " "
                ilIdx = ilIdx + 1
            Wend
            llReturn = Left(slResponse, ilIdx)
            gExecWebSQLWithRowsEffected = llReturn
            Exit Function
        Else
            gLogMsg "gExecWebSQLWithRowsEffected is retrying" & "  User: " & sgUserName, "RepNetLink.txt", False
            Call gSleep(2)  ' Delay for 2 seconds between requests.
        End If
    Next
    Exit Function
    
ErrHandler:
    llErrorCode = err.Number
    gMsg = "A general error has occured in RepNetWeb-gExecWebSQLWithRowsEffected: Retries = " & ilRetries
    gLogMsg gMsg & err.Description & " Error #" & err.Number, "RepNetLink.txt", False
    gLogMsg "    SQLQuery = " & sSQL, "RepNetLink.txt", False
    Resume Next
End Function

Public Function gGetDataNoQuotes(sDataStr As String) As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slNewStr As String
    Dim clOneChar As String
    
    ilLen = Len(sDataStr)
    slNewStr = ""
    For ilLoop = 1 To ilLen
        clOneChar = Mid(sDataStr, ilLoop, 1)
        If clOneChar <> """" Then
            slNewStr = slNewStr + clOneChar
        End If
    Next
    gGetDataNoQuotes = Trim(slNewStr)
End Function

'***************************************************************************************
'*
'* Procedure Name: gFTPFileToWebServer
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function obtains all necessary information from the ini file to
'*           FTP the specified file to the Web Server.
'*
'***************************************************************************************
Public Function gFTPFileToWebServer(sPathFileName As String, sFileName As String) As Boolean
    On Error GoTo ERR_gFTPFileToWebServer
    Dim hINetSession As Long
    Dim hSession As Long
    Dim FTPIsOn As String
    Dim FTPAddress As String
    Dim FTPPort As String
    Dim FTPUID As String
    Dim FTPPWD As String
    Dim FTPWebDir As String
    Dim FTPImportDir As String
    Dim ServerFileName As String
    Dim slRegSection As String

    gFTPFileToWebServer = False
    FTPAddress = Trim(tgServerNRF(igNrfIndex).sFTPAddress)
    FTPPort = Trim(tgServerNRF(igNrfIndex).iFTPPort)
    FTPUID = Trim(tgServerNRF(igNrfIndex).sFTPUserID)
    FTPPWD = Trim(tgServerNRF(igNrfIndex).sFTPUserPW)

    'Debug Only - User Id and Password have to be blank for Jeff's FTP server
    If FTPUID = "BLANK" Then
        FTPUID = ""
    End If
    If FTPPWD = "BLANK" Then
        FTPPWD = ""
    End If
    
    FTPImportDir = Trim(tgServerNRF(igNrfIndex).sFTPImportDir)
    slRegSection = Trim(tgServerNRF(igNrfIndex).sIISRegSection)
    ' Connect to the internet
    hINetSession = InternetOpen(slRegSection, 0, vbNullString, vbNullString, 0)
    If hINetSession < 1 Then
        gMsgBox "FAIL: gFTPFileToWebServer: InternetOpen Failed", vbCritical + vbOKOnly, "Traffic.Ini Error"
        Exit Function
    End If
    hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, 0, 0)
    If hINetSession < 1 Then
        Call InternetCloseHandle(hINetSession)
        gMsgBox "FAIL: gFTPFileToWebServer: InternetConnect Failed", vbCritical + vbOKOnly, "Traffic.Ini Error"
        Exit Function
    End If

    FTPWebDir = Trim(tgServerNRF(igNrfIndex).sFTPImportDir)
    ' Send the data to the server
    FTPWebDir = gSetPathEndSlash(FTPWebDir)
    ServerFileName = FTPWebDir & sFileName
 
    If FtpPutFile(hSession, sPathFileName, ServerFileName, 1, 0) = False Then
        gLogMsg "FAIL: gFTPFileToWebServer: FtpPutFile Failed. " & sPathFileName & ", " & ServerFileName & "  User: " & sgUserName, "RepNetLink.txt", False
        gMsgBox "FAIL: gFTPFileToWebServer: FtpPutFile Failed. " & sPathFileName & ", " & ServerFileName, vbCritical + vbOKOnly, "Traffic.Ini Error"
        Call InternetCloseHandle(hSession)
        Call InternetCloseHandle(hINetSession)
        Exit Function
    End If
    
    Call InternetCloseHandle(hSession)
    Call InternetCloseHandle(hINetSession)

    gFTPFileToWebServer = True
    Exit Function

ERR_gFTPFileToWebServer:
    ' Exit the function if any errors occur
End Function


Public Function gFTPFileFromWebServer(sPathFileName As String, sFileName As String) As Boolean

    Dim hINetSession As Long
    Dim hSession As Long
    Dim FTPIsOn As String
    Dim FTPAddress As String
    Dim FTPPort As String
    Dim FTPUID As String
    Dim FTPPWD As String
    Dim FTPWebDir As String
    Dim ServerFileName As String
    Dim slLocalPath As String
    Dim slRegSection As String

    On Error GoTo ERR_gFTPFileFromWebServer
    
    gFTPFileFromWebServer = False
    FTPAddress = Trim(tgServerNRF(igNrfIndex).sFTPAddress)
    FTPPort = Trim(tgServerNRF(igNrfIndex).iFTPPort)
    FTPUID = Trim(tgServerNRF(igNrfIndex).sFTPUserID)
    FTPPWD = Trim(tgServerNRF(igNrfIndex).sFTPUserPW)

    'Debug Only - User Id and Password have to be blank for Jeff's FTP server
    If FTPUID = "BLANK" Then
        FTPUID = ""
    End If
    If FTPPWD = "BLANK" Then
        FTPPWD = ""
    End If
    
    FTPWebDir = Trim(tgServerNRF(igNrfIndex).sFTPExportDir)
    FTPWebDir = gSetPathEndSlash(FTPWebDir)
    slRegSection = Trim(tgServerNRF(igNrfIndex).sIISRegSection)
    
    ' Connect to the internet
    hINetSession = InternetOpen(slRegSection, 0, vbNullString, vbNullString, 0)
    If hINetSession < 1 Then
        gLogMsg "FAIL: FAIL: gFTPFileFromWebServer: InternetOpen Failed" & sgUserName, "RepNetLink.txt", False
        Exit Function
    End If
    hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, 0, 0)
    If hINetSession < 1 Then
        Call InternetCloseHandle(hINetSession)
        gLogMsg "FAIL: gFTPFileFromWebServer: InternetConnect Failed", "RepNetLink.txt", False
        Exit Function
    End If

    ' Receive the data from the server
    
    slLocalPath = sPathFileName & sFileName
    ServerFileName = FTPWebDir & sFileName
    If FtpGetFile(hSession, ServerFileName, slLocalPath, False, 0, 1, 0) = False Then
        gLogMsg "FAIL: RecvImportFileFromWebServer: FtpPutFile Failed. " & slLocalPath & ", " & ServerFileName, "RepNetLink.txt", False
        Call InternetCloseHandle(hSession)
        Call InternetCloseHandle(hINetSession)
        Exit Function
    End If
    
    Call InternetCloseHandle(hSession)
    Call InternetCloseHandle(hINetSession)

    gFTPFileFromWebServer = True
    Exit Function

ERR_gFTPFileFromWebServer:
    ' Exit the function if any errors occur
End Function

