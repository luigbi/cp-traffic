Attribute VB_Name = "modWebSubs"
Option Compare Text
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private tmCsiFtpInfo As CSIFTPINFO
Private tmCsiFtpStatus As CSIFTPSTATUS



Public Function gBuildWebHeaderDetail() As String

    '   D.S. 08/06/07
    'Build the detail/first line portion of the web header
    '**** Important Note: If you change the below fields you must change all other occurances of it.   ****
    '**** Currently this includes gBuildWebHeaders, gBuildwebHeaderDetail and mSetVefDate in 3 places. ****
    'D.S. 11/4/12 new header below
    'FYM  2/6/19 added two fields: Market, Rank
    gBuildWebHeaderDetail = "attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, LogType, PostType, StartTime, StationEmail, StationPW, AggreementEmail, AggreementPW, SendLogEmail, VehicleFTPSite, TimeZone, ShowAvailNames, Multicast, WebLogSummary, WebLogFeedTime, Mode, LogStartDate, LogEndDate, MonthlyPosting, InterfaceType, UseActual, SuppressLog, PledgeByEvent, altVehName, MGsOnWeb, ReplacementsOnWeb, WebSiteVersion, Market, Rank, ShowCart"
End Function
Public Function VerifyVersionTable() As Boolean
    
    Dim smPath As String
    Dim smToFile As String
    Dim smPathFileName As String
    Dim hmFileHandle As Integer
    Dim ilRet As Integer
    Dim slVersion As String
    Dim slVersionDate As String
    Dim slNotes As String
    Dim ilRetry As Integer
    Dim ilSuccess As Boolean
    
    VerifyVersionTable = False
    On Error GoTo ErrHandler:
    
    smToFile = "TableTest" & "_" & sgUserName & ".txt"
    Call gLoadOption(sgWebServerSection, "WebImports", smPath)
    smPath = gSetPathEndSlash(smPath, True)
    smPathFileName = smPath & smToFile
    ilSuccess = False
    For ilRetry = 0 To 4 Step 1
        If Not gRemoteExecSql("Select Top 1 * from DBVersion", smToFile, "WebImports", True, True, 10) Then
            Sleep (1000)
        Else
            ilSuccess = True
            Exit For
        End If
    Next ilRetry
        
    If Not ilSuccess Then
        gLogMsg "Error: modWebSubs - VerifyVersionTable: Select Top 1 * from DBVersion", "AffWebErrorLog.Txt", False
        Exit Function
    End If
    
    On Error GoTo CatchFileError:
    
    'hmFileHandle = FreeFile
    ilRet = 0
    'Open smPathFileName For Input Access Read Lock Write As hmFileHandle
    ilRet = gFileOpen(smPathFileName, "Input Access Read Lock Write", hmFileHandle)
    If ilRet <> 0 Then
        gLogMsg "Error: modWebSubs - VerifyVersionTable: Unable to open: " & smPathFileName, "AffWebErrorLog.Txt", False
        Close #hmFileHandle
        Exit Function
    End If
    
    On Error GoTo ErrHandler:
    ' Read in the column header and verify each column exist.
    
    Input #hmFileHandle, slVersion, slVersionDate, slNotes
 
    If slVersion <> "Version" Then
        gLogMsg "Error: modWebSubs - VerifyVersionTable: Version <> Version " & slVersion, "AffWebErrorLog.Txt", False
        Close #hmFileHandle
        Exit Function
    End If
    If slVersionDate <> "VersionDate" Then
        gLogMsg "Error: modWebSubs - VerifyVersionTable: VersionDate <> slVersionDate " & slVersionDate, "AffWebErrorLog.Txt", False
        Close #hmFileHandle
        Exit Function
    End If
    If slNotes <> "Notes" Then
        gLogMsg "Error: modWebSubs - VerifyVersionTable: Notes <> Notes " & slNotes, "AffWebErrorLog.Txt", False
        Close #hmFileHandle
        Exit Function
    End If
    
    Input #hmFileHandle, slVersion, slVersionDate, slNotes
    sgWebSiteVersion = slVersion
    sgWebSiteDate = slVersionDate
    Close #hmFileHandle
    
    If ilRet = 0 Then
        VerifyVersionTable = True
    Else
        VerifyVersionTable = False
    End If
    
    Exit Function
    
CatchFileError:
    ilRet = 1
    Resume Next
    
ErrHandler:
    Exit Function
 
End Function


Public Function CreateVersionTable() As Boolean
    
    ' Jeff D. 8/22/08
    
    Dim slSQLQuery As String
    Dim slArray(0 To 1) As String
    Dim ilResult As Long
    
    CreateVersionTable = False
    slSQLQuery = "Create Table DBVersion"
    slSQLQuery = slSQLQuery & "("
    slSQLQuery = slSQLQuery & "   Version        [char] (10) NOT NULL ,"
    slSQLQuery = slSQLQuery & "   VersionDate    [datetime] default '1899-12-31' NOT NULL,"
    slSQLQuery = slSQLQuery & "   Notes          [char] (255) NULL"
    slSQLQuery = slSQLQuery & ") ON [PRIMARY]"
    
    ilResult = gExecWebSQL(slArray, slSQLQuery, False)
    If ilResult <> 0 Then
        gLogMsg "ERROR: modWebSubs-CreateVersionTable:  " & slSQLQuery, "AffWebErrorLog.Txt", False
        Exit Function
    End If
    CreateVersionTable = True
    Exit Function
 
End Function
 
Function InsertVersionID(sVersion As String, sNotes As String) As Boolean

    ' Jeff D. 8/22/08

    Dim slSQLQuery As String
    Dim slArray(0 To 1) As String
    Dim ilResult As Long
    Dim sDateTime As String
 
    InsertVersionID = False
    sDateTime = Now()
    slSQLQuery = "Insert Into DBVersion (Version, VersionDate, Notes) Values ("
    slSQLQuery = slSQLQuery & "'" & sVersion & "', "
    slSQLQuery = slSQLQuery & "'" & sDateTime & "', "
    slSQLQuery = slSQLQuery & "'" & sNotes & "')"
 
    ilResult = gExecWebSQL(slArray, slSQLQuery, False)
    If ilResult <> 0 Then
        gLogMsg "ERROR: modWebSubs-InsertVersionID:  " & slSQLQuery, "AffWebErrorLog.Txt", False
        Exit Function
    End If
    InsertVersionID = True
    Exit Function
 
End Function

Public Function gBuildWebHeaders(cprst As ADODB.Recordset, iVefCode As Integer, sVefName As String, iShttCode As Integer, sAttWebInterface As String, iSendEmails As Integer, sMode As String, sStartDate As String, sEndDate As String, sUseActual As String, sSuppressLog As String) As String
 
    'D.S. 1/5/05
    'Used to build header records for the web site.  This function is called by frmStations,
    'frmWebExportSchdSpot and frmAgmnt.  This is where you need to be if your going to change
    'the format of the header agreements.

    Dim rst_Temp As ADODB.Recordset
    Dim slFTP As String
    Dim slTemp As String
    Dim slStr As String
    Dim llTemp As Long
    Dim rst_Gsf As ADODB.Recordset
    Dim rst_emt As ADODB.Recordset
    Dim llRow As Long
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slGsfCode As Long
    Dim slGameDate As String
    Dim slGameStartTime As String
    Dim slVisitTeamName As String
    Dim slVisitTeamAbbr As String
    Dim slHomeTeamName As String
    Dim slHomeTeamAbbr As String
    Dim slLanguageCode As String
    Dim slFeedSource As String
    Dim slWebLogSummary As String
    Dim slWebLogFeedTime As String
    Dim slMultiCastFlag As String
    Dim sWebEMail As String
    Dim ilLen As Integer
    Dim sSDate As String
    Dim sEDate As String
    Dim slPledgeByEvent As String
    Dim ilVff As Integer
    Dim llWebAttID As Long
    Dim slShowVehName As String
    Dim slClientAbrvName As String
    Dim slShowCart As String        'FYM 2/28/19    Added ShowCart from vff
    On Error GoTo ErrHand
    
    gBuildWebHeaders = ""
    
    sSDate = Format$(sStartDate, sgSQLDateForm)
    sEDate = Format$(sEndDate, sgSQLDateForm)
    
'    SQLQuery = "SELECT emtEmail"
'    SQLQuery = SQLQuery + " FROM emt"
'    SQLQuery = SQLQuery + " WHERE (emtShttCode = " & iShttCode & ")"
'    Set rst_Emt = gSQLSelectCall(SQLQuery)
'
    sWebEMail = ""
'    While Not rst_Emt.EOF
'        DoEvents
'        sWebEMail = sWebEMail & Trim$(rst_Emt!emtEmail)
'        sWebEMail = sWebEMail & ","
'        rst_Emt.MoveNext
'    Wend
'    ilLen = Len(sWebEMail)
'
'    If ilLen > 0 Then
'        sWebEMail = Left(sWebEMail, ilLen - 1)
'    End If
    
    slMultiCastFlag = Trim$(cprst!attMulticast)
    llWebAttID = gGetLogAttID(iVefCode, iShttCode, cprst!attCode)
    'D.S. 11/4/12
    ilVff = gBinarySearchVff(iVefCode)
    slShowVehName = ""
    If ilVff <> -1 Then
        slShowVehName = Trim$(tgVffInfo(ilVff).sWebName)
    End If
    
    'D.S. TTP #6511
    SQLQuery = "SELECT mnfName"
    SQLQuery = SQLQuery & " FROM SPF_Site_Options, MNF_Multi_Names"
    SQLQuery = SQLQuery & " WHERE spfCode = 1"
    SQLQuery = SQLQuery & " AND spfMnfClientAbbr = mnfCode"
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    'FTP site for audio to show on the log screen
    If Not rst_Temp.EOF Then
        slClientAbrvName = Trim$(rst_Temp!mnfName)
    Else
        slClientAbrvName = sgClientName
    End If
    If slClientAbrvName = "" Then
        slClientAbrvName = sgClientName
    End If
    
    
    ' Build the header information record
    'slStr = Trim$(Str$(cprst!attCode)) & ","
    slStr = Trim$(Str$(llWebAttID)) & ","
    slStr = slStr & """" & "Counterpoint Software" & ""","       'Software Provider Name
    slStr = slStr & """" & sgClientName & ""","                  'Web Software Provider
    slStr = slStr & """" & slClientAbrvName & ""","                  'Station Station Provider
    slStr = slStr & """" & sgClientName & ""","                  'Network Provider Name
        
    'D.S. 11/13/12
    If slShowVehName <> "" Then
        'If the web is to show an alternate vehicle name change out the original vehicle name to the alternate name
        slStr = slStr & """" & slShowVehName & ""","   'Alternate Vehicle name
    Else
        'No alternate vehicle name so use the original vehicle name
        slStr = slStr & """" & Trim$(sVefName) & ""","              'Vehicle name & ""","     'Vehicle name
    End If
    slStr = slStr & """" & Trim$(cprst!shttCallLetters) & ""","
    slStr = slStr & Trim$(Str(cprst!attLogType)) & ","
    slStr = slStr & Trim$(Str(cprst!attPostType)) & ","
    slStr = slStr & Format$(cprst!attStartTime, "hh:mma/p") & ","
    'slStr = slStr & """" & Trim$(cprst!shttWebEmail) & ""","
    slStr = slStr & """" & Trim$(sWebEMail) & ""","
    slStr = slStr & """" & Trim$(cprst!shttWebPW) & ""","
    'slStr = slStr & """" & Trim$(cprst!attWebEmail) & ""","
    slStr = slStr & """" & "" & ""","
    slStr = slStr & """" & Trim$(cprst!attWebPW) & ""","
    If iSendEmails Then
        slStr = slStr & """" & Trim$(cprst!attSendLogEmail) & ""","
    Else
        slStr = slStr & """" & "1" & ""","
    End If

    SQLQuery = "SELECT vpfAvailnameonweb, arfFTP"
    SQLQuery = SQLQuery & " FROM VPF_Vehicle_Options, ARF_Addresses"
    SQLQuery = SQLQuery & " WHERE vpfVefKCode = " & iVefCode
    SQLQuery = SQLQuery & " AND vpfFTPArfCode = arfCode"
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    'FTP site for audio to show on the log screen
    If Not rst_Temp.EOF Then
        slFTP = Trim$(rst_Temp!arfFTP)
        If InStr(1, slFTP, "://", vbTextCompare) = 0 Then
            If InStr(1, slFTP, "www", vbTextCompare) = 1 Then
                slFTP = "http://" & slFTP
            ElseIf InStr(slFTP, "ftp") Then
                slFTP = "ftp://" & slFTP
            End If
        End If
        slStr = slStr & """" & slFTP & ""","
    Else
        slStr = slStr & """" & ""","
    End If
    
    'Time Zones
    slTemp = Trim$(cprst!shttTimeZone)
    Select Case slTemp
        Case "EST"
            slStr = slStr & """" & "Eastern Zone" & ""","
        Case "CST"
            slStr = slStr & """" & "Central Zone" & ""","
        Case "MST"
            slStr = slStr & """" & "Mountain Zone" & ""","
        Case "PST"
            slStr = slStr & """" & "Pacific Zone" & ""","
        Case Else
            slStr = slStr & """" & "" & ""","
            gLogMsg "WARNING: Time zone is missing for station " & Trim$(cprst!shttCallLetters), "NoTimeZoneStations.Txt", False
    End Select
    
    'Show Avail Names?
    SQLQuery = "SELECT vpfAvailnameonweb"
    SQLQuery = SQLQuery & " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery & " WHERE vpfVefKCode = " & iVefCode
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        If rst_Temp!vpfAvailNameOnWeb = "Y" Then
            slStr = slStr & """" & 1 & ""","
            igSendAvails = True
        Else
            slStr = slStr & """" & 0 & ""","
            igSendAvails = False
        End If
    Else
        'default to no
        slStr = slStr & """" & 0 & ""","
    End If
    
    llTemp = gGetStaMulticastGroupID(iShttCode)
    If llTemp = 0 Then
        slStr = slStr & """" & ""","
    Else
        'Even though there may be a group ID it still might not be multicast.
        If slMultiCastFlag = "Y" Then
            slStr = slStr & """" & llTemp & ""","
        Else
            slStr = slStr & """" & ""","
        End If
    End If
    
    ' JD - 06-01-07, Moved this code from after the gaming fields to before.
    SQLQuery = "SELECT vpfWebLogSummary, vpfWebLogFeedTime"
    SQLQuery = SQLQuery & " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery & " WHERE vpfVefKCode = " & iVefCode
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        slWebLogSummary = Trim$(rst_Temp!vpfWebLogSummary)
        slWebLogFeedTime = Trim$(rst_Temp!vpfWebLogFeedTime)
        slStr = slStr & """" & slWebLogSummary & ""","
        slStr = slStr & """" & slWebLogFeedTime & ""","
        
    Else
        slStr = slStr & """" & ""","
        slStr = slStr & """" & ""","
    End If
    
    slStr = slStr & """" & sMode & ""","
    slStr = slStr & sSDate & ","
    slStr = slStr & sEDate & ","
    
    slTemp = Trim$(cprst!attMonthlyWebPost)
    If slTemp = "" Then
        'Default value
        slTemp = "N"
    End If
    slStr = slStr & """" & slTemp & ""","
    
    slStr = slStr & """" & sAttWebInterface & ""","
    
    slStr = slStr & """" & sUseActual & ""","
    
    slStr = slStr & """" & sSuppressLog & ""","
    
    'Pledge By Event
    If iVefCode <= 0 Then
        slPledgeByEvent = "N"
    Else
        ilVff = gBinarySearchVff(iVefCode)
        If ilVff <> -1 Then
            slPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
            If slPledgeByEvent = "" Then
                slPledgeByEvent = "N"
            End If
        End If
    End If
    slStr = slStr & """" & slPledgeByEvent & ""","
    
    'D.S. 11/13/12
    If slShowVehName <> "" Then
        'If the show alternate vehicle name is not blank then replace it with the original vehicle name.
        'Above the original vehicle name was replaced by the alternate vehicle name
        slStr = slStr & """" & Trim$(sVefName) & ""","             'Vehicle name
    Else
        slStr = slStr & """" & " " & ""","
    End If
    
    SQLQuery = "SELECT vffMGsOnWeb, vffReplacementOnWeb, vffCartOnWeb "
    SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
    SQLQuery = SQLQuery + " WHERE (vffVefCode = " & iVefCode & ")"
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        If Trim$(rst_Temp!vffMGsOnWeb) <> "" Then
            'slStr = slStr & """" & Trim$(rst_Temp!vffMGsOnWeb) & ""","
            If Trim$(rst_Temp!vffMGsOnWeb) = "Y" Then
                slStr = slStr & 1 & ","
            Else
                slStr = slStr & 0 & ","
            End If
        Else
            slStr = slStr & 0 & ","
        End If
        
        If Trim$(rst_Temp!vffReplacementOnWeb) <> "" Then
            If Trim$(rst_Temp!vffReplacementOnWeb) = "Y" Then
                slStr = slStr & 1 & ","
            Else
                slStr = slStr & 0 & ","
            End If
        Else
            slStr = slStr & 0 & ","
        End If
        
        'FYM 2/28/19 Added Show Cart from vff
        slShowCart = "N"
        If Trim(rst_Temp!vffCartOnWeb) <> "" Then
            slShowCart = rst_Temp!vffCartOnWeb
        End If
        
    End If
    
    'D.S. 06/08/16 added web number
    'FYM  02/06/19 added Market and Rank
    SQLQuery = "SELECT * from shtt WHERE shttCode = " & iShttCode  'shttWebNumber, shttMarket, shttRank
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        ' JD TTP 10861
        slStr = slStr & Trim(rst_Temp!shttWebNumber)
        slStr = slStr & "," & """" & Trim(rst_Temp!shttMarket) & """," & rst_Temp!shttRank
    Else
        slStr = slStr & "1" & "," & """" & "" & """," & "0"
    End If
    slStr = slStr & ",""" & slShowCart & """"
    gBuildWebHeaders = slStr
'    rst_Gsf.Close
'    rst_Emt.Close
    
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gBuildWebHeaders"
    Exit Function
End Function


'***************************************************************************************
'*
'* Procedure Name: gVerifyWebIniSettings
'*
'* Created: 5/12/04 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************
Public Function gVerifyWebIniSettings() As Boolean
    Dim sBuffer As String

    gVerifyWebIniSettings = False
    ' Check the web site parameters 05-11-04 JD
    If Not gLoadOption(sgWebServerSection, "RootURL", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'RootURL' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "RegSection", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'RegSection' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "WebExports", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'WebExports' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "WebImports", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'WebImports' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPAddress", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPAddress' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPPort", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPPort' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPImportDir", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPImportDir' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPExportDir", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPExportDir' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPUID", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPUID' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPPWD", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPPWD' key is missing.", vbCritical
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPIsOn", sBuffer) Then
        gMsgBox "Affiliat.Ini [WebServer] 'FTPIsOn' key is missing.", vbCritical
        Exit Function
    End If
    gVerifyWebIniSettings = True
    
    
End Function


'***************************************************************************************
'*
'* Procedure Name: gTestAccessToWebServer
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function tests to see whether this PC has access to the web server.
'*           This function should be called one time only during startup and will set
'*           the gWebAccessTestedOk variable. From then on, you should call the function
'*           gHasWebAccess to test whether or not to perform a web command.
'*
'***************************************************************************************
Public Function gTestAccessToWebServer() As Boolean
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

    gTestAccessToWebServer = False
    gWebAccessTestedOk = False
    If Not gUsingWeb Then
        ' If this is not turned on in the site options table then there for sure will not be any
        ' web access calls made.
        Exit Function
    End If
    
    Call gLoadOption(sgWebServerSection, "WebLogLevel", slWebLogLevel)
    If slWebLogLevel = "" Then
        slWebLogLevel = "0"
    End If
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If Not gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        Exit Function
    End If
    ' Make a request to view the headers. If this operation succeeds, then all is ok.
    ' NOTE: The VH.ASP page will not return any data from the header table because the password here
    ' is not correct due to the date and time being added. But this is ok because it will return the
    ' text that the password is invalid and access to the web page itself returns success.
    slWebPage = slRootURL & "Help_Index.htm?" & Now()
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
    
    gTestAccessToWebServer = True
    gWebAccessTestedOk = True
    Exit Function

ERR_gTestAccessToWebServer:
    Exit Function
    ' Exit the function if any errors occur

End Function

Public Function gWebUpdateAccessControl() As Boolean

    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slWebPage As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slUpdateAccessControl As String

    'D.S. 09/23/16  This was written for the code that Phil and Greg wrote and is no longer going to be implemented.
    'Its purpose was to make a call to sync the the first tier DB credentials with the second tier DB
    gWebUpdateAccessControl = True
    Exit Function
    
'    '3/4/15
'    'D.S. 2/4/15 Until the new web is ready just return true and get out.  The new web will start at version 7.0
'    'If sgWebSiteVersion < 7 Then
'    If CDbl(sgWebSiteVersion) < 7 Then
'        gWebUpdateAccessControl = True
'        Exit Function
'    End If
'
'    On Error GoTo ERR_gWebUpdateAccessControl
'    gWebUpdateAccessControl = False
'    If Not gUsingWeb Then
'        ' If this is not turned on in the site options table then there for sure will not be any
'        ' web access calls made.
'        Exit Function
'    End If
'
'    If gGetEMailDistribution Then
'        If Not gLoadOption("WebServer", "UpdateAccessControl", slUpdateAccessControl) Then
'            gLogMsg "Error: modWebSubs - gWebUpdateAccessControl: LoadOption UpdateAccessControl Error", "AffWebErrorLog.Txt", False
'            Exit Function
'        End If
'    End If
'
'    If Not gLoadOption("WebServer", "RootURL", slRootURL) Then
'        gLogMsg "Error: modWebSubs - gWebUpdateAccessControl: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
'        Exit Function
'    End If
'    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
'    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
'    ' registry to gather additional information. This is necessary to run multiple databases on the
'    ' same IIS platform. The password is hardcoded and never changes.
'    If Not gLoadOption("WebServer", "RegSection", slRegSection) Then
'        gLogMsg "Error: modWebSubs - gWebUpdateAccessControl: LoadOption RegSection Error", "AffWebErrorLog.Txt", False
'        Exit Function
'    End If
'    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
'    objXMLHTTP.Open "GET", slUpdateAccessControl, False
'    objXMLHTTP.Send
'    llReturn = objXMLHTTP.Status
'    slResponse = objXMLHTTP.responseText
'    Set objXMLHTTP = Nothing
'    ' Very the return code here. Anything but 200 is an error.
'    If llReturn <> 200 Then
'        Exit Function
'    End If
'    gWebUpdateAccessControl = True
'    Exit Function

ERR_gWebUpdateAccessControl:
    Exit Function
    ' Exit the function if any errors occur

End Function



'***************************************************************************************
'*
'* Procedure Name: gHasWebAccess
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Call gTestAccessToWebServer before calling this function.
'*           gTestAccessToWebServer should be called one time only and will set the gHasWebAccess
'*           variable to its proper state.
'*
'***************************************************************************************
Public Function gHasWebAccess() As Boolean
    gHasWebAccess = False
    '10000
'    If igTestSystem Then
    If igDemoMode Then
        Exit Function
    End If
    If Not gUsingWeb Then
        ' This site does not have the Using Web Server turned on in the Site Options.
        Exit Function
    End If
    If Not gWebAccessTestedOk Then
        ' The function gTestAccessToWebServer was not able to access the web server and therfore
        ' this session should not attempt to make any web calls.
        Exit Function
    End If
    gHasWebAccess = True
End Function


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
Public Function gSendCmdToWebServer(sWebPageToAccess As String, sFileName As String)
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
    Dim WebCmds As New WebCommands

    gSendCmdToWebServer = False
    If Not gHasWebAccess() Then
        gSendCmdToWebServer = True
        Exit Function
    End If
    Call gLoadOption(sgWebServerSection, "WebLogLevel", slWebLogLevel)
    If slWebLogLevel = "" Then
        slWebLogLevel = "0"
    End If
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gSendCmdToWebServer: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gSendCmdToWebServer: LoadOption RootURL Error"
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & sWebPageToAccess & "?PWData?PW=jfdl&RK=" & Trim(slRegSection) & "&P1=" & slWebLogLevel & sFileName
    Else
        ' Here we provide for backward compatibility incase the parameter is missing. If running
        ' with the new ISAPI extensions, they will Error with an error if no parameters are supplied.
        slISAPIExtensionDLL = slRootURL & sWebPageToAccess
    End If

    For ilRetries = 0 To 14
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
            'gLogMsg "gSendCmdToWebServer is retrying 1", "AffWebSubsRetryLog.Txt", False
            Call gSleep(2)
        End If
    Next ilRetries

    ' We were never successful if we make it to here.
    gLogMsg "Error - gSendCmdToWebServer, retries were exceeded.", "AffWebErrorLog.txt", False
    gLogMsg "   " & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", "AffWebErrorLog.Txt", False

    Exit Function

ERR_gSendCmdToWebServer:
    'llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-gExecWebSQL: "
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.Txt", False
    Resume Next
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
    Dim ServerFileName As String
    Dim ilMaxRetries As Integer
    Dim ilRetries As Integer
    Dim iRet As Integer

    ilMaxRetries = 4

    If Not gHasWebAccess() Then
        gFTPFileToWebServer = True
        Exit Function
    End If

    gFTPFileToWebServer = False
    ' First load all the information we need from the ini file.
    Call gLoadOption(sgWebServerSection, "FTPIsOn", FTPIsOn)
    If Val(FTPIsOn) < 1 Then
        ' FTP is turned off. Return success.
        ' Note: This will be the case when the affiliate system and IIS is running on the same machine.
        '       Usually only while testing.
        gFTPFileToWebServer = True
        Exit Function
    End If

    If Not gLoadOption(sgWebServerSection, "FTPAddress", FTPAddress) Then
        gLogMsg "Error: gFTPFileToWebServer: LoadOption FTPAddress Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gFTPFileToWebServer: LoadOption FTPAddress Error"
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPPort", FTPPort) Then
        gLogMsg "Error: gFTPFileToWebServer: LoadOption FTPPort Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gFTPFileToWebServer: LoadOption FTPPort Error"
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPUID ", FTPUID) Then
        gLogMsg "Error: gFTPFileToWebServer: LoadOption FTPUID Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gFTPFileToWebServer: LoadOption FTPUID Error"
        Exit Function
    End If
    If Not gLoadOption(sgWebServerSection, "FTPPWD ", FTPPWD) Then
        gLogMsg "Error: gFTPFileToWebServer: LoadOption FTPPWD Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gFTPFileToWebServer: LoadOption FTPPWD Error"
        Exit Function
    End If
    
    'Debug Only - User Id and Password have to be blank for Jeff's FTP server
    If FTPUID = "BLANK" Then
        FTPUID = ""
    End If
    If FTPPWD = "BLANK" Then
        FTPPWD = ""
    End If
    
    If Not gLoadOption(sgWebServerSection, "FTPImportDir ", FTPWebDir) Then
        gLogMsg "Error: gFTPFileToWebServer: LoadOption FTPImportDir Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gFTPFileToWebServer: LoadOption FTPImportDir Error"
        Exit Function
    End If
    
    ' 02-01-21 Jeff D.
    ' Use FTP in CSI_Utils instead so it can also be sent via TCP/IP sockets.
    '
    iRet = mInitFTP()
    iRet = csiFTPInit(tmCsiFtpInfo)
    iRet = csiFTPFileToServer(sFileName)
    
    iRet = csiFTPGetStatus(tmCsiFtpStatus)
    ilRetries = 30
    '1 = Busy, 0 = Not Busy
    While (tmCsiFtpStatus.iState = 1 And ilRetries > 0)
        Sleep (200) ' 200 * 30 = 6 seconds
        DoEvents
        iRet = csiFTPGetStatus(tmCsiFtpStatus)
        ilRetries = ilRetries - 1
    Wend
    If tmCsiFtpStatus.iState = 0 Then
        gFTPFileToWebServer = True
    End If
    Exit Function
    
'    ' Open internet
'    For ilRetries = 0 To ilMaxRetries
'    hINetSession = InternetOpen("CSI_Affiliate", 0, vbNullString, vbNullString, 0)
'    If hINetSession < 1 Then
'            ilRetries = ilRetries + 1
'            gSleep (1)
'        Else
'            Exit For
'        End If
'    Next ilRetries
'    If ilRetries = ilMaxRetries Then
'        gMsgBox "Error: gFTPFileFromWebServer: InternetOpen Error"
'        gLogMsg "Error: gFTPFileFromWebServer: InternetOpen Error", "AffWebErrorLog.txt", False
'        Exit Function
'    End If
'
'    ' Connect to the internet
'    For ilRetries = 0 To ilMaxRetries
'    'hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, 0, 0)
'    hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'
'        If hINetSession < 1 Then
'            ilRetries = ilRetries + 1
'        Else
'            Exit For
'        End If
'    Next ilRetries
'    If ilRetries = ilMaxRetries Then
'        Call InternetCloseHandle(hINetSession)
'        gMsgBox "Error: gFTPFileFromWebServer: InternetConnect Error"
'        gLogMsg "Error: gFTPFileFromWebServer: InternetConnect Error", "AffWebErrorLog.txt", False
'        Exit Function
'    End If
'
'    ' Send the data to the server
'    FTPWebDir = gSetPathEndSlash(FTPWebDir, False)
'    ServerFileName = FTPWebDir & sFileName
'    For ilRetries = 0 To ilMaxRetries
'    If FtpPutFile(hSession, sPathFileName, ServerFileName, 1, 0) = False Then
'            ilRetries = ilRetries + 1
'        Else
'            Exit For
'        End If
'    Next ilRetries
'    If ilRetries = ilMaxRetries Then
'        gLogMsg "Error: gFTPFileToWebServer: FtpPutFile Error. " & sPathFileName & ", " & ServerFileName, "AffWebErrorLog.Txt", False
'        gMsgBox "Error: gFTPFileToWebServer: FtpPutFile Error. " & sPathFileName & ", " & ServerFileName
'        Call InternetCloseHandle(hSession)
'        Call InternetCloseHandle(hINetSession)
'        Exit Function
'    End If
'
'    Call InternetCloseHandle(hSession)
'    Call InternetCloseHandle(hINetSession)
'
'    gFTPFileToWebServer = True
'    Exit Function

ERR_gFTPFileToWebServer:
    ' Exit the function if any errors occur
    Exit Function
End Function

Private Function mInitFTP() As Boolean

    Dim slTemp As String
    Dim ilRet As Integer
    Dim slSection As String
    
    mInitFTP = False
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    
    'Support for CSI_Utils FTP functions
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tmCsiFtpInfo.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tmCsiFtpInfo.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tmCsiFtpInfo.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tmCsiFtpInfo.sPWD)
    Call gLoadOption(sgWebServerSection, "WebExports", tmCsiFtpInfo.sSendFolder)
    Call gLoadOption(sgWebServerSection, "WebImports", tmCsiFtpInfo.sRecvFolder)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tmCsiFtpInfo.sServerDstFolder)
    Call gLoadOption(sgWebServerSection, "FTPExportDir", tmCsiFtpInfo.sServerSrcFolder)
    Call gLoadOption(slSection, "DBPath", tmCsiFtpInfo.sLogPathName)
    tmCsiFtpInfo.sLogPathName = Trim$(tmCsiFtpInfo.sLogPathName) & "\" & "Messages\FTPLog.txt"
    ilRet = csiFTPInit(tmCsiFtpInfo)
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tgCsiFtpFileListing.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tgCsiFtpFileListing.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tgCsiFtpFileListing.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tgCsiFtpFileListing.sPWD)
    Call gLoadOption(slSection, "DBPath", tgCsiFtpFileListing.sLogPathName)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tgCsiFtpFileListing.sPathFileMask)
    Exit Function
End Function

'***************************************************************************************
'*
'* Procedure Name: gFTPFileFromWebServer
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: This function obtains all necessary information from the ini file to
'*           FTP the specified file to the Web Server.
'*
'***************************************************************************************
Public Function gFTPFileFromWebServer(sPathFileName As String, sFileName As String) As Boolean
    On Error GoTo ERR_gFTPFileFromWebServer
    Dim hINetSession As Long
    Dim hSession As Long
    Dim FTPIsOn As String
    Dim FTPAddress As String
    Dim FTPPort As String
    Dim FTPUID As String
    Dim FTPPWD As String
    Dim FTPWebDir As String
    Dim ServerFileName As String
    Dim ilMaxRetries As Integer
    Dim ilRetries As Integer

    ilMaxRetries = 4

    If Not gHasWebAccess() Then
        gFTPFileFromWebServer = True
        Exit Function
    End If
    
    gFTPFileFromWebServer = False
    ' First load all the information we need from the ini file.
    Call gLoadOption(sgWebServerSection, "FTPIsOn", FTPIsOn)
    If Val(FTPIsOn) < 1 Then
        ' FTP is turned off. Return success.
        ' Note: This will be the case when the affiliate system and IIS is running on the same machine.
        '       Usually only while testing.
        gFTPFileFromWebServer = True
        Exit Function
    End If

    If Not gLoadOption(sgWebServerSection, "FTPAddress", FTPAddress) Then
        gMsgBox "Error: gFTPFileFromWebServer: LoadOption FTPAddress Error"
        gLogMsg "Error: gFTPFileFromWebServer: LoadOption FTPAddress Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    If Not gLoadOption(sgWebServerSection, "FTPPort", FTPPort) Then
        gMsgBox "Error: gFTPFileFromWebServer: LoadOption FTPPort Error"
        gLogMsg "Error: gFTPFileFromWebServer: LoadOption FTPPort Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    If Not gLoadOption(sgWebServerSection, "FTPUID ", FTPUID) Then
        gMsgBox "Error: gFTPFileFromWebServer: LoadOption FTPUID Error"
        gLogMsg "Error: gFTPFileFromWebServer: LoadOption FTPUID Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    If Not gLoadOption(sgWebServerSection, "FTPPWD ", FTPPWD) Then
        gMsgBox "Error: gFTPFileFromWebServer: LoadOption FTPPWD Error"
        gLogMsg "Error: gFTPFileFromWebServer: LoadOption FTPPWD Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    'Debug Only - User Id and Password have to be blank for Jeff's FTP server
    If FTPUID = "BLANK" Then
        FTPUID = ""
    End If
    If FTPPWD = "BLANK" Then
        FTPPWD = ""
    End If
    
    If Not gLoadOption(sgWebServerSection, "FTPExportDir ", FTPWebDir) Then
        gMsgBox "Error: gFTPFileFromWebServer: LoadOption FTPExportDir Error"
        gLogMsg "Error: gFTPFileFromWebServer: LoadOption FTPExportDir Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    ' Open internet
    For ilRetries = 0 To ilMaxRetries
    hINetSession = InternetOpen("CSI_Affiliate", 0, vbNullString, vbNullString, 0)
    If hINetSession < 1 Then
            ilRetries = ilRetries + 1
            gSleep (1)
        Else
            Exit For
        End If
    Next ilRetries
    If ilRetries = ilMaxRetries Then
        gMsgBox "Error: gFTPFileFromWebServer: InternetOpen Error"
        gLogMsg "Error: gFTPFileFromWebServer: InternetOpen Error", "AffWebErrorLog.txt", False
        Exit Function
    End If
    
    
    ' Connect to the internet
    For ilRetries = 0 To ilMaxRetries
    'hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, 0, 0)
    hSession = InternetConnect(hINetSession, FTPAddress, FTPPort, FTPUID, FTPPWD, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
    
        If hSession < 1 Then
            ilRetries = ilRetries + 1
        Else
            Exit For
        End If
    Next ilRetries
    If ilRetries = ilMaxRetries Then
        Call InternetCloseHandle(hINetSession)
        gMsgBox "Error: gFTPFileFromWebServer: InternetConnect Error"
        gLogMsg "Error: gFTPFileFromWebServer: InternetConnect Error", "AffWebErrorLog.txt", False
        Exit Function
    End If

    ' Receive the data from the server
    FTPWebDir = gSetPathEndSlash(FTPWebDir, False)
    ServerFileName = FTPWebDir & sFileName
    For ilRetries = 0 To ilMaxRetries
    If FtpGetFile(hSession, ServerFileName, sPathFileName, False, 0, 1, 0) = False Then
            gSleep (2)
        Else
            Exit For
        End If
    Next ilRetries
    If ilRetries >= ilMaxRetries Then
        gLogMsg "Error: FtpGetFile Error. " & sPathFileName & ", " & ServerFileName, "AffWebErrorLog.txt", False
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
    Exit Function
End Function

Public Function gRemoteExecSql(mSqlStr As String, sFileName As String, mIniValue As String, iKill As Integer, iWriteFile, mNumRetries As Integer) As Boolean

    On Error GoTo ERR_gRemoteExecSql
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slWebLogLevel As String
    Dim ilRet As Integer
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim WebCmds As New WebCommands
    
    If Not gHasWebAccess() Then
        gRemoteExecSql = True
        Exit Function
    End If
    gRemoteExecSql = False
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gRemoteExecSql: LoadOption RootURL Error", "AffWebErrorLog.txt", False
        gMsgBox "Error: gRemoteExecSql: LoadOption RootURL Error"
        Exit Function
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    mSqlStr = gUrlEncoding(mSqlStr)
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & mSqlStr
    End If

    'If Not gExecXMLHTTPRequest(slISAPIExtensionDLL) Then
    '    Exit Function
    'End If
    'We will retry every 2 seconds and wait up to mNumRetries
    For ilRetries = 0 To mNumRetries
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(mSqlStr)
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
        
        ' Very the return code here. Anything but 200 is an error.
        ' Also if the error handler is called llErrorCode will be set.
        If llReturn = 200 Then
            If iWriteFile Then
                gRemoteSqlResults slResponse, sFileName, mIniValue, iKill
            End If
            gRemoteExecSql = True
            Exit Function
        Else
            DoEvents
            Call gSleep(2)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "gRemoteExecSql, retries were exceeded. SQL = " & mSqlStr, "AffWebErrorLog.txt", False
    Exit Function

ERR_gRemoteExecSql:
    'llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-gRemoteExecSQL: "
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.txt", False
    Resume Next
End Function


Public Sub gRemoteSqlResults(sMsg As String, sFileName As String, sIniValue As String, iKill As Integer)
    'D.S. 6/04
    'Purpose: A general file routine that writes the results from a reomte sql call
    '         for processing later
    
    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'sIniValue is the name of the entry in the the Affiliat.ini used for a path
    'iKill = True then delete the file first, iKill = False then append to the file
    
    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slToFile As String
    Dim slDateTime As String

    If Not gHasWebAccess() Then
        Exit Sub
    End If

    Call gLoadOption(sgWebServerSection, sIniValue, slToFile)
    slToFile = gSetPathEndSlash(slToFile, True)
    slToFile = slToFile & sFileName
    'On Error GoTo Error
    If iKill = True Then
        ilRet = 0
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
    End If
    
    'hlLogFile = FreeFile
    'Open slToFile For Append As hlLogFile
    ilRet = gFileOpen(slToFile, "Append", hlLogFile)
    If ilRet = 0 Then
        slFullMsg = sMsg
        Print #hlLogFile, slFullMsg
    End If
    Close hlLogFile
    Exit Sub
    
'Error:
'    ilRet = 1
'    Resume Next
End Sub

Public Function gRemoteProcessPWResults(sFileName As String, sIniValue As String) As Long
    
    'D.S. 6/04
    'Purpose:

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim slLocation As String
    Dim slTemp As String
    Dim ilLineNumber As Integer
    Dim hlFrom As Integer
    Dim slLine As String
    Dim ilRet  As Integer
    Dim ilIdx As Integer
    Dim slAttCode As String
    Dim slStationName As String
    Dim slOldStationPW As String
    Dim slNewStationPW As String
    Dim slOldAgreementPW As String
    Dim slNewAgreementPW As String
    'Dim slLine As String
    ''Dim slFields(1 To 6) As String
    'Dim slFields(0 To 5) As String
    Dim slFields() As String
    Dim ilField As Integer
    Dim blProcessLine As Boolean
    
    If Not gHasWebAccess() Then
        gRemoteProcessPWResults = True
        Exit Function
    End If
    
    On Error GoTo ErrHand
    
    lgLine = 0
    gRemoteProcessPWResults = True
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    'slLocation = gSetPathEndSlash(slLocation, True)
    If (StrComp(sIniValue, "WebImports", vbTextCompare) = 0) Or (StrComp(sIniValue, "WebExports", vbTextCompare) = 0) Then
        slLocation = gSetPathEndSlash(slLocation, True)
    Else
        slLocation = gSetPathEndSlash(slLocation, False)
    End If
    slLocation = slLocation & sFileName
    
    On Error GoTo FileErrHand:
    'hlFrom = FreeFile
    ilRet = 0
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        gMsgBox "Error: modWebSubsg-RemoteProcessPWResults was unable to open the file."
        GoTo ErrHand
        Exit Function
    End If
    
    'Move past the header information
    ilIdx = 0
    Input #hlFrom, slAttCode, slStationName, slOldStationPW, slNewStationPW, slOldAgreementPW, slNewAgreementPW
    lgLine = lgLine + 1
    While Not EOF(hlFrom)
        DoEvents
        ''If the slAttCode is not numeric then we have a problem
        'Input #hlFrom, slAttCode, slStationName, slOldStationPW, slNewStationPW, slOldAgreementPW, slNewAgreementPW
        Line Input #hlFrom, slLine
        blProcessLine = False
        If Len(slLine) > 0 Then
            slFields = Split(slLine, ",")
            If IsArray(slFields) Then
                blProcessLine = True
                For ilField = 0 To UBound(slFields)
                    Select Case ilField
                        Case 0
                             slAttCode = gGetDataNoQuotes(slFields(ilField))
                        Case 1
                            slStationName = gGetDataNoQuotes(slFields(ilField))
                        Case 2
                            slOldStationPW = gGetDataNoQuotes(slFields(ilField))
                        Case 3
                            slNewStationPW = gGetDataNoQuotes(slFields(ilField))
                        Case 4
                            slOldAgreementPW = gGetDataNoQuotes(slFields(ilField))
                        Case 5
                            slNewAgreementPW = gGetDataNoQuotes(slFields(ilField))
                    End Select
                Next ilField
            End If
            blProcessLine = True
        End If
        If blProcessLine Then
            If IsNumeric(CLng(slAttCode)) Then
                lgLine = lgLine + 1
                'Process the agreement passwords
                ilIdx = ilIdx + 1
                SQLQuery = "UPDATE att SET attWebPW = '" & Trim$(slNewAgreementPW) & "'"
                SQLQuery = SQLQuery + " WHERE (attCode= " & slAttCode & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffWebErrorLog.txt", "modWebSubs-gRemoteProcessPWResults"
                    gRemoteProcessPWResults = False
                    Close hlFrom
                    Exit Function
                End If
                'Process the agreement passwords
                SQLQuery = "UPDATE shtt SET shttWebPW = '" & Trim$(slNewStationPW) & "'"
                SQLQuery = SQLQuery + " WHERE (shttCallLetters = '" & Trim$(slStationName) & "')"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffWebErrorLog.txt", "modWebSubs-gRemoteProcessPWResults"
                    gRemoteProcessPWResults = False
                    Close hlFrom
                    Exit Function
                End If
                '11/26/17
                mUpdateShttTables slStationName, slNewStationPW
            End If
        End If
    Wend
    Erase slFields
    Close hlFrom
    Exit Function
    
FileErrHand:
    Close hlFrom
    If ilIdx = 0 Then
        gRemoteProcessPWResults = False
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in modWebSubs-gRemoteProcessPWResults: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffWebErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Function

'Public Function gRemoteProcessEmailResults(sFileName As String, sIniValue As String) As Long
'
'    'D.S. 1/3/11
'    'Purpose: Convert from using the old EMT file to the new ARTT file
'
'    Dim tlTxtStream As TextStream
'    Dim fs As New FileSystemObject
'    Dim slRetString As String
'    Dim slLocation As String
'    Dim slTemp As String
'    Dim ilLineNumber As Integer
'    Dim hlFrom As Integer
'    Dim slLine As String
'    Dim ilRet  As Integer
'    Dim ilIdx As Integer
'    Dim slCode As String
'    Dim slStationName As String
'    Dim slOldStationEmail As String
'    Dim slNewStationEmail As String
'    Dim slOldAgreementEmail As String
'    Dim slNewAgreementEmail As String
'    'Dim slLine As String
'    'Dim slFields(1 To 5) As String
'    Dim slShttCode As String
'    Dim ilShttCode As Integer
'    Dim slSeqNo As String
'    Dim slEMail As String
'    Dim slStatus As String
'    Dim ilRowsEffected As Integer
'    Dim max_rst As ADODB.Recordset
'    Dim slCallLetters As String
'    Dim slFirstName As String
'    Dim slLastName As String
'    Dim slTitle As String
'    Dim iltntCode As Integer
'    Dim slFields() As String
'    Dim ilField As Integer
'    Dim blProcessLine As Boolean
'
'
'    If Not gHasWebAccess() Then
'        gRemoteProcessEmailResults = True
'        Exit Function
'    End If
'
'    On Error GoTo ErrHand
'
'    lgLine = 0
'    gRemoteProcessEmailResults = True
'    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
'    'slLocation = gSetPathEndSlash(slLocation, False)
'    If (StrComp(sIniValue, "WebImports", vbTextCompare) = 0) Or (StrComp(sIniValue, "WebExports", vbTextCompare) = 0) Then
'        slLocation = gSetPathEndSlash(slLocation, True)
'    Else
'        slLocation = gSetPathEndSlash(slLocation, False)
'    End If
'    slLocation = slLocation & sFileName
'
'    On Error GoTo FileErrHand:
'    'hlFrom = FreeFile
'    ilRet = 0
'    'Open slLocation For Input Access Read As hlFrom
'    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
'    If ilRet <> 0 Then
'        gMsgBox "Error: modWebSubsg-RemoteProcessEmailResults was unable to open the file."
'        GoTo ErrHand
'        Exit Function
'    End If
'
'    'Move past the header information
'    ilIdx = 0
'    Input #hlFrom, slCode, slShttCode, slSeqNo, slEMail, slStatus, slCallLetters, slFirstName, slLastName, slTitle
'    lgLine = lgLine + 1
'    While Not EOF(hlFrom)
'        DoEvents
'        'Input #hlFrom, slCode, slShttCode, slSeqNo, slEMail, slStatus, slCallLetters, slFirstName, slLastName, slTitle
'        Line Input #hlFrom, slLine
'        blProcessLine = False
'        If Len(slLine) > 0 Then
'            slFields = Split(slLine, ",")
'            If IsArray(slFields) Then
'                blProcessLine = True
'                For ilField = 0 To UBound(slFields)
'                    Select Case ilField
'                        Case 0
'                            slCode = slFields(ilField)
'                        Case 1
'                            slShttCode = slFields(ilField)
'                        Case 2
'                            slSeqNo = slFields(ilField)
'                        Case 3
'                            slEMail = slFields(ilField)
'                        Case 4
'                            slStatus = slFields(ilField)
'                        Case 5
'                            slCallLetters = slFields(ilField)
'                        Case 6
'                            slFirstName = slFields(ilField)
'                        Case 7
'                            slLastName = slFields(ilField)
'                        Case 8
'                            slTitle = slFields(ilField)
'                    End Select
'                Next ilField
'            End If
'            blProcessLine = True
'        End If
'        If blProcessLine Then
'            'Note: if slCode = 90,000,000 or more it designates that the record was newly added on the web
'            'It could have been added and then edited or deleted afterwards.
'
'            'If it's greater than 90000000 and the status is a "E" then it's because
'            'it was added on the web then edited so, it's still new to the Aff database
'            If CLng(slCode) >= 90000000 And slStatus = "E" Then
'                slStatus = "A"
'            End If
'
'            'If it's greater than 900000 and the status is a "D" then it's because
'            'it was added on the web then deleted so, we don't need to insert in the Aff database
'            'If CLng(slCode) >= 90000 And slStatus = "D" Then
'            '    slStatus = "Z"
'            'End If
'
'            'Do nothing, it's Jeff's bogus station that does not exist
'            If Trim$(slCallLetters) = "TEST-FM" Then
'                slStatus = "Z"
'            End If
'
'            If IsNumeric(CLng(slShttCode)) Then
'                iltntCode = gGetTntCodeByTitle(slTitle)
'                ilShttCode = gGetShttCodeFromCallLetters(slCallLetters)
'                slShttCode = ilShttCode
'                lgLine = lgLine + 1
'                ilIdx = ilIdx + 1
'                If slStatus = "E" Then
'                    SQLQuery = "Update artt Set "
'                    SQLQuery = SQLQuery & "arttEmail = '" & gFixQuote(Trim$(slEMail)) & "',"
'                    SQLQuery = SQLQuery & "arttEmailToWeb = '" & " " & "',"
'                    SQLQuery = SQLQuery & "arttFirstName = '" & gFixQuote(Trim$(slFirstName)) & "',"
'                    SQLQuery = SQLQuery & "arttLastName = '" & gFixQuote(Trim$(slLastName)) & "',"
'                    SQLQuery = SQLQuery & "arttTntCode =  " & iltntCode & " ,"
'                    SQLQuery = SQLQuery & " arttWebEMailRefID = '" & slSeqNo & "'"
'                    SQLQuery = SQLQuery & " Where arttshttCode = " & slShttCode & " And arttWebEMailRefID = " & slSeqNo & " And arttType = " & "'" & "P" & "'"
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        '6/13/16: Replaced GoSub
'                        'GoSub ErrHand:
'                        Screen.MousePointer = vbDefault
'                        gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
'                        gRemoteProcessEmailResults = False
'                        On Error Resume Next
'                        Close hlFrom
'                        Exit Function
'                    End If
'
'                    SQLQuery = "UPDATE WebEmt SET "
'                    SQLQuery = SQLQuery + "Status = ' '" & ","
'                    SQLQuery = SQLQuery + "DateModified = " & "'" & Format(Now, "ddddd ttttt") & "' "
'                    SQLQuery = SQLQuery + " WHERE (ShttCode = " & slShttCode
'                    SQLQuery = SQLQuery + " AND SeqNo = " & slSeqNo & ")"
'                    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
'                    If ilRowsEffected = -1 Then
'                        gLogMsg "Error: Auto Update Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
'                    End If
'
'                End If
'
'                If slStatus = "D" Then
'                    SQLQuery = "Delete From artt Where arttWebEMailRefID = " & slSeqNo & " And arttShttCode = " & slShttCode & " And arttType = " & "'" & "P" & "'"
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        '6/13/16: Replaced GoSub
'                        'GoSub ErrHand:
'                        Screen.MousePointer = vbDefault
'                        gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
'                        gRemoteProcessEmailResults = False
'                        On Error Resume Next
'                        Close hlFrom
'                        Exit Function
'                    End If
'
'                    SQLQuery = "Delete From webEmt Where Code = " & slCode
'                    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
'                    If ilRowsEffected = -1 Then
'                        gLogMsg "Error: Auto Delete Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
'                    End If
'                End If
'
'                If slStatus = "A" Then
'                    If slShttCode = 0 Then
'                        slShttCode = gGetShttCodeFromCallLetters(slCallLetters)
'                    End If
'                    SQLQuery = "Select Max(arttWebEMailRefID) From artt Where arttShttCode = " & slShttCode & " And arttType = " & "'" & "P" & "'"
'                    Set max_rst = gSQLSelectCall(SQLQuery)
'
'                    If IsNull(max_rst(0).Value) Then
'                        slSeqNo = 1
'                    Else
'                        slSeqNo = max_rst(0).Value + 1
'                    End If
'
'                    iltntCode = gGetTntCodeByTitle(slTitle)
'
'                    'Insert into the local Affiliate
'                    SQLQuery = "Insert Into artt ( "
'                    SQLQuery = SQLQuery & "arttCode, "
'                    SQLQuery = SQLQuery & "arttShttCode, "
'                    SQLQuery = SQLQuery & "arttWebEMailRefID, "
'                    SQLQuery = SQLQuery & "arttEMail, "
'                    SQLQuery = SQLQuery & "arttEMailRights, "
'                    SQLQuery = SQLQuery & "arttEmailToWeb, "
'                    SQLQuery = SQLQuery & "arttType, "
'                    SQLQuery = SQLQuery & "arttWebEmail, "
'                    SQLQuery = SQLQuery & "arttFirstName, "
'                    SQLQuery = SQLQuery & "arttLastName, "
'                    SQLQuery = SQLQuery & "arttTntCode, "
'                    SQLQuery = SQLQuery & "arttUnused "
'                    SQLQuery = SQLQuery & ") "
'                    SQLQuery = SQLQuery & "Values ( "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & slShttCode & ", "
'                    SQLQuery = SQLQuery & slSeqNo & ", "
'                    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slEMail)) & "', "
'                    SQLQuery = SQLQuery & "'" & "N" & "', "
'                    'Mark the email as Sent as the web already has it
'                    SQLQuery = SQLQuery & "'" & "S" & "', "
'                    SQLQuery = SQLQuery & "'" & "P" & "', "
'                    'Mark the web email as yes, it's a web email
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slFirstName)) & "', "
'                    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slLastName)) & "', "
'                    SQLQuery = SQLQuery & iltntCode & ", "
'                    SQLQuery = SQLQuery & "'" & " " & "' "
'                    SQLQuery = SQLQuery & ") "
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        '6/13/16: Replaced GoSub
'                        'GoSub ErrHand:
'                        Screen.MousePointer = vbDefault
'                        gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
'                        gRemoteProcessEmailResults = False
'                        On Error Resume Next
'                        Close hlFrom
'                        Exit Function
'                    End If
'
'                    SQLQuery = "Select * From artt Where arttShttCode = " & slShttCode
'                    SQLQuery = SQLQuery + " AND arttWebEMailRefID = " & "'" & slSeqNo & "' "
'                    Set max_rst = gSQLSelectCall(SQLQuery)
'
'                    'Now update the web
'                    SQLQuery = "UPDATE WebEmt SET "
'                    SQLQuery = SQLQuery + "Code = " & max_rst!arttCode & ", "
'                    SQLQuery = SQLQuery + "Status = ' '" & ", "
'                    SQLQuery = SQLQuery + "DateModified = " & "'" & Format(Now, "ddddd ttttt") & "', "
'                    SQLQuery = SQLQuery + "SeqNo = " & CLng(slSeqNo) & ", "
'                    SQLQuery = SQLQuery + "ShttCode = " & slShttCode & " ,"
'                    SQLQuery = SQLQuery & "FirstName = " & "'" & gFixQuote(Trim$(slFirstName)) & "',"
'                    SQLQuery = SQLQuery & "LastName = " & "'" & gFixQuote(Trim$(slLastName)) & "',"
'                    SQLQuery = SQLQuery & "Title = " & "'" & gFixQuote(Trim$(slTitle)) & "'"
'                    SQLQuery = SQLQuery + " WHERE Code = " & slCode
'                    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
'                    If ilRowsEffected = -1 Then
'                        gLogMsg "Error: Auto Update Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
'                    End If
'                End If
'
'            End If
'        End If
'    Wend
'    Erase slFields
'    Close hlFrom
'    Exit Function
'
'FileErrHand:
'    Close hlFrom
'    If ilIdx = 0 Then
'        gRemoteProcessEmailResults = False
'    End If
'    Exit Function
'
'ErrHand:
'    Screen.MousePointer = vbDefault
'
'    If (Err.Number <> 0) And (gMsg = "") Then
'        gMsg = "A general error has occured in frmGenSubs-gRemoteProcessEmailResults: "
'        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "WebEmailInsertLog.Txt", False
'        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
'    End If
'End Function


Public Function gRemoteProcessEmailResults(sFileName As String, sIniValue As String) As Long
    
    'D.S. 1/3/11
    'Purpose: Convert from using the old EMT file to the new ARTT file

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slRetString As String
    Dim slLocation As String
    Dim slTemp As String
    Dim ilLineNumber As Integer
    Dim hlFrom As Integer
    Dim slLine As String
    Dim ilRet  As Integer
    Dim ilIdx As Integer
    Dim slCode As String
    Dim slStationName As String
    Dim slOldStationEmail As String
    Dim slNewStationEmail As String
    Dim slOldAgreementEmail As String
    Dim slNewAgreementEmail As String
    'Dim slLine As String
    Dim slFields(1 To 5) As String
    Dim slShttCode As String
    Dim ilShttCode As Integer
    Dim slSeqNo As String
    Dim slEMail As String
    Dim slStatus As String
    Dim ilRowsEffected As Integer
    Dim max_rst As ADODB.Recordset
    Dim slCallLetters As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slTitle As String
    Dim iltntCode As Integer
    
    If Not gHasWebAccess() Then
        gRemoteProcessEmailResults = True
        Exit Function
    End If
    
    On Error GoTo ErrHand
    
    lgLine = 0
    gRemoteProcessEmailResults = True
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    'slLocation = gSetPathEndSlash(slLocation, False)
    If (StrComp(sIniValue, "WebImports", vbTextCompare) = 0) Or (StrComp(sIniValue, "WebExports", vbTextCompare) = 0) Then
        slLocation = gSetPathEndSlash(slLocation, True)
    Else
        slLocation = gSetPathEndSlash(slLocation, False)
    End If
    slLocation = slLocation & sFileName
    
    On Error GoTo FileErrHand:
    hlFrom = FreeFile
    ilRet = 0
    Open slLocation For Input Access Read As hlFrom
    If ilRet <> 0 Then
        gMsgBox "Error: modWebSubsg-RemoteProcessEmailResults was unable to open the file."
        GoTo ErrHand
        Exit Function
    End If
    
    'Move past the header information
    ilIdx = 0
    Input #hlFrom, slCode, slShttCode, slSeqNo, slEMail, slStatus, slCallLetters, slFirstName, slLastName, slTitle
    lgLine = lgLine + 1
    While Not EOF(hlFrom)
        DoEvents
        Input #hlFrom, slCode, slShttCode, slSeqNo, slEMail, slStatus, slCallLetters, slFirstName, slLastName, slTitle
        
        'Note: if slCode = 90,000,000 or more it designates that the record was newly added on the web
        'It could have been added and then edited or deleted afterwards.
        
        'If it's greater than 90000000 and the status is a "E" then it's because
        'it was added on the web then edited so, it's still new to the Aff database
        If CLng(slCode) >= 90000000 And slStatus = "E" Then
            slStatus = "A"
        End If
        
        'If it's greater than 900000 and the status is a "D" then it's because
        'it was added on the web then deleted so, we don't need to insert in the Aff database
        'If CLng(slCode) >= 90000 And slStatus = "D" Then
        '    slStatus = "Z"
        'End If
        
        'Do nothing, it's Jeff's bogus station that does not exist
        If Trim$(slCallLetters) = "TEST-FM" Then
            slStatus = "Z"
        End If
        
        If IsNumeric(CLng(slShttCode)) Then
            iltntCode = gGetTntCodeByTitle(slTitle)
            ilShttCode = gGetShttCodeFromCallLetters(slCallLetters)
            slShttCode = ilShttCode
            lgLine = lgLine + 1
            ilIdx = ilIdx + 1
            If slStatus = "E" Then
                SQLQuery = "Update artt Set "
                SQLQuery = SQLQuery & "arttEmail = '" & gFixQuote(Trim$(slEMail)) & "',"
                SQLQuery = SQLQuery & "arttEmailToWeb = '" & " " & "',"
                SQLQuery = SQLQuery & "arttFirstName = '" & gFixQuote(Trim$(slFirstName)) & "',"
                SQLQuery = SQLQuery & "arttLastName = '" & gFixQuote(Trim$(slLastName)) & "',"
                SQLQuery = SQLQuery & "arttTntCode =  " & iltntCode & " ,"
                SQLQuery = SQLQuery & " arttWebEMailRefID = '" & slSeqNo & "'"
                SQLQuery = SQLQuery & " Where arttshttCode = " & slShttCode & " And arttWebEMailRefID = " & slSeqNo & " And arttType = " & "'" & "P" & "'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
                    gRemoteProcessEmailResults = False
                    On Error Resume Next
                    Close hlFrom
                    Exit Function
                End If
                
                SQLQuery = "UPDATE WebEmt SET "
                SQLQuery = SQLQuery + "Status = ' '" & ","
                SQLQuery = SQLQuery + "DateModified = " & "'" & Format(Now, "ddddd ttttt") & "' "
                SQLQuery = SQLQuery + " WHERE Code = " & slCode

                ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
                If ilRowsEffected = -1 Then
                    gLogMsg "Error: Auto Update Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
                End If
            
            End If
        
            If slStatus = "D" Then
                SQLQuery = "Delete From artt Where arttWebEMailRefID = " & slSeqNo & " And arttShttCode = " & slShttCode & " And arttType = " & "'" & "P" & "'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
                    gRemoteProcessEmailResults = False
                    On Error Resume Next
                    Close hlFrom
                    Exit Function
                End If
            
                SQLQuery = "Delete From webEmt Where Code = " & slCode
                ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
                If ilRowsEffected = -1 Then
                    gLogMsg "Error: Auto Delete Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
                End If
            End If
        
            If slStatus = "A" Then
                If slShttCode = 0 Then
                    slShttCode = gGetShttCodeFromCallLetters(slCallLetters)
                End If
                SQLQuery = "Select Max(arttWebEMailRefID) From artt Where arttShttCode = " & slShttCode & " And arttType = " & "'" & "P" & "'"
                Set max_rst = cnn.Execute(SQLQuery)
                
                If IsNull(max_rst(0).Value) Then
                    slSeqNo = 1
                Else
                    slSeqNo = max_rst(0).Value + 1
                End If
                
                iltntCode = gGetTntCodeByTitle(slTitle)
                
                'Insert into the local Affiliate
                SQLQuery = "Insert Into artt ( "
                SQLQuery = SQLQuery & "arttCode, "
                SQLQuery = SQLQuery & "arttShttCode, "
                SQLQuery = SQLQuery & "arttWebEMailRefID, "
                SQLQuery = SQLQuery & "arttEMail, "
                SQLQuery = SQLQuery & "arttEmailToWeb, "
                SQLQuery = SQLQuery & "arttType, "
                SQLQuery = SQLQuery & "arttWebEmail, "
                SQLQuery = SQLQuery & "arttFirstName, "
                SQLQuery = SQLQuery & "arttLastName, "
                SQLQuery = SQLQuery & "arttTntCode, "
                SQLQuery = SQLQuery & "arttUnused "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & 0 & ", "
                SQLQuery = SQLQuery & slShttCode & ", "
                SQLQuery = SQLQuery & slSeqNo & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slEMail)) & "', "
                'Mark the email as Sent as the web already has it
                SQLQuery = SQLQuery & "'" & "S" & "', "
                SQLQuery = SQLQuery & "'" & "P" & "', "
                'Mark the web email as yes, it's a web email
                SQLQuery = SQLQuery & "'" & "Y" & "', "
                SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slFirstName)) & "', "
                SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slLastName)) & "', "
                SQLQuery = SQLQuery & iltntCode & ", "
                SQLQuery = SQLQuery & "'" & " " & "' "
                SQLQuery = SQLQuery & ") "
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gRemoteProcessEmailResults"
                    gRemoteProcessEmailResults = False
                    On Error Resume Next
                    Close hlFrom
                    Exit Function
                End If
            
                SQLQuery = "Select * From artt Where arttShttCode = " & slShttCode
                SQLQuery = SQLQuery + " AND arttWebEMailRefID = " & "'" & slSeqNo & "' "
                Set max_rst = cnn.Execute(SQLQuery)
                
                'Now update the web
                SQLQuery = "UPDATE WebEmt SET "
                SQLQuery = SQLQuery + "Code = " & max_rst!arttCode & ", "
                SQLQuery = SQLQuery + "Status = ' '" & ", "
                SQLQuery = SQLQuery + "DateModified = " & "'" & Format(Now, "ddddd ttttt") & "', "
                SQLQuery = SQLQuery + "SeqNo = " & CLng(slSeqNo) & ", "
                SQLQuery = SQLQuery + "ShttCode = " & slShttCode & " ,"
                SQLQuery = SQLQuery & "FirstName = " & "'" & gFixQuote(Trim$(slFirstName)) & "',"
                SQLQuery = SQLQuery & "LastName = " & "'" & gFixQuote(Trim$(slLastName)) & "',"
                SQLQuery = SQLQuery & "Title = " & "'" & gFixQuote(Trim$(slTitle)) & "'"
                SQLQuery = SQLQuery + " WHERE Code = " & slCode
                ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
                If ilRowsEffected = -1 Then
                    gLogMsg "Error: Auto Update Error: " & SQLQuery, "WebEmailInsertLog.Txt", False
                End If
            End If
            
        End If
    Wend
    Close hlFrom
    Exit Function
    
FileErrHand:
    Close hlFrom
    If ilIdx = 0 Then
        gRemoteProcessEmailResults = False
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmGenSubs-gRemoteProcessEmailResults: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "WebEmailInsertLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Function



Public Sub gRemoteTestForNewWebPW()

    If gIsUsingNovelty Then
        Exit Sub
    End If
    
    Dim sFTPAddress As String
    Dim ilRet As Integer
    Dim slTemp As String
    Dim slFullPath As String
    Dim slLocation As String
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    If Not gHasWebAccess() Then
        Exit Sub
    End If
    
    Call gLoadOption(sgWebServerSection, "FTPAddress", sFTPAddress)
    frmProgressMsg.Show
    frmProgressMsg.SetMessage 0, "Importing Password Changes from the Website..." & vbCrLf & vbCrLf & "[" & sFTPAddress & "]"
    SQLQuery = "Select attCode, StationName, OldStationPW, NewStationPW, OldAgreementPW, NewAgreementPW from PWChanges"
    
    slTemp = "WebPasswords" & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    ilRet = gRemoteExecSql(SQLQuery, slTemp, "WebImports", True, True, 30)
    DoEvents
    If ilRet = True Then
        ilRet = gRemoteProcessPWResults(slTemp, "WebImports")
        lgLine = lgLine
    End If
    If ilRet = True Then
        SQLQuery = "DELETE FROM PWChanges"
        ilRet = gRemoteExecSql(SQLQuery, slTemp, "WebImports", True, False, 30)
    End If
    'slFullPath = sgImportDirectory & slTemp
    Call gLoadOption(sgWebServerSection, "WebImports", slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    slFullPath = slLocation & slTemp
    Kill slFullPath
    Unload frmProgressMsg

End Sub
Public Sub gRemoteTestForNewEmail()

    Dim sFTPAddress As String
    Dim ilRet As Integer
    Dim slTemp As String
    Dim slFullPath As String
    Dim slLocation As String
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    If Not gHasWebAccess() Then
        Exit Sub
    End If
    
    Call gLoadOption(sgWebServerSection, "FTPAddress", sFTPAddress)
    frmProgressMsg.Show
    frmProgressMsg.SetMessage 0, "Importing Email Changes from the Website..." & vbCrLf & vbCrLf & "[" & sFTPAddress & "]"
    SQLQuery = "Select Code, shttCode, seqNo, Email, Status, CallLetters, FirstName, LastName, Title from WebEmt where status <> ' ' Order By Code"
    
    slTemp = "WebEmail" & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    ilRet = gRemoteExecSql(SQLQuery, slTemp, "WebImports", True, True, 30)
    DoEvents
    If ilRet = True Then
        ilRet = gRemoteProcessEmailResults(slTemp, "WebImports")
        lgLine = lgLine
    End If
    
    Call gLoadOption(sgWebServerSection, "WebImports", slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    slFullPath = slLocation & slTemp
    Kill slFullPath
    Unload frmProgressMsg

End Sub


Public Function gExecExtStoredProc(sFileName As String, sExeToRun As String, iKill As Integer, iWriteFile) As Boolean

    On Error GoTo ERR_gExecExtStoredProc
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slWebLogLevel As String
    Dim ilRet As Integer
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim WebCmds As New WebCommands
    
    If Not gHasWebAccess() Then
        gExecExtStoredProc = True
        Exit Function
    End If
    
    If igDemoMode Then
        gExecExtStoredProc = True
        Exit Function
    End If
    
    gExecExtStoredProc = False
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gMsgBox "Error: gExecExtStoredProc: LoadOption RootURL Error"
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        'slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & mSqlStr
        'slISAPIExtensionDLL = slRootURL & "ExecExSP.dll?ExecExSP?PW=jfdl&RK=CSI_Web&SQL=Exec                   xp_cmdshell 'WebDispatcher.exe " & sFileName & "," & slRegSection & "," & sExeToRun & "'"
        slISAPIExtensionDLL = slRootURL & "ExecExSP.dll?ExecExSP?PW=jfdl&RK=" & slRegSection & "&" & "SQL=Exec xp_cmdshell " & "'" & "WebDispatcher.exe " & sFileName & "," & slRegSection & "," & sExeToRun & "'"
    End If
    
    For ilRetries = 0 To 14
        llReturn = 1
        If bgUsingSockets Then
            slResponse = WebCmds.RunEXE(sExeToRun, sFileName)
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
            If iWriteFile Then
                gRemoteSqlResults slResponse, sFileName, slRegSection, iKill
            End If
            gExecExtStoredProc = True
            Exit Function
        Else
            'gLogMsg "gExecExtStoredProc  is retrying", "WebSubsRetryLog.Txt", False
            Call gSleep(2)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "gExeExtStoredProc, retries were exceeded. slResponse = " & slResponse, "AffWebErrorLog.txt", False
    gLogMsg "gExecExtStoredProc Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", "AffWebErrorLog.Txt", False
    Exit Function
    
ERR_gExecExtStoredProc:
    'llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-gExecExtStoredProc: "
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.Txt", False
    Resume Next
End Function



' This function attempts to take control of the WebSession table by inserting an entry in it.
' If it can insert an entry using UKey = 1 then is gains control.
' If there is already an entry in the table it checks to make sure the PCName is not ours. Otherwise
' it means we have it locked already. If another PC has an entry in it, then the date and time
' is checked. If the date and time is greater than the amount specified, this function will take
' control of the web session anyway.
' This function returns one of two things.
' Either 0 meaning we now have control or a number indicating the total number of minutes to
' wait for the other PC.
Public Function gStartWebSession(LogFileName As String) As Integer
    On Error GoTo ErrHandler
    Dim slDateTime As String
    Dim slServersDateTime As String
    Dim slPCName As String
    Dim ilTotalTime As Integer
    Dim llTotalRecords As Long
    Dim slExportLockoutMinutes As String
    Dim slComputerName As String
    Dim alDataArray() As String
    Dim smMaxWebLockWaitMinutes As String
    
    On Error GoTo ErrHandler
    gStartWebSession = 0
    If Not gHasWebAccess() Then
        Exit Function
    End If

    Call gLoadOption(sgWebServerSection, "MaxWebLockWaitMinutes", smMaxWebLockWaitMinutes)
    If smMaxWebLockWaitMinutes = "" Then
        slExportLockoutMinutes = 60
    Else
        slExportLockoutMinutes = CInt(smMaxWebLockWaitMinutes)
    End If
    
    Call gLoadOption(sgWebServerSection, "ExportLockoutMinutes", slExportLockoutMinutes)

    llTotalRecords = gExecWebSQL(alDataArray, "Select DTStamp, PCName From WebSession Where UKey = 1", True)
    If llTotalRecords < 1 Then
        ' There are no records in the WebSession. Attempt to take control now.
        If Not gStartNewWebSession() Then
            ' This could only occur if someone obtained a lock right after we just looked at it.
            gStartWebSession = slExportLockoutMinutes   ' Try again later.
            Exit Function
        End If
        gStartWebSession = 0        ' We now have control.
        Exit Function
    End If

    ' If we made it here, the WebSession table has an entry in it so someone has it locked.
    ' Find out who it is and for how long it's been locked.

    ' Retrieve the two data items we got back from the SQL statement above.
    slDateTime = gGetDataNoQuotes(alDataArray(0))
    slPCName = gGetDataNoQuotes(alDataArray(1))

    slComputerName = gGetComputerName() & "_" & sgUserName
    If UCase(Trim$(slPCName)) = UCase(Trim$(slComputerName)) Then
        ' We're the one who has the record locked. This can happen if we exited abnormally.
        Call gEndWebSession(LogFileName)
        If Not gStartNewWebSession() Then
            gStartWebSession = Int(slExportLockoutMinutes)   ' Try again later.
            Exit Function
        End If
        gStartWebSession = 0        ' We now have control.
        Exit Function
    End If

    slServersDateTime = GetServersDateTime()
    ' Someone else has the export locked. check how long its been locked and if longer than
    ' the time specified allow this check to pass.
    ilTotalTime = DateDiff("n", slDateTime, slServersDateTime)
    If ilTotalTime >= Int(slExportLockoutMinutes) Then
        'D.S. 08/16/14 added new gKillTimedOutWebSession
        'Call gEndWebSession(LogFileName)
        Call gKillTimedOutWebSession(LogFileName)
        gLogMsg "Timed out waiting for web session.  Starting new session.", "AffWebErrorLog.Txt", False
        If Not gStartNewWebSession() Then
            gStartWebSession = Int(slExportLockoutMinutes)   ' Try again later.
            Exit Function
        End If
        gStartWebSession = 0        ' We now have control.
        Exit Function
    End If
    ' Return how much time is left to wait.
    gStartWebSession = Int(slExportLockoutMinutes) - ilTotalTime
    Exit Function

ErrHandler:
    gStartWebSession = 30
    'gMsg = "A general error has occured in modWebSubs-gStartWebSession: "
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, LogFileName, False
    'gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
End Function

Private Function gStartNewWebSession() As Boolean
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slURLResponse As String
    Dim slWebLogLevel As String
    Dim ilRet As Integer
    Dim slCommand As String
    Dim WebCmds As New WebCommands

    gStartNewWebSession = False
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gStartOrEndWebSession: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gStartOrEndWebSession: LoadOption RootURL Error"
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        'D.S. 09/16/14 added user name to the call below.
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=SWS-" & gGetComputerName() & "_" & sgUserName
    End If
    
    If bgUsingSockets Then
        llReturn = 1
        slResponse = WebCmds.ExecSQL("SWS-" & gGetComputerName() & "_" & sgUserName)
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
    
    ' Very the return code here. Anything but 200 is an error.
    If llReturn = 200 Then
        gStartNewWebSession = True
    Else
        gLogMsg "gStartOrEndWebSession Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", "AffWebErrorLog.Txt", False
        gMsgBox "gStartOrEndWebSession Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]"
        Exit Function
    End If
    Exit Function

ErrHandler:
    gMsg = "A general error has occured in modWebSubs-gStartNewWebSession: "
    gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.Txt", False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
End Function

Public Function gEndWebSession(LogFileName As String, Optional sLogMsg As String) As Boolean
    On Error GoTo ErrHandler
    Dim alDataArray() As String
    Dim slTemp As String
    Dim llRet As Long

    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    
    slTemp = slTemp & "_" & sgUserName
    
    gEndWebSession = True
    'D.S. 09/16/14 new call to gExecWebSQL using the computer name and user rather than just UKey.
    'Now it should be unlikely to end someone elses web session
    llRet = gExecWebSQL(alDataArray, "Delete From WebSession Where PCName = " & "'" & slTemp & "'", False)
    'Call gExecWebSQL(alDataArray, "Delete From WebSession Where UKey = 1", True)
    If sLogMsg <> "N" Then
        If llRet = 0 Then
            gLogMsg "Web Session Ended Successfully", LogFileName, False
        Else
            gLogMsg "Error: Web Session Failed to End Successfully.  gExecWebSQL returned: " & llRet, LogFileName, False
        End If
    End If
    Exit Function

ErrHandler:
    gMsg = "A general error has occured in modWebSubs-gEndWebSession: "
    gLogMsg gMsg & Err.Description & " Error #" & Err.Number, LogFileName, False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
End Function
Public Function gKillTimedOutWebSession(LogFileName As String) As Boolean

    'D.S. 09/16/14 Created function
    On Error GoTo ErrHandler
    Dim alDataArray() As String
    Dim slTemp As String
    Dim llRet As Long

    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    
    slTemp = slTemp & "_" & sgUserName
    
    gKillTimedOutWebSession = True
    Call gExecWebSQL(alDataArray, "Delete From WebSession Where UKey = 1", False)
    If llRet = 0 Then
        gLogMsg "Timed Out Web Session Ended Successfully", LogFileName, False
    Else
        gLogMsg "Error: Timed Out Web Session Failed to End Successfully.  gExecWebSQL returned: " & llRet, LogFileName, False
    End If
    Exit Function

ErrHandler:
    gMsg = "A general error has occured in modWebSubs-gKillTimedOutWebSession: "
    gLogMsg gMsg & Err.Description & " Error #" & Err.Number, LogFileName, False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
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

    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gCheckWebSession: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gCheckWebSession: LoadOption RootURL Error"
        Exit Function
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'D.S. 4/6/18 Start logging all remote updates and deletes
    If InStr(1, sSQL, "Delete") > 0 Then
        If InStr(1, sSQL, "WebSession") = 0 Then
        gLogMsg "gExecWebSQL: " & sSQL, "RemoteSQLCalls.txt", False
    End If
    End If
    If InStr(1, sSQL, "Update") > 0 Then
        If InStr(1, sSQL, "vendors") = 0 Then
        gLogMsg "gExecWebSQL: " & sSQL, "RemoteSQLCalls.txt", False
    End If
    End If
    
    slResponse = ""
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To 20
        llReturn = 1
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
            'If gExecWebSQL < 1 Then
            If gExecWebSQL = 1 Then
                ' The SQL Statement asked for two fields. DTStamp and PCName
                Exit Function
            End If
        End If
        Call gSleep(1)
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "gExecWebSQL, retries were exceeded. sSQL = " & sSQL & " llErrorCode = " & llErrorCode, "AffWebErrorLog.txt", False
    Exit Function
    
ErrHandler:
    llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-gExecWebSQL: "
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number & vbCrLf & "  " & sSQL, "AffWebErrorLog.Txt", False
    'gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    Resume Next
End Function

'Public Function gAdjustWebTimeZone(lAttCode As Long, lDelta As Long) As Boolean
'
'    On Error GoTo ERR_gAdjustWebTimeZone
'    Dim objXMLHTTP
'    Dim llReturn As Long
'    Dim slISAPIExtensionDLL As String
'    Dim slRootURL As String
'    Dim slResponse As String
'    Dim slRegSection As String
'    Dim slURLResponse As String
'    Dim ilRet As Integer
'
'    If Not gHasWebAccess() Then
'        gAdjustWebTimeZone = True
'        Exit Function
'    End If
'    gAdjustWebTimeZone = False
'    If Not gLoadOption("WebServer", "RootURL", slRootURL) Then
'        gLogMsg "Error: gAdjustWebTimeZone: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
'        gMsgBox "Error: gAdjustWebTimeZone: LoadOption RootURL Error"
'        Exit Function
'    End If
'    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
'    If gLoadOption("WebServer", "RegSection", slRegSection) Then
'        slISAPIExtensionDLL = slRootURL & "ExecExSP.dll?ChangeDateTime?PW=jfdl&RK=" & slRegSection & "&LL=0" & "&attCode=" & lAttCode & "&Delta=" & lDelta
'    End If
'
''    If bgUsingSockets Then
''        slResponse = WebCmds.
''    Else
'        Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
'        objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
'        objXMLHTTP.Send
'        llReturn = objXMLHTTP.Status
'        slResponse = objXMLHTTP.responseText
'        Set objXMLHTTP = Nothing
''    End If
'
'    ' Very the return code here. Anything but 200 is an error.
'    If llReturn <> 200 Then
'        gMsgBox "gAdjustWebTimeZone Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]"
'        gLogMsg "gAdjustWebTimeZone Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", "AffWebErrorLog.Txt", False
'        Exit Function
'    End If
'    gLogMsg slResponse, "AffWebErrorLog.Txt", False
'    gAdjustWebTimeZone = True
'    Exit Function
'
'ERR_gAdjustWebTimeZone:
'    gLogMsg "gAdjustWebTimeZone Error - " & vbCrLf & slISAPIExtensionDLL & vbCrLf & vbCrLf & "Response = [" & slResponse & "]", "AffWebErrorLog.Txt", False
'End Function

'*****************************************************************************************************
' Returns the number of rows effected.
'
'*****************************************************************************************************
Public Function gExecWebSQLWithRowsEffected(sSQL As String) As Long
    
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim ilRet As Integer
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim ilIdx As Integer
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim ilMaxRetries As Integer
    Dim ilStepNo As Integer
    Dim WebCmds As New WebCommands
    
    
ilStepNo = 1
    '10000
'    If igDemoMode Or igTestSystem Then
    If igDemoMode Then
        gExecWebSQLWithRowsEffected = 0
        Exit Function
    End If
    
    gExecWebSQLWithRowsEffected = -1    ' -1 is an error condition.

    'D.S. 02/03/09
    'First make sure that we can establish a connection
    ilMaxRetries = 4
    ilStepNo = 2
    ilRet = gHasWebAccess()
    ilStepNo = 3
    If Not ilRet Then
        '10472 code does nothing
'        For ilRetries = 0 To ilMaxRetries Step 1
'            Call Sleep(3000)  'Sleep 3 seconds
'            ilRet = gHasWebAccess()
'            If ilRet = True Then
'                'We got access
'                Exit For
'            End If
'        Next ilRetries
'        If Not ilRet And ilRetries = ilMaxRetries Then
            gLogMsg "Error: gExecWebSQLWithRowsEffected was called and gHasWebAccess returned False", "AffWebErrorLog.txt", False
            Exit Function
'        End If
    
    End If
    ilStepNo = 4
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gExecWebSQLWithRowsEffected: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: gExecWebSQLWithRowsEffected: LoadOption RootURL Error"
        Exit Function
    End If
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
   
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'D.S. 4/6/18 Start logging all remote updates and deletes
    If InStr(1, sSQL, "Delete") > 0 Then
        If InStr(1, sSQL, "WebSession") = 0 Then
            gLogMsg "gExecWebSQLWithRowsEffected: " & sSQL, "RemoteSQLCalls.txt", False
    End If
    End If
    If InStr(1, sSQL, "Update") > 0 Then
        If InStr(1, sSQL, "vendors") = 0 Then
            gLogMsg "gExecWebSQLWithRowsEffected: " & sSQL, "RemoteSQLCalls.txt", False
    End If
    End If
    
    For ilRetries = 0 To 14
        llReturn = 1
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
            'gLogMsg "gExecWebSQLWithRowsEffected is retrying", "WebSubsRetryLog.Txt", False
            Call gSleep(2)  ' Delay for 2 seconds between requests.
        End If
    Next
    
    ' We were never successful if we make it to here.
    gLogMsg "gExecWebSQLWithRowsEffected, retries were exceeded. slResponse = " & slResponse & " StepNo = " & ilStepNo, "AffWebErrorLog.txt", False
    
    Exit Function
    
ErrHandler:
    'llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-gExecWebSQLWithRowsEffected: Retries = " & ilRetries & " Step Num = " & ilStepNo & " slResponse = " & slResponse & " llReturn = " & llReturn
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.txt", False
    'gLogMsg "    SQLQuery = " & sSQL, "AffWebErrorLog.txt", False
    Resume Next
End Function

Public Function gIsWebAgreement(lAttCode As Long) As Boolean

    Dim temp_rst As ADODB.Recordset
    
    gIsWebAgreement = False
    '7701
    SQLQuery = "Select attExportType FROM Att WHERE AttCode = " & lAttCode
    Set temp_rst = gSQLSelectCall(SQLQuery)

    If temp_rst!attExportType = 1 Then
        gIsWebAgreement = True
    End If
'    SQLQuery = "Select attExportType, attExportToWeb, attWebInterface FROM Att WHERE AttCode = " & lAttCode
'    Set temp_rst = gSQLSelectCall(SQLQuery)
'
'    If temp_rst!attExportType = 1 And ((temp_rst!attExportToWeb = "Y") Or (temp_rst!attWebInterface = "C")) Then
'        gIsWebAgreement = True
'    End If

End Function
Public Function gWebDeleteSpots(attCode As Long, startDate As String, endDate As String) As Boolean

    'D.S.
    Dim temp_rst As ADODB.Recordset
    Dim slSQLQuery As String
    Dim slStr As String
    Dim llRet As Long
    Dim llPrevCount As Long
    Dim llSQLCount As Long
    
    On Error GoTo ErrHand
    
    gWebDeleteSpots = False
    
    SQLQuery = "Select Count(astCode) FROM Ast WHERE (astAtfCode = " & attCode
    SQLQuery = SQLQuery & " AND astFeedDate >= '" & Format$(startDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(endDate, sgSQLDateForm) & "')"
    Set temp_rst = gSQLSelectCall(SQLQuery)
    
    If Not temp_rst.EOF Then
        If temp_rst(0).Value > 0 Then
            llPrevCount = temp_rst(0).Value
        Else
            llPrevCount = 0
        End If
    Else
        llPrevCount = 0
    End If
    
    SQLQuery = "Delete from Spots Where attCode = " & attCode
    SQLQuery = SQLQuery & " AND FeedDate >= '" & Format$(startDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND FeedDate <= '" & Format$(endDate, sgSQLDateForm) & "'"
    llSQLCount = gExecWebSQLWithRowsEffected(SQLQuery)
    
    If llPrevCount <> llSQLCount Then
        gLogMsg "Warning:  ATT Code: " & attCode & " for the date range: " & startDate & " - " & endDate, "AffWebErrorLog.Txt", False
        gLogMsg "    gWebDeleteSpots from web. Affilite shows: " & llPrevCount & " and the web deleted: " & llSQLCount, "AffWebErrorLog.Txt", False
    End If
    
    'D.S. 01/24/20
    'Delete spots from the revisions table
    SQLQuery = "Delete from SpotRevisions Where attCode = " & attCode
    SQLQuery = SQLQuery & " AND FeedDate >= '" & Format$(startDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND FeedDate <= '" & Format$(endDate, sgSQLDateForm) & "'"
    llSQLCount = gExecWebSQLWithRowsEffected(SQLQuery)
    
    gWebDeleteSpots = True
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebDeleteSpots"
    Exit Function
End Function



Public Function gWebUpdateEmail(lCode As Long, iShttCode As Integer, iSeqNum As Integer, sEmail As String, sFirstName As String, sLastName As String, iTntCode As Integer) As Boolean

    'D.S. 1/5/11
    'Update email information on the web site if it exists.  Otherwise, insert it.
    
    Dim ilRowsEffected As Integer
    Dim ilRet As Integer
    Dim slTitle As String

    On Error GoTo ErrHand
    
    gWebUpdateEmail = False
    
    SQLQuery = "Select Count(*) from WebEmt where shttCode = " & iShttCode & " And seqNo = " & iSeqNum
    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
    
    slTitle = gGetTitleByTntCode(iTntCode)
    
    'This is the normal case where a record needs updating on the web
    If ilRowsEffected = 1 Then
        SQLQuery = "Update WebEmt Set Email = '" & gFixQuote(Trim$(sEmail)) & "',"
        SQLQuery = SQLQuery & " FirstName = '" & gFixQuote(Trim$(sFirstName)) & "',"
        SQLQuery = SQLQuery & " LastName = '" & gFixQuote(Trim$(sLastName)) & "',"
        SQLQuery = SQLQuery & " Title = '" & gFixQuote(Trim$(slTitle)) & "',"
        SQLQuery = SQLQuery & "DateModified = '" & Format(Now, "ddddd ttttt") & "'"
        SQLQuery = SQLQuery & " Where shttCode = " & iShttCode & " And seqNo = " & iSeqNum
        ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
        If ilRowsEffected = -1 Then
            gLogMsg "Update Failed: " & SQLQuery, "WebEmailInsertLog.Txt", False
        End If
        'D.S. 8/29/19 TTP 9509
        If CDbl(sgWebSiteVersion) < 7.1 Then
            gWebUpdateEmail = True
        End If

        If CDbl(sgWebSiteVersion) >= 7.1 Then
            If gWebUpdateAccessControl Then
                gWebUpdateEmail = True
            End If
        End If

    Else
        'This is the case where on the Affiliate a web email was created and sent to the web. Then they unchecked
        'the web email which caused the record to be deleted on the web.  Now they re-checked the web email box which
        'causes an update on Artt, but the record no longer exists on the web.  So, we need to insert it on the web.
        ilRet = gWebInsertEmail(lCode, iShttCode, iSeqNum, sEmail, sFirstName, sLastName, iTntCode)
        If ilRet Then
            If CDbl(sgWebSiteVersion) >= 7.1 Then
                If gWebUpdateAccessControl Then
                    gWebUpdateEmail = True
                End If
            End If
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebUpdateEmail"
    gWebUpdateEmail = False
    Exit Function
End Function


Public Function gWebInsertEmail(lCode As Long, iShttCode As Integer, iSeqNum As Integer, sEmail As String, sFirstName As String, sLastName As String, iTntCode As Integer) As Boolean

    'D.S. 1/5/11
    'Insert new email information into the web site
    
    Dim ilRowsEffected As Integer
    Dim llAttCode As Long
    Dim slTitle As String
    
    On Error GoTo ErrHand
    
    gWebInsertEmail = False
    
    'init, we really don't use this
    llAttCode = 0

    slTitle = gGetTitleByTntCode(iTntCode)

    If gIsUsingNovelty Then
        SQLQuery = "usp_AddStationUser "
        SQLQuery = SQLQuery & lCode & ", "
        SQLQuery = SQLQuery & "'" & Trim$(gGetCallLettersByShttCode(iShttCode)) & "', "
        SQLQuery = SQLQuery & iShttCode & ", "
        SQLQuery = SQLQuery & llAttCode & ", "
        SQLQuery = SQLQuery & iSeqNum & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sEmail)) & " ', "
        SQLQuery = SQLQuery & "'" & " " & "',"
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sFirstName)) & " ',"    ' Do not remove the
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sLastName)) & " ',"     ' spaces you see before
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slTitle)) & " '"        ' the comments!
    Else

        SQLQuery = "Insert Into WebEmt ( "
        SQLQuery = SQLQuery & "Code, "
        SQLQuery = SQLQuery & "CallLetters, "
        SQLQuery = SQLQuery & "ShttCode, "
        SQLQuery = SQLQuery & "AttCode, "
        SQLQuery = SQLQuery & "SeqNo, "
        SQLQuery = SQLQuery & "EMail, "
        SQLQuery = SQLQuery & "Status, "
        SQLQuery = SQLQuery & "FirstName, "
        SQLQuery = SQLQuery & "LastName, "
        SQLQuery = SQLQuery & "Title, "
        SQLQuery = SQLQuery & "DateModified "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & lCode & ", "
        SQLQuery = SQLQuery & "'" & Trim$(gGetCallLettersByShttCode(iShttCode)) & "', "
        SQLQuery = SQLQuery & iShttCode & ", "
        SQLQuery = SQLQuery & llAttCode & ", "
        SQLQuery = SQLQuery & iSeqNum & ", "
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sEmail)) & "', "
        SQLQuery = SQLQuery & "'" & " " & "',"
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sFirstName)) & "',"
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(sLastName)) & "',"
        SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(slTitle)) & "',"
        SQLQuery = SQLQuery & "'" & Format(Now, "ddddd ttttt") & "' "
        SQLQuery = SQLQuery & ") "
    End If
    
    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
    If ilRowsEffected = -1 Then
        gLogMsg "Insert Failed: " & SQLQuery, "WebEmailInsertLog.Txt", False
    Else
        gWebInsertEmail = True
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebInsertEmail"
    gWebInsertEmail = False
    Exit Function
End Function

Public Function gWebDeleteEmail(iShttCode As Integer, iSeqNum As Integer) As Boolean

    'D.S. 1/5/11
    'Delete email from web site
    
    Dim ilRowsEffected As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    gWebDeleteEmail = False
    
    SQLQuery = "Delete From webEmt Where ShttCode = " & iShttCode & " And SeqNo = " & iSeqNum
    ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
    If ilRowsEffected = -1 Then
        gLogMsg "Error: Delete: " & SQLQuery, "WebEmailInsertLog.Txt", False
    End If
    gWebDeleteEmail = True


Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebDeleteEmail"
    gWebDeleteEmail = False
    Exit Function
End Function


Public Function gWebTestEmailChange() As Boolean

    'D.S. 1/5/11
    'Check to see if the ARTT table contains any email records that need to be sent to the web
    'i.e. Insert, Update or Delete
    
    
    Dim rst_artt As ADODB.Recordset
    Dim ilRet As Integer
    Dim slvalue As String
    
    On Error GoTo ErrHand
    
    gWebTestEmailChange = False
    
    'Handle the case where they either added a new web email or updated a web email
    SQLQuery = "SELECT * FROM artt"
    SQLQuery = SQLQuery & " WHERE ("
    SQLQuery = SQLQuery & " arttWebEmail = " & "'" & "Y" & "'" & " And arttEmailToweb <> " & "'" & "S" & "'" & ")"
    Set rst_artt = gSQLSelectCall(SQLQuery)
    
    If Not rst_artt.EOF Then
        Do While Not rst_artt.EOF
            slvalue = rst_artt!arttEmailToWeb
            Select Case slvalue
                Case "I"   'Insert
                    ilRet = gWebInsertEmail(rst_artt!arttCode, rst_artt!arttShttCode, rst_artt!arttWebEMailRefID, rst_artt!arttEmail, rst_artt!arttFirstName, rst_artt!arttLastName, rst_artt!arttTntCode)
                Case "U"   'Update
                    ilRet = gWebUpdateEmail(rst_artt!arttCode, rst_artt!arttShttCode, rst_artt!arttWebEMailRefID, rst_artt!arttEmail, rst_artt!arttFirstName, rst_artt!arttLastName, rst_artt!arttTntCode)
            End Select
            If ilRet Then
                ilRet = gWebSetEmailAsSent(rst_artt!arttCode)
            End If
        rst_artt.MoveNext
        Loop
    End If
    
    'Handle the case where they changed from sending to the web to not sending to the web
    SQLQuery = "SELECT * FROM artt"
    SQLQuery = SQLQuery & " WHERE ("
    SQLQuery = SQLQuery & " arttWebEmail = " & "'" & "N" & "'" & " And arttEmailToweb = " & "'" & "D" & "'" & ")"
    Set rst_artt = gSQLSelectCall(SQLQuery)
    
    If Not rst_artt.EOF Then
        Do While Not rst_artt.EOF
           ilRet = gWebDeleteEmail(rst_artt!arttShttCode, rst_artt!arttWebEMailRefID)
            If ilRet Then
                ilRet = gWebSetEmailAsSent(rst_artt!arttCode)
            End If
        rst_artt.MoveNext
        Loop
    End If
    
    
    rst_artt.Close
    gWebTestEmailChange = True
    
Exit Function
    
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebTestEmailChange"
    gWebTestEmailChange = False
    Exit Function
End Function


Public Function gWebSetEmailAsSent(sCode As String) As Integer

    'D.S. 1/5/11
    'Set the email sent flag in the ARTT table so we don't send it again
    
    On Error GoTo ErrHand
    
    gWebSetEmailAsSent = False
    
    SQLQuery = "UPDATE artt"
    SQLQuery = SQLQuery & " SET arttEmailToWeb = " & "'" & "S" & "'"
    SQLQuery = SQLQuery & " WHERE arttCode = " & "'" & sCode & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "WebEmailInsertLog.Txt", "modWebSubs-gWebSetEmailAsSent"
        gWebSetEmailAsSent = False
        Exit Function
    End If
    
    gWebSetEmailAsSent = True
    Exit Function


ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebSetEmailAsSent"
    gWebSetEmailAsSent = False
    Exit Function
End Function
    

Public Function gTestWebVersion() As Integer
    'D.S. 08-26-08
    'D.S. 03/22/12 re-worked version code.
    '10000
'    If igDemoMode Or igTestSystem Then
    If igDemoMode Then
        gTestWebVersion = True
        Exit Function
    End If
    'NOTE: sgWebSiteExpectedByAffiliate must be in sync with the web or it's an error
    
    gTestWebVersion = 0
    
    'Check for SQL version table, if not there then try to create it and insert the current version
    If Not igDemoMode Then
        If gUsingWeb Then
            If gHasWebAccess() Then
                sgWebSiteNeedsUpdating = "False"
                sgWebSiteVersion = ""
                'D.S. 03/20/12 Rolled the version to 1.0 for NOT handling MG, Missed etc.
                'D.S. 03/20/12 Rolled the version to 2.0 for handling MG, Missed etc.
                sgWebSiteExpectedByAffiliate = "6.5"  'hanldes all MG etc. and 90 million for new web version 3.0
                If VerifyVersionTable() Then
                    If CDbl(sgWebSiteVersion) < CDbl(sgWebSiteExpectedByAffiliate) Then
                        sgWebSiteNeedsUpdating = "True"
                        gMsgBox "Call Counterpoint. An Incorrect WebSite Version was Found." & sgCRLF & sgCRLF & "               Web = " & sgWebSiteVersion & " Local = " & sgWebSiteExpectedByAffiliate & sgCRLF & sgCRLF & "                You May Continue but, " & sgCRLF & sgCRLF & "        No Imports or Exports will be Allowed.", vbCritical
                        gLogMsg "Call Counterpoint. An Incorrect WebSite Version was Found." & " Web = " & sgWebSiteVersion & " Local = " & sgWebSiteExpectedByAffiliate & " You May Continue but, No Imports or Exports will be Allowed.", "AffErrorLog.Txt", False
                        Exit Function
                    End If
                Else
                    gMsgBox "Call Counterpoint. No WebSite Version Number was Found", vbCritical
                    gLogMsg "Call Counterpoint. No WebSite Version Number was Found", "AffErrorLog.Txt", False
                    sgWebSiteNeedsUpdating = "True"
                    Exit Function
                End If
            Else
                'If the client is using the web and the ini file is wrong then it's an error
                gMsgBox "Error: Call Counterpoint. Could Not Connect to the Web.", vbCritical
                gLogMsg "Error: Call Counterpoint. Could Not Connect to the Web.", "AffErrorLog.Txt", False
                sgWebSiteNeedsUpdating = "True"
                Exit Function
            End If
        End If
    End If
    gTestWebVersion = -1
    
End Function

Public Function gInitGlobals() As Integer
    Dim slTemp As String
    
    Call gLoadOption(sgWebServerSection, "UsingSockets", slTemp)
    If UCase(slTemp) = "Y" Then
        bgUsingSockets = True
    Else
        bgUsingSockets = False
    End If
    gUsingUnivision = False
    gUsingWeb = False
    'sgExportISCI = "B"
    sgExportISCI = ""
    sgShowByVehType = "N"
    sgRCSExportCart4 = "Y"
    sgRCSExportCart5 = "N"
    sgUsingStationID = "N"
    sgMissedMGBypass = "N"
    sgUsingServiceAgreement = "N"
    igLastPoolAdfCode = -1
    '8/1/14: Not used with v7.0
    'sgMarketronCompliant = "P"
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        '8000, re-allow
        gUsingUnivision = rst!siteMarketron
        'gUsingUnivision = False
        gUsingWeb = rst!siteWeb
        'sgExportISCI = rst!siteISCIExport
        sgShowByVehType = rst!siteShowVehType
        sgRCSExportCart4 = rst!siteExportCart4
        sgRCSExportCart5 = rst!siteExportCart5
        sgUsingStationID = rst!siteUsingStationID
        sgMissedMGBypass = rst!siteMissedMGBypass
        sgUsingServiceAgreement = rst!siteUsingServAgree
        '8/1/14: Not used with v7.0
        'sgMarketronCompliant = rst!siteCompliantBy
        'If Trim$(sgMarketronCompliant) = "" Then
        '    sgMarketronCompliant = "P"
        'End If
        If gUsingWeb Then
            gInitGlobals = 0
            'While Not gVerifyWebIniSettings()
               ' frmWebIniOptions.Show vbModal
               ' If Not igWebIniOptionsOK Then
               '     Unload frmLogin
               '     Exit Function
               ' End If
            'Wend
            '10000
           ' If Not igTestSystem Then
                If Not igDemoMode Then
                    If Not gTestAccessToWebServer() Then
                        gMsgBox "WARNING!" & vbCrLf & vbCrLf & _
                               "Web Server Access Error: The Affiliate System does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
                        "No data will be exported to the web site." & vbCrLf & _
                        "No data will be imported from the web site." & vbCrLf & _
                        "Sign off system immediately and contact system administrator.", vbExclamation
                    Else
                        Exit Function
                    End If
                End If
           ' End If
        End If
    End If
    gInitGlobals = 1
End Function

Public Function gWebUpdatePW(iShttCode As Integer) As Boolean

    'D.S. 04/11/12
    
    Dim rst_Shtt As ADODB.Recordset
    Dim ilRowsEffected As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    gWebUpdatePW = False
    
    SQLQuery = "Select shttCallLetters, shttWebPW From SHTT Where shttCode = " & iShttCode
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        SQLQuery = "Update Header Set attStationPW = '" & Trim(rst_Shtt!shttWebPW) & "'"
        SQLQuery = SQLQuery & " Where stationName = " & "'" & Trim$(rst_Shtt!shttCallLetters) & "'"
        For ilLoop = 0 To 5 Step 1
            ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
            If ilRowsEffected <> -1 Then
                Exit For
            End If
            Sleep (1000)
        Next ilLoop
        If ilRowsEffected = -1 Then
            gLogMsg "Update Failed: modWebSubs - gWebUpdatePW " & SQLQuery, "AffWebErrorLog.Txt", False
        Else
            gWebUpdatePW = True
            If CDbl(sgWebSiteVersion) >= 7.1 Then
                ilRet = gWebUpdateAccessControl
            End If
        End If
    End If
    rst_Shtt.Close
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gWebUpdatePW"
    Exit Function
End Function



Public Function gGetLogAttID(ilVefCode As Integer, ilShttCode As Integer, llAttCode As Long) As Long
    Dim slMaxQuery As String
    Dim llWebLogAttID As Long
    Dim llLatCode As Long
    Dim ilRet As Integer
    Dim llVef As Long
    Dim ilVff As Integer
    Dim ilLogVefCode As Integer
    Dim rst_Lat As ADODB.Recordset
    
    On Error GoTo ErrHand
    gGetLogAttID = llAttCode
    llVef = gBinarySearchVef(CLng(ilVefCode))
    If llVef = -1 Then
        Exit Function
    End If
    If tgVehicleInfo(llVef).iVefCode > 0 Then
        'Determine if vehicle is to be merged into Log vehicle on the web
        For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
            If tgVehicleInfo(llVef).iCode = tgVffInfo(ilVff).iVefCode Then
                If tgVffInfo(ilVff).sMergeWeb <> "S" Then
                    ilLogVefCode = tgVehicleInfo(llVef).iVefCode
                Else
                    Exit Function
                End If
                Exit For
            End If
        Next ilVff
    ElseIf tgVehicleInfo(llVef).sVehType = "L" Then
        ilLogVefCode = ilVefCode
    Else
        Exit Function
    End If
    SQLQuery = "SELECT latWebLogAttID "
    SQLQuery = SQLQuery + " FROM lat"
    SQLQuery = SQLQuery + " WHERE (latLogVefCode = " & ilLogVefCode
    SQLQuery = SQLQuery + " AND latShttCode = " & ilShttCode & ")"
    Set rst_Lat = gSQLSelectCall(SQLQuery)
    If Not rst_Lat.EOF Then
        gGetLogAttID = -rst_Lat!latWebLogAttID
        Exit Function
    Else
        slMaxQuery = "SELECT MAX(latWebLogAttID) from lat"
        Do
            Set rst = gSQLSelectCall(slMaxQuery)
            'Dan M 9/14/10 take care of Null
            If IsNull(rst(0).Value) Then
                llWebLogAttID = 1
            Else
                If Not rst.EOF Then
                    llWebLogAttID = rst(0).Value + 1
                Else
                    llWebLogAttID = 1
                End If
            End If
            ilRet = 0
            SQLQuery = "Insert Into lat ( "
            SQLQuery = SQLQuery & "latCode, "
            SQLQuery = SQLQuery & "latLogVefCode, "
            SQLQuery = SQLQuery & "latShttCode, "
            SQLQuery = SQLQuery & "latWebLogAttID, "
            SQLQuery = SQLQuery & "latUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & "Replace" & ", "
            SQLQuery = SQLQuery & ilLogVefCode & ", "
            SQLQuery = SQLQuery & ilShttCode & ", "
            SQLQuery = SQLQuery & llWebLogAttID & ", "
            SQLQuery = SQLQuery & "'" & gFixQuote("") & "' "
            SQLQuery = SQLQuery & ") "
            llLatCode = gInsertAndReturnCode(SQLQuery, "lat", "latCode", "Replace")
            If llLatCode <= 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                Screen.MousePointer = vbDefault
                If Not gHandleError4994("WebExportLog.Txt", "modWebSubs-gGetLogAttID") Then
                    gGetLogAttID = False
                    On Error Resume Next
                    rst_Lat.Close
                    Exit Function
                End If
                ilRet = 1
            End If
            gGetLogAttID = -llWebLogAttID
        Loop While ilRet <> 0
    End If
    On Error Resume Next
    rst_Lat.Close
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "frmWebSubs-gGetLogAtt"
End Function
Public Function gUrlEncoding(sStr As String) As String

    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    
    Dim slTemp As String
    
    gUrlEncoding = ""
    
    slTemp = sStr
    'slTemp = Replace$(slTemp, "$", "%24", 1, Len(slTemp), vbTextCompare)
    slTemp = Replace$(slTemp, "&", "%26", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "+", "%2B", 1, Len(slTemp), vbTextCompare)
'    'slTemp = Replace$(slTemp, ",", "%2C", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "/", "%2F", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, ":", "%3A", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, ";", "%3B", 1, Len(slTemp), vbTextCompare)
'    'slTemp = Replace$(slTemp, "=", "%3D", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "?", "%3F", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "@", "%40", 1, Len(slTemp), vbTextCompare)
'    'slTemp = Replace$(slTemp, " ", "%20", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, """", "%22", 1, Len(slTemp), vbTextCompare)
'    'slTemp = Replace$(slTemp, "<", "%3C", 1, Len(slTemp), vbTextCompare)
'    'slTemp = Replace$(slTemp, ">", "%3E", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "#", "%23", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "%", "%25", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "{", "%7B", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "}", "%7D", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "|", "%7C", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "\", "%5C", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "^", "%5E", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "~", "%7E", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "[", "%5B", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "]", "%5D", 1, Len(slTemp), vbTextCompare)
'    slTemp = Replace$(slTemp, "`", "%60", 1, Len(slTemp), vbTextCompare)
    gUrlEncoding = slTemp
End Function


Public Function gTestFTPFileExists(sFileName As String) As Integer

    Dim ilRet As Integer
    Dim slTemp As String
    Dim llCount As Long

    gTestFTPFileExists = 0
    
    'degug
    'sFileName = "zz.txt" '& sFileName
    
    tgCsiFtpFileListing.nTotalFiles = 0
    
    slTemp = Trim$(tgCsiFtpFileListing.sPathFileMask) & "\" & sFileName
    tgCsiFtpFileListing.sPathFileMask = slTemp
    tgCsiFtpFileListing.sSavePathFileName = tgCsiFtpFileListing.sLogPathName
    tgCsiFtpFileListing.sLogPathName = Trim$(tgCsiFtpFileListing.sLogPathName) & "\" & "Messages\FTPLog.txt"
    tgCsiFtpFileListing.sSavePathFileName = tgCsiFtpFileListing.sLogPathName
    ilRet = csiFTPGetFileListing(tgCsiFtpFileListing)
    llCount = tgCsiFtpFileListing.nTotalFiles
    If ilRet = False Then
        gLogMsg "csiFTPGetFileListing failed.", "AffWebErrorLog.Txt", False
        MsgBox "csiFTPGetFileListing failed. Please notify Counterpoint."
        gTestFTPFileExists = -1
        Exit Function
    End If
        
    'Debug
    'MsgBox tgCsiFtpFileListing.nTotalFiles & " Files were found and written to " & tgCsiFtpFileListing.sSavePathFileName
    
    gTestFTPFileExists = llCount
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebSubs-gTestFTPFileExists"
    Exit Function
End Function

Sub mUpdateShttTables(slCallLetters As String, slWebPW As String)
    '11/26/17
    Dim ilIndex As Integer
    Dim blRepopRequired As Boolean
    
    blRepopRequired = False
    ilIndex = gBinarySearchStation(slCallLetters)
    If ilIndex <> -1 Then
        tgStationInfo(ilIndex).sWebPW = Trim(slWebPW)
        ilIndex = gBinarySearchStationInfoByCode(tgStationInfo(ilIndex).iCode)
        If ilIndex <> -1 Then
            tgStationInfoByCode(ilIndex).sWebPW = Trim(slWebPW)
        Else
            blRepopRequired = True
        End If
    Else
        blRepopRequired = True
    End If

    gFileChgdUpdate "shtt.mkd", blRepopRequired
End Sub

Public Function GetServersDateTime() As String
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slCommand As String
    Dim slRegSection As String
    Dim alRecordsArray() As String
    Dim aDataArray() As String
    Dim WebCmds As New WebCommands

    GetServersDateTime = ""

    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gMsgBox "FAIL: gCheckWebSession: LoadOption RootURL Failed"
        Exit Function
    End If
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    slCommand = "Select GetDate() as ServerDate"
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & slCommand
    End If
    llReturn = 1
    If bgUsingSockets Then
        slResponse = WebCmds.ExecSQL(slCommand)
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
    If llReturn <> 200 Then
        Exit Function
    End If
    alRecordsArray = Split(slResponse, vbCrLf)
    If Not IsArray(alRecordsArray) Then
        Exit Function
    End If
    If UBound(alRecordsArray) < 1 Then
        Exit Function
    End If

    GetServersDateTime = gGetDataNoQuotes(alRecordsArray(1))
    Exit Function
    
ErrHandler:
    gMsg = "A general error has occurred in modGenSubs-GetServersDateTime: "
    gLogMsg gMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "WebImportLog.Txt", False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
End Function
Public Function gIsTestWebServer() As Boolean
    Dim blRet As Boolean
    blRet = False
    If InStr(sgWebServerSection, "Test") > 0 Then
        blRet = True
    End If
    gIsTestWebServer = blRet
End Function



