Attribute VB_Name = "WEBSPOTS"


'*********************************************************************************
'
' Copyright 2004 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: WebSpots.BAS
'
' Description:
'   This file contains routines and definitions to support the Station Traffic
'   system's interaction with the network web site
'
'*********************************************************************************

Option Compare Text
Option Explicit

Private tmFnf As FNF
Private tmFsf As FSF
Private tmLcfArray() As Ext_LCF
Private tmSdfArray() As SDF
Private hmFsf As Integer
Private imFsfRecLen As Integer
Private smURLToPostTo As String

Private tmCif As CIF
Private tmCifSrchKey As LONGKEY0
Private tmCpf As CPF
Private tmCpfSrchKey As LONGKEY0
Private hmAdf As Integer
Private tmAdfSrchKey As INTKEY0
Private hmMcf As Integer
Private tmMcf As MCF
Private tmMcfSrchKey As INTKEY0
Private hmPrf As Integer
Private tmPrfSrchKey As INTKEY0
Private hmCif As Integer
Private hmCpf As Integer
Private imPurgeInvMcf As Integer
Private tmPurgeInv() As SORTCODE

Type WEBSPOTS
    lAstCode As Long
    lAttCode As Long
    sStationName As String * 40
    sVehicleName As String * 40
    sAdvt As String * 30
    sProd As String * 35
    sPledgeStartDate As String * 10
    sPledgeEndDate As String * 10
    sPledgeStartTime As String * 11
    sPledgeEndTime As String * 11
    iSpotLen As Integer
    sCart As String * 7
    sCreativeTitle As String * 30
    sISCI As String * 20
End Type
Public tgWebSpots() As WEBSPOTS

Type WEBFEED
    iCode As Integer
    sName As String * 40
    sIPAddress As String * 70
    sPassword As String * 10
    sDaysToPoll(0 To 6) As String * 1
    lInterval As Long
    lStartPoll As Long
    lEndPoll As Long
    lNextPoll As Long
    iMcfCode As Integer
End Type
Public tgWebFeed() As WEBFEED

Type WEBUSER
    sPrimaryUser As String * 40
    iPrimaryCode As Integer
    sSecondaryUser As String * 40
    iSecondaryCode As Integer
End Type
Public tgWebUser As WEBUSER

'Lcf record layout
Type Ext_LCF
    iVefCode As Integer         'Vehicle Code (combos not allowed)
    iLogDate(0 To 1) As Integer 'Log Date
                                'Date Byte 0:Day, 1:Month, followed by 2 byte year
                                'TFN: iLogDate(0) = 1 (mon), 2 (tue),... 7 (sun) and iLogDate(1) = 0
    iSeqNo As Integer           'Sequency number (used to order log dates)
    sType As String * 1         'O=On Air; A=Altered (partial day defined)
    sStatus As String * 1       'C=Current; P=Pending
    sTiming As String * 1       'Log timing indicator (N=Not started; I=Timing Incomplete; C=Timing Completed)
    sAffPost As String * 1      'Affidavit posting (N=No posting done yet; I=Posting Incomplete; C=Posting completed)
    iLastTime(0 To 1) As Integer 'Last Time Timed or posted (signoff time if completed)
                                '(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    'lLvfCode(1 To 50) As Long  'Log version library code
    lLvfCode(0 To 49) As Long  'Log version library code
    'iTime(0 To 1, 1 To 50) As Integer 'Log library start time
    iTime(0 To 1, 0 To 49) As Integer 'Log library start time
                                '(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iUrfCode As Integer         'Last user who modified Log library
    lRecPos As Long             'So we can use a btrGetDirect to update this record later on.
End Type

Public sgFnfFileDate As String
Public Const INET_BUF_SIZE = 8196
Public hmWebFile As Integer
Public iDebugIdx As Integer

'***************************************************************************************
'*
'* Procedure Name: gTestAccessToWebServer
'*
'* Created: 8/22/03 - J. Dutschke
'*
'* Modified:          By:
'*
'* Comments: This function tests to see whether this PC has access to the web server.
'*
'***************************************************************************************
Public Function gTestAccessToWebServer() As Boolean
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slURLResponse                                                                         *
'******************************************************************************************

    On Error GoTo ERR_gTestAccessToWebServer
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slWebPage As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim slWebLogLevel As String
    Dim WebCmds As New WebCommands

    gTestAccessToWebServer = False
    If igTestSystem Then
        Exit Function
    End If
    Call gLoadOption("WebServer", "WebLogLevel", slWebLogLevel)
    If slWebLogLevel = "" Then
        slWebLogLevel = "0"
    End If
    If Not gLoadOption("WebServer", "RootURL", slRootURL) Then
        Exit Function
    End If
    slRootURL = gSetPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If Not gLoadOption("WebServer", "RegSection", slRegSection) Then
        Exit Function
    End If
    ' Make a request to view the headers. If this operation succeeds, then all is ok.
    ' NOTE: The VH.ASP page will not return any data from the header table because the password here
    ' is not correct due to the date and time being added. But this is ok because it will return the
    ' text that the password is invalid and access to the web page itself returns success.
    slWebPage = slRootURL & "VH.asp?PW=jfdl" & Now()
    
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
    objXMLHTTP.Close
    Exit Function

ERR_gTestAccessToWebServer:
    ' Exit the function if any errors occur
End Function

'***************************************************************************************
'*
'* Procedure Name: gWebExecSql
'*
'* Created: 8/23/03      By: D. Smith
'*
'* Modified: 08/09/04    By: J. Dutschke
'*
'* Comments:
'*
'***************************************************************************************
Public Function gWebExecSql(mSqlStr As String, sURLToUse As String, sFileName As String) As Boolean
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llTemp                                                                                *
'******************************************************************************************


    On Error GoTo ERR_gWebExecSql
    Dim sURL As String
    Dim slRootURL As String
    Dim slRegSection As String
    Dim slBuffer As String * INET_BUF_SIZE
    Dim slText As String
    Dim lhInternet As Long
    Dim lhINetSession As Long
    Dim llBytesRead As Long
    Dim FileHndl As Integer
    Dim ilRet As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slFileName As String
    Dim slDateTime As String

    iDebugIdx = iDebugIdx + 1

    gWebExecSql = False
    If Len(sURLToUse) < 1 Then
        If Not gLoadOption("WebServer", "RootURL", slRootURL) Then
            'MsgBox "FAIL: gWebExecSql: LoadOption RootURL Failed"
            gLogMsg "FAIL: gWebExecSql: LoadOption RootURL Failed", "FeedImport.txt", False
            Exit Function
        End If
        slRootURL = gSetPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
        ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
        ' registry to gather additional information. This is necessary to run multiple databases on the
        ' same IIS platform. The password is hardcoded and never changes.
        If gLoadOption("WebServer", "RegSection", slRegSection) Then
            sURL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & mSqlStr
        End If
    Else
        sURL = sURLToUse
    End If

    lhINetSession = InternetOpen("CSI", 0, vbNullString, vbNullString, 0)
    If lhINetSession < 1 Then
        Exit Function
    End If

    llStartTime = Timer
    lhInternet = InternetOpenUrl(lhINetSession, sURL, "", -1, INTERNET_FLAG_RAW_DATA, 0)
    llEndTime = Timer
    If lhInternet < 1 Then
        ilRet = InternetCloseHandle(lhINetSession)
        Exit Function
    End If

    ' The first thing the ExecuteSQL.DLL returns is Counterpoint Software. This is done to
    ' qualify the data.
    ilRet = InternetReadFile(lhInternet, slBuffer, 21, llBytesRead)
    If llBytesRead <> 21 Then
        ilRet = InternetCloseHandle(lhINetSession)
        ilRet = InternetCloseHandle(lhInternet)
        Exit Function
    End If

    If Left(slBuffer, 21) <> "Counterpoint Software" Then
        gWebExecSql = "Unknown data was returned --> [" & slBuffer & "]"
        ilRet = InternetCloseHandle(lhINetSession)
        ilRet = InternetCloseHandle(lhInternet)
        Exit Function
    End If

    ilRet = InternetReadFile(lhInternet, slBuffer, INET_BUF_SIZE, llBytesRead)
    If llBytesRead < 1 Then
        ilRet = InternetCloseHandle(lhINetSession)
        ilRet = InternetCloseHandle(lhInternet)
        Exit Function
    End If

    'On Error GoTo FileErr
    'FileHndl = FreeFile
    slFileName = sFileName & ".txt"
    'ilRet = 0
    'slDateTime = FileDateTime(slFileName)
    ilRet = gFileExist(slFileName)
    If ilRet = 0 Then
        Kill slFileName
    End If

    'Open slFileName For Binary Access Write As #FileHndl
    ilRet = gFileOpen(slFileName, "Binary Access Write", FileHndl)

    Do
        If llBytesRead < INET_BUF_SIZE Then
            ' This is the last block of data. Make sure we write only what we received.
            slText = Left(slBuffer, llBytesRead)
            Put #FileHndl, , slText
            Exit Do
        End If
        Put #FileHndl, , slBuffer

        ilRet = InternetReadFile(lhInternet, slBuffer, INET_BUF_SIZE, llBytesRead)
    Loop While (llBytesRead > 0)

    Close #FileHndl
    ilRet = InternetCloseHandle(lhINetSession)
    ilRet = InternetCloseHandle(lhInternet)

    gWebExecSql = True
    Exit Function

'FileErr:
'    ilRet = 1
'    Resume Next

ERR_gWebExecSql:
    ' Exit the function if any errors occur
End Function

'***************************************************************************************
'*
'* Procedure Name: gBuildFeedArray
'*
'* Created: August, 2004    By: D. Smith
'*
'* Modified:                By:
'*
'* Comments: Build array of all the different feed names and their assoc. data
'*
'***************************************************************************************

Public Function gBuildFeedArray() As Integer

    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim ilRecLen As Integer
    Dim hlFnf As Integer

    On Error GoTo ErrHand

    If tgSpf.sSystemType = "R" Then
        ReDim tgWebFeed(0 To 0) As WEBFEED
        hlFnf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hlFnf, "", sgDBPath & "Fnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        ilRecLen = Len(tmFnf)
        ilIdx = 0
        ilRet = btrGetFirst(hlFnf, tmFnf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            'If there's no FTP address then don't add it to the array
            If Trim$(tmFnf.sFTP) <> "" Then
                tgWebFeed(ilIdx).iCode = tmFnf.iCode
                tgWebFeed(ilIdx).sName = Trim$(tmFnf.sName)
                tgWebFeed(ilIdx).sPassword = Trim$(tmFnf.sPW)
                tgWebFeed(ilIdx).sIPAddress = Trim$(tmFnf.sFTP)
                tgWebFeed(ilIdx).lInterval = tmFnf.lChkInterval
                gUnpackTimeLong tmFnf.iChkStartHr(0), tmFnf.iChkStartHr(1), True, tgWebFeed(ilIdx).lStartPoll
                gUnpackTimeLong tmFnf.iChkEndHr(0), tmFnf.iChkEndHr(1), True, tgWebFeed(ilIdx).lEndPoll
                tgWebFeed(ilIdx).iMcfCode = tmFnf.iMcfCode
                'Poll Monday - Friday
                If StrComp(tmFnf.sChkDays, "YYYYYNN", vbTextCompare) = 0 Then
                    tgWebFeed(ilIdx).sDaysToPoll(0) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(1) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(2) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(3) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(4) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(5) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(6) = "N"
                'Poll Monday - Saturday
                ElseIf StrComp(tmFnf.sChkDays, "YYYYYYN", vbTextCompare) = 0 Then
                    tgWebFeed(ilIdx).sDaysToPoll(0) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(1) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(2) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(3) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(4) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(5) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(6) = "N"
                'Poll Monday - Sunday
                ElseIf StrComp(tmFnf.sChkDays, "YYYYYYY", vbTextCompare) = 0 Then
                    tgWebFeed(ilIdx).sDaysToPoll(0) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(1) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(2) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(3) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(4) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(5) = "Y"
                    tgWebFeed(ilIdx).sDaysToPoll(6) = "Y"
                'Do Not Poll
                Else
                    tgWebFeed(ilIdx).sDaysToPoll(0) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(1) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(2) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(3) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(4) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(5) = "N"
                    tgWebFeed(ilIdx).sDaysToPoll(6) = "N"
                End If
                tgWebFeed(ilIdx).lNextPoll = gTimeToLong(Format(gNow(), "hh:mm:ss"), True)
                ilIdx = ilIdx + 1
                ReDim Preserve tgWebFeed(0 To ilIdx) As WEBFEED
                ilRet = btrGetNext(hlFnf, tmFnf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            End If
        Loop
        ilRet = btrClose(hlFnf)
        btrDestroy hlFnf
    End If
    gBuildFeedArray = True
    Exit Function

ErrHand:
    ilRet = btrClose(hlFnf)
    btrDestroy hlFnf
    gDbg_HandleError "WebSpots: gBuildFeedArray"
    gBuildFeedArray = False
End Function

'***************************************************************************************
'*
'* Procedure Name: gBuildWebUser
'*
'* Created: August, 2004    By: D. Smith
'*
'* Modified:                By:
'*
'* Comments: Get the URF codes of the Primary and Secondary Web users
'*
'***************************************************************************************
Public Function gBuildWebUser() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIdx                                                                                 *
'******************************************************************************************



    On Error GoTo ErrHand


    If tgSpf.sSystemType = "R" Then
        tgWebUser.iPrimaryCode = tgSpf.iPriUrfCode
        tgWebUser.iSecondaryCode = tgSpf.iSecUrfCode
    End If

    gBuildWebUser = True
    Exit Function

ErrHand:
    gDbg_HandleError "WebSpots: gBuildWebUser"
    gBuildWebUser = False

End Function

'***************************************************************************************
'*
'* Function Name: gPollWebForSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: Does all of the polling of the web site
'*
'***************************************************************************************
Public Function gPollWebForSpots()

    Dim ilLoop As Integer
    Dim ilVehLoop As Integer
    Dim ilRet As Integer
    Dim slDay As String
    Dim llCurTime As Long
    Dim slTempStr As String

    If tgWebUser.iPrimaryCode = tgUrf(0).iCode Or tgWebUser.iSecondaryCode = tgUrf(0).iCode Then
        'Test to see if the FNF file has changed since we last built the Feed Array
        slTempStr = gFileDateTime(sgDBPath & "FNF.btr")
        If (gDateValue(slTempStr) <> gDateValue(sgFnfFileDate)) Then
            ilRet = gBuildFeedArray
            ilRet = gBuildWebUser
        End If

        If UBound(tgWebFeed) = 0 Then
            Exit Function
        End If

        ilRet = gObtainVef()
        ilRet = mOpenFiles()
        llCurTime = gTimeToLong(Format(gNow(), "hh:mm:ss"), True)
        slDay = gWeekDayStr(Format(gNow(), "mm-dd-yyyy"))

        For ilLoop = 0 To UBound(tgWebFeed) - 1 Step 1
            gLogMsg "Feed Name = " & Trim$(tgWebFeed(ilLoop).sName), "Debug.txt", False
            If tgWebFeed(ilLoop).sDaysToPoll(slDay) = "Y" Then
                If llCurTime >= tgWebFeed(ilLoop).lStartPoll And llCurTime <= tgWebFeed(ilLoop).lEndPoll Then
                    If llCurTime >= tgWebFeed(ilLoop).lNextPoll Then
                        gLogMsg "Passed the time interval test to poll", "Debug.txt", False
                        'For ilVehLoop = 1 To UBound(tgMVef) - 1 Step 1
                        For ilVehLoop = 0 To UBound(tgMVef) - 1 Step 1
                            gLogMsg "Vehicle = " & Trim$(tgMVef(ilVehLoop).sName), "Debug.txt", False
                            If tgMVef(ilVehLoop).sType = "C" Or tgMVef(ilVehLoop).sType = "S" Then
                                DoEvents

                                ilRet = mWebGetDeleteSpots(ilLoop, ilVehLoop)
                                ilRet = mReadWebSpots(True, "WebDeleted.txt")
                                ilRet = mBuildFsfFile(True, ilLoop)
                                ilRet = mWebRemDeletedSpots(ilLoop, ilVehLoop)

                                ilRet = mWebGetSpots(ilLoop, ilVehLoop)
                                ilRet = mReadWebSpots(False, "WebNewChanged.txt")
                                ilRet = mBuildFsfFile(False, ilLoop)
                                ilRet = mWebUpdateSpots(ilLoop, ilVehLoop)
                            End If
                        Next ilVehLoop
                        tgWebFeed(ilLoop).lNextPoll = tgWebFeed(ilLoop).lNextPoll + (tgWebFeed(ilLoop).lInterval * 60)
                        If tgWebFeed(ilLoop).lNextPoll > 86400 Then
                            tgWebFeed(ilLoop).lNextPoll = llCurTime
                        End If
                    End If
                End If
            End If
        Next ilLoop
        mCloseFiles
        'Schedule the spots
        gFeedSchSpots (False)
        'debug
        'ilRet = mWebDebugReset
        ilRet = gExportWebSpots()
    End If
    gEndWebPoll tgWebUser.iPrimaryCode
    gLogMsg "***** Exiting gPollWebFor Spots!!!", "Debug.txt", False
    gLogMsg "", "Debug.txt", False



End Function

'***************************************************************************************
'*
'* Function Name: gStartWebPoll
'*
'* Created: August, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************

Public Function gStartWebPoll(urfCode As Integer) As Boolean

    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlRlf As RLF
    Dim tlRlfKey As RLFKEY1
    Dim llDate As Long
    Dim llTime As Long
    Dim llCurTime As Long
    Dim llNowDate As Long
    Dim llRecDateTime As Long
    Dim llCurDateTime As Long

    gStartWebPoll = False
    If Not gTestAccessToWebServer() Then
        Exit Function
    End If

    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "RLF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_SHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "btrOpen RLF.BTR Failed" & vbCrLf & vbCrLf & "Error code = " & Str(ilRet)
        gLogMsg "btrOpen RLF.BTR Failed, Error code = " & str(ilRet), "FeedImport.txt", False
        Exit Function
    End If

    ' Prepare to insert a new record.
    tlRlf.sType = "P"
    tlRlf.iUrfCode = urfCode
    tlRlf.lRecCode = urfCode

    gPackDate Format$(gNow(), "m/d/yy"), tlRlf.iEnteredDate(0), tlRlf.iEnteredDate(1)
    gPackTime Format(gNow(), "hh:mm:ss"), tlRlf.iEnteredTime(0), tlRlf.iEnteredTime(1)
    tlRlf.sSubType = 0
    ilRecLen = btrRecordLength(hlFile)
    ilRet = btrInsert(hlFile, tlRlf, ilRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet <> BTRV_ERR_DUPLICATE_KEY Then
            ' Any error other than a duplicate error is reported back to the user right now.
            'MsgBox "btrInsert RLF.BTR Failed" & vbCrLf & vbCrLf & "Error code = " & Str(ilRet)
            gLogMsg "btrInsert RLF.BTR Failed, Error code = " & str(ilRet), "FeedImport.txt", False
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            Exit Function
        End If
        ' This user already has an entry in the table.
        tlRlfKey.sType = "P"
        tlRlfKey.lRecCode = urfCode
        ilRet = btrGetEqual(hlFile, tlRlf, ilRecLen, tlRlfKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            ' Couldn't insert and now we can't find it! Maybe it got deleted since we tried?
            ' Either case, I'm going to let the caller deal with this.
            ilRet = btrClose(hlFile)
            btrDestroy hlFile
            Exit Function
        End If

        ' Get the date and time this record was inserted.
        gUnpackDateLong tlRlf.iEnteredDate(0), tlRlf.iEnteredDate(1), llDate
        gUnpackTimeLong tlRlf.iEnteredTime(0), tlRlf.iEnteredTime(1), False, llTime
        llRecDateTime = llDate + llTime

        ' Get the current date and time.
        llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
        llCurTime = gTimeToLong(Format(gNow(), "hh:mm:ss"), True)
        llCurDateTime = llNowDate + llCurTime

        If (llCurDateTime - llRecDateTime) > 3600 Then
            ' The record is older than 1 hour. This most likely indicates a record that was inserted
            ' and then never deleted. Delete the record and return false. The caller will have to
            ' initiate the call again.
            ilRet = btrDelete(hlFile)
        End If

        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If

    gStartWebPoll = True
    ilRet = btrClose(hlFile)
    btrDestroy hlFile

End Function

'***************************************************************************************
'*
'* Function Name: gEndWebPoll
'*
'* Created: August, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************

Public Sub gEndWebPoll(urfCode As Integer)

    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlRlf As RLF
    Dim tlRlfKey As RLFKEY1

    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "RLF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_SHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "btrOpen RLF.BTR Failed" & vbCrLf & vbCrLf & "Error code = " & Str(ilRet)
        gLogMsg "btrOpen RLF.BTR Failed, Error code = " & str(ilRet), "FeedImport.txt", False
        Exit Sub
    End If

    tlRlf.iUrfCode = urfCode
    ilRecLen = btrRecordLength(hlFile)
    tlRlfKey.sType = "P"
    tlRlfKey.lRecCode = urfCode
    ilRet = btrGetEqual(hlFile, tlRlf, ilRecLen, tlRlfKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = btrDelete(hlFile)
        If ilRet <> BTRV_ERR_NONE Then
            'MsgBox "btrDelete RLF.BTR Failed" & vbCrLf & vbCrLf & "Error code = " & Str(ilRet)
            gLogMsg "btrDelete RLF.BTR Failed, Error code = " & str(ilRet), "FeedImport.txt", False
        End If
    End If

    ilRet = btrClose(hlFile)
    btrDestroy hlFile

End Sub

'***************************************************************************************
'*
'* Function Name: mWebGetSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: 'Receive the new and changed spots from the web
'*
'***************************************************************************************
Private Function mWebGetSpots(iFeedLoop As Integer, iVehLoop As Integer) As Integer

    Dim ilRet As Integer

    SQLQuery = "Select astCode, attCode, StationName, VehicleName, Advt, Prod, PledgeStartDate, PledgeEndDate, PledgeStartTime , PledgeEndTime, SpotLen, Cart, CreativeTitle, ISCI"
    SQLQuery = SQLQuery + " From VW_RadioSpots "
    SQLQuery = SQLQuery + " Where VehicleName = '" & Trim$(tgWebFeed(iFeedLoop).sName) & "'"
    SQLQuery = SQLQuery + " And StationName = '" & Trim$(tgMVef(iVehLoop).sName) & "'"
    SQLQuery = SQLQuery + " And Password = '" & Trim$(tgWebFeed(iFeedLoop).sPassword) & "'"
    SQLQuery = SQLQuery + " Order By Advt"
    ilRet = gWebExecSql(SQLQuery, "", sgImportPath & "WebNewChanged")

    mWebGetSpots = ilRet

End Function

'***************************************************************************************
'*
'* Function Name: mWebUpdateSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: 'Update the web so we don't get these spots again
'*
'***************************************************************************************
Private Function mWebUpdateSpots(iFeedLoop As Integer, iVehLoop As Integer) As Integer

    Dim ilRet As Integer

    SQLQuery = "sp_UpdateSentDate "
    SQLQuery = SQLQuery + "'" & Trim$(tgMVef(iVehLoop).sName) & "',"
    SQLQuery = SQLQuery + "'" & Trim$(tgWebFeed(iFeedLoop).sName) & "',"
    SQLQuery = SQLQuery + "'" & Trim$(tgWebFeed(iFeedLoop).sPassword) & "',"
    SQLQuery = SQLQuery + "'" & Format$(gNow(), "yyyy-mm-dd") & "'"
    ilRet = gWebExecSql(SQLQuery, "", sgImportPath & "WebUpdateSpots")

    mWebUpdateSpots = ilRet

End Function

'***************************************************************************************
'*
'* Function Name: mWebGetDeleteSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: 'Receive the deleted spots from the web
'*
'***************************************************************************************
Private Function mWebGetDeleteSpots(iFeedLoop As Integer, iVehLoop As Integer) As Integer

    Dim ilRet As Integer

    SQLQuery = "Select astCode, attCode, StationName, VehicleName, Advt, Prod, PledgeStartDate, PledgeEndDate,"
    SQLQuery = SQLQuery + " PledgeStartTime , PledgeEndTime, SpotLen, Cart, CreativeTitle, ISCI"
    SQLQuery = SQLQuery + " From VW_DeletedRadioSpots"
    SQLQuery = SQLQuery + " Where VehicleName = '" & Trim$(tgWebFeed(iFeedLoop).sName) & "'"
    SQLQuery = SQLQuery + " And StationName = '" & Trim$(tgMVef(iVehLoop).sName) & "'"
    SQLQuery = SQLQuery + " And Password = '" & Trim$(tgWebFeed(iFeedLoop).sPassword) & "'"
    SQLQuery = SQLQuery + " And astCode NOT IN (select astCode from Spots)"
    ilRet = gWebExecSql(SQLQuery, "", sgImportPath & "WebDeleted")

    mWebGetDeleteSpots = ilRet

End Function

'***************************************************************************************
'*
'* Function Name: mWebRemDeletedSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: 'Remove the deleted spots on the web
'*
'***************************************************************************************
Private Function mWebRemDeletedSpots(iFeedLoop As Integer, iVehLoop As Integer) As Integer

    Dim ilRet As Integer

    SQLQuery = "sp_RemoveDeletedSpots "
    SQLQuery = SQLQuery + "'" & Trim$(tgMVef(iVehLoop).sName) & "',"
    SQLQuery = SQLQuery + "'" & Trim$(tgWebFeed(iFeedLoop).sName) & "',"
    SQLQuery = SQLQuery + "'" & Trim$(tgWebFeed(iFeedLoop).sPassword) & "'"
    ilRet = gWebExecSql(SQLQuery, "", sgImportPath & "WebUpdateDeletes")

    mWebRemDeletedSpots = ilRet

End Function

'***************************************************************************************
'*
'* Function Name: mWebDebugReset
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: '*** For debug only - This resets the database so spots will be sent again
'*
'***************************************************************************************






'***************************************************************************************
'*
'* Function Name: mReadWebSpots
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: Used to read the Station Traffic spot files received from the web
'*           and store them into tgWebSpots
'*
'***************************************************************************************

Private Function mReadWebSpots(iIsDelete As Integer, sFileName As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  hmFile                        slFlie                        slDateTime                *
'*                                                                                        *
'******************************************************************************************


    Dim ilRet As Integer
    Dim llIdx As Long
    Dim slFileAndPath As String

    Dim sAstCode As String
    Dim sAttCode As String
    Dim sStationName As String
    Dim sVehicleName As String
    Dim sAdvt As String
    Dim sProd As String
    Dim sPledgeStartDate As String
    Dim sPledgeEndDate As String
    Dim sPledgeStartTime As String
    Dim sPledgeEndTime As String
    Dim sSpotLen As String
    Dim sCart As String
    Dim sCreativeTitle As String
    Dim sISCI As String


    'On Error GoTo ErrHand

    'hmWebFile = FreeFile
    ilRet = 0
    slFileAndPath = sgImportPath & sFileName
    ReDim tgWebSpots(0 To 0)

    'Open slFileAndPath For Input Access Read As hmWebFile
    ilRet = gFileOpen(slFileAndPath, "Input Access Read", hmWebFile)
    If ilRet <> 0 Then
        mReadWebSpots = False
        Exit Function
    End If

    ' Skip past the header definition record.
    Input #hmWebFile, sAstCode, sAttCode, sStationName, sVehicleName, sAdvt, sProd, sPledgeStartDate, sPledgeEndDate, sPledgeStartTime, sPledgeEndTime, sSpotLen, sCart, sCreativeTitle, sISCI
    If ilRet <> 0 Then
        mReadWebSpots = True
        Exit Function
    End If

    llIdx = 0

    Do While Not EOF(hmWebFile)
        DoEvents
        Input #hmWebFile, sAstCode, sAttCode, sStationName, sVehicleName, sAdvt, sProd, sPledgeStartDate, sPledgeEndDate, sPledgeStartTime, sPledgeEndTime, sSpotLen, sCart, sCreativeTitle, sISCI
        tgWebSpots(llIdx).lAstCode = CLng(sAstCode)
        tgWebSpots(llIdx).lAttCode = CLng(sAttCode)
        tgWebSpots(llIdx).sStationName = sStationName
        tgWebSpots(llIdx).sVehicleName = sVehicleName
        tgWebSpots(llIdx).sAdvt = sAdvt
        tgWebSpots(llIdx).sProd = sProd
        tgWebSpots(llIdx).sPledgeStartDate = sPledgeStartDate
        If iIsDelete = True Then
            'If we're deleting recs then set the end date to the start date minus one day
            tgWebSpots(llIdx).sPledgeEndDate = DateAdd("d", -1, sPledgeStartDate)
        Else
            tgWebSpots(llIdx).sPledgeEndDate = sPledgeEndDate
        End If
        tgWebSpots(llIdx).sPledgeStartTime = sPledgeStartTime
        tgWebSpots(llIdx).sPledgeEndTime = sPledgeEndTime
        tgWebSpots(llIdx).iSpotLen = sSpotLen
        tgWebSpots(llIdx).sCart = sCart
        tgWebSpots(llIdx).sCreativeTitle = sCreativeTitle
        tgWebSpots(llIdx).sISCI = sISCI
        llIdx = llIdx + 1
        ReDim Preserve tgWebSpots(0 To llIdx)
    Loop

    Close hmWebFile
    mReadWebSpots = True
    Exit Function

'ErrHand:
'    ilRet = 1
'    Resume Next

End Function



'***************************************************************************************
'*
'* Function Name: mGetVehicleCode
'*
'* Created: August, 18  By: J. Dutschke
'*
'* Modified:            By:
'*
'* Comments: Returns the station code for a given station name
'*
'***************************************************************************************
Private Function mGetVehicleCode(sName As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

    On Error GoTo Err_GetVehicleCode
    Dim ilLoop As Integer

    mGetVehicleCode = -1
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
        If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(sName), 1) = 0 Then
            mGetVehicleCode = tgMVef(ilLoop).iCode
            Exit Function
        End If
    Next ilLoop

    Exit Function
Err_GetVehicleCode:
End Function




'***************************************************************************************
'*
'* Function Name: mGetAdvertiserCode
'*
'* Created: August, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Returns the advertiser code for a given advertiser name
'*
'***************************************************************************************
Private Function mGetAdvertiserCode(sName As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilCRet                                                  *
'******************************************************************************************

    Dim ilLoop As Integer

    If Trim(sName) = "" Then
        mGetAdvertiserCode = -1
        Exit Function
    End If
    For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
        DoEvents
        If StrComp(Trim$(tgCommAdf(ilLoop).sName), Trim$(sName), 1) = 0 Then
            mGetAdvertiserCode = tgCommAdf(ilLoop).iCode
            Exit Function
        End If
    Next ilLoop
    mGetAdvertiserCode = mAddAdvertiser(sName, "D")
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mAddAdvertiser                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add advertiser                 *
'*                                                     *
'*******************************************************
Private Function mAddAdvertiser(slName As String, slDirect As String) As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim tlAdf As ADF
'    Dim hlAdf As Integer    'file handle
    Dim ilAdfRecLen As Integer  'Record length
    Dim slSyncDate As String
    Dim slSyncTime As String

    mAddAdvertiser = -1
    gGetSyncDateTime slSyncDate, slSyncTime
    tlAdf.iCode = 0         'Internal code number for advertiser
    tlAdf.sName = slName    'Name
    tlAdf.sAbbr = Left$(slName, 7) 'Abbreviation
    tlAdf.sProduct = ""     'Product name
    tlAdf.iSlfCode = 0      'Salesperson code number
    tlAdf.iAgfCode = 0      'Agency code number
    tlAdf.sBuyer = ""       'Buyers name
    tlAdf.sCodeRep = ""     'Rep advertiser Code
    tlAdf.sCodeAgy = ""
    tlAdf.sCodeStn = ""     'Station advertiser Code
    tlAdf.iMnfComp(0) = 0   'Competitive code
    tlAdf.iMnfComp(1) = 0   'Competitive code
    tlAdf.iMnfExcl(0) = 0   'Program Exclusions code
    tlAdf.iMnfExcl(1) = 0   'Program Exclusions code
    tlAdf.sCppCpm = "N"     'P=CPP; M=CPM; N=N/A
    For ilLoop = 0 To 3 Step 1
        tlAdf.sDemo(ilLoop) = ""    'First-four Demo target
        tlAdf.iMnfDemo(ilLoop) = 0
        tlAdf.lTarget(ilLoop) = 0
    Next ilLoop
    tlAdf.sCreditRestr = "N"
    tlAdf.lCreditLimit = 0
    tlAdf.sPaymRating = "1"
    tlAdf.sShowISCI = "N"
    tlAdf.iMnfSort = 0
    tlAdf.sBillAgyDir = "D"
    tlAdf.sCntrAddr(0) = "*************************"
    For ilLoop = 1 To 2 Step 1
        tlAdf.sCntrAddr(ilLoop) = ""
        tlAdf.sBillAddr(ilLoop) = ""
    Next ilLoop
    tlAdf.iArfLkCode = 0
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tlAdf.sPhone = "______________"
    tlAdf.sFax = "__________"
    tlAdf.iArfCntrCode = 0
    tlAdf.iArfInvCode = 0
    tlAdf.sCntrPrtSz = "N"
    '12/17/06-Change to tax by agency or vehicle
    'tlAdf.sSlsTax(0) = "N"
    'tlAdf.sSlsTax(1) = "N"
    tlAdf.iTrfCode = 0
    tlAdf.sCrdApp = "A" '"R" changed 6/30/00 via jim request
    tlAdf.sCrdRtg = ""
    tlAdf.iPnfBuyer = 0
    tlAdf.iPnfPay = 0
    tlAdf.iPct90 = 0
    slStr = ""
    gStrToPDN slStr, 2, 6, tlAdf.sCurrAR
    slStr = ""
    gStrToPDN slStr, 2, 6, tlAdf.sUnbilled
    slStr = ""
    gStrToPDN slStr, 2, 6, tlAdf.sHiCredit
    slStr = ""
    gStrToPDN slStr, 2, 6, tlAdf.sTotalGross
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tlAdf.iDateEntrd(0), tlAdf.iDateEntrd(1)
    tlAdf.iNSFChks = 0
    tlAdf.iDateLstInv(0) = 0  'No date
    tlAdf.iDateLstInv(1) = 0
    tlAdf.iDateLstPaym(0) = 0  'No date
    tlAdf.iDateLstPaym(1) = 0
    tlAdf.iAvgToPay = 0
    tlAdf.iLstToPay = 0
    tlAdf.iNoInvPd = 0
    tlAdf.sNewBus = "N"
    tlAdf.iEndDate(0) = 0
    tlAdf.iEndDate(1) = 0
    tlAdf.iMerge = 0
    tlAdf.iUrfCode = 2
    tlAdf.sState = "A"
    tlAdf.iCrdAppDate(0) = 0
    tlAdf.iCrdAppDate(1) = 0
    tlAdf.iCrdAppTime(0) = 0
    tlAdf.iCrdAppTime(1) = 0
    tlAdf.sPkInvShow = "T"
    tlAdf.sBkoutPoolStatus = "N"
    tlAdf.sUnused2 = ""
    tlAdf.sRateOnInv = "Y"
    tlAdf.iMnfBus = 0
    tlAdf.lGuar = 0
    tlAdf.sAllowRepMG = "N"
    tlAdf.sBonusOnInv = "Y"
    tlAdf.sRepInvGen = "I"
    tlAdf.iMnfInvTerms = 0
    tlAdf.sPolitical = "N"
    tlAdf.sAddrID = ""
    tlAdf.iTrfCode = 0

'    hlAdf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        MsgBox "Open Advertiser Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'        Exit Function
'    End If
    ilAdfRecLen = Len(tlAdf)
    ilRet = btrInsert(hmAdf, tlAdf, ilAdfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        'MsgBox "Insert Advertiser Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Insert Error"
        gLogMsg "Insert Advertiser " & slName & ", Error:" & str(ilRet), "FeedImport.txt", False
        Exit Function
    End If
    ilUpper = UBound(tgCommAdf)
    tgCommAdf(ilUpper).iCode = tlAdf.iCode
    tgCommAdf(ilUpper).sName = Trim$(tlAdf.sName)
    tgCommAdf(ilUpper).sAbbr = Trim$(tlAdf.sAbbr)
    tgCommAdf(ilUpper).sBillAgyDir = tlAdf.sBillAgyDir
    tgCommAdf(ilUpper).sState = tlAdf.sState
    tgCommAdf(ilUpper).sAllowRepMG = tlAdf.sAllowRepMG
    tgCommAdf(ilUpper).sBonusOnInv = tlAdf.sBonusOnInv
    tgCommAdf(ilUpper).sRepInvGen = tlAdf.sRepInvGen
    tgCommAdf(ilUpper).sFirstCntrAddr = tlAdf.sCntrAddr(0)
    tgCommAdf(ilUpper).sAddrID = Trim$(tlAdf.sAddrID)
    tgCommAdf(ilUpper).iTrfCode = tlAdf.iTrfCode
    ilUpper = ilUpper + 1
    'ReDim Preserve tgCommAdf(1 To ilUpper) As ADFEXT
    ReDim Preserve tgCommAdf(0 To ilUpper) As ADFEXT
    mAddAdvertiser = tlAdf.iCode
End Function

'***************************************************************************************
'*
'* Function Name: mGetAdvtProdCode
'*
'* Created: August, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments: Returns the product code for a given advertiser product name
'*
'***************************************************************************************
Private Function mGetAdvtProdCode(ilAdvtCode As Integer, slAdvtProdName As String) As Boolean
    On Error GoTo Err_mGetAdvtProdCode
    Dim ilRet As Integer
'    Dim hlPrf As Integer
    Dim ilPrfRecLen As Integer
    Dim tlPrf As PRF

    If Trim(slAdvtProdName) = "" Then
        tmFsf.lPrfCode = 0
        tmFsf.iMnfComp1 = 0
        tmFsf.iMnfComp2 = 0
        mGetAdvtProdCode = True
        Exit Function
    End If

    mGetAdvtProdCode = False

'    hlPrf = CBtrvTable(TWOHANDLES)
'    ilRet = btrOpen(hlPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        MsgBox "Open Prf Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'        Exit Function
'    End If
    ilPrfRecLen = Len(tlPrf)
    tmPrfSrchKey.iCode = ilAdvtCode
    ilRet = btrGetGreaterOrEqual(hmPrf, tlPrf, ilPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If tlPrf.iAdfCode <> ilAdvtCode Then
            Exit Do
        End If
        If StrComp(Trim(slAdvtProdName), Trim(tlPrf.sName), vbTextCompare) = 0 Then
            tmFsf.lPrfCode = tlPrf.lCode
            tmFsf.iMnfComp1 = tlPrf.iMnfComp(0)
            tmFsf.iMnfComp2 = tlPrf.iMnfComp(1)

            mGetAdvtProdCode = True
'            ilRet = btrClose(hlPrf)
'            btrDestroy hlPrf
            Exit Function
        End If
        ilRet = btrGetNext(hmPrf, tlPrf, ilPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    'Add product

'    ilRet = btrClose(hlPrf)
'    btrDestroy hlPrf
    tmFsf.lPrfCode = mAddProduct(ilAdvtCode, slAdvtProdName)
    tmFsf.iMnfComp1 = 0
    tmFsf.iMnfComp2 = 0
    Exit Function

Err_mGetAdvtProdCode:
    mGetAdvtProdCode = False
End Function

'***************************************************************************************
'*
'* Procedure Name: mBuildFsfFile
'*
'* Created: August, 2004    By: D. Smith
'*
'* Modified:                By:
'*
'* Comments: Prepare tmFsf record to either be inserted or updated
'*
'***************************************************************************************

Private Function mBuildFsfFile(iIsDelete As Integer, iFeedLoop As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIdx                         ilTemp                        llTemp                    *
'*  slTemp                                                                                *
'******************************************************************************************


    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim tlFsf As FSF
    Dim ilRevNumber As Integer
    Dim llPrevFsfCode As Long
    Dim llRecCode As Long

    ilRet = True
    mBuildFsfFile = False
    hmFsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imFsfRecLen = Len(tlFsf)

    For ilLoop = 0 To UBound(tgWebSpots) - 1 Step 1
        DoEvents
        ilFound = False
        ilRet = btrGetEqual(hmFsf, tlFsf, imFsfRecLen, tgWebSpots(ilLoop).lAstCode, INDEXKEY5, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE And (tgWebSpots(ilLoop).lAstCode = tlFsf.lAstCode) Then
            tmFsf = tlFsf
            ilFound = True
        End If
        Do While ilRet = BTRV_ERR_NONE And (tgWebSpots(ilLoop).lAstCode = tlFsf.lAstCode)
            ilRet = btrGetNext(hmFsf, tlFsf, imFsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            'Always save the version with the highest revision number
            If tmFsf.iRevNo < tlFsf.iRevNo And (tgWebSpots(ilLoop).lAstCode = tlFsf.lAstCode) Then
                tmFsf = tlFsf
            End If
        Loop

        If ilFound Then
            If tmFsf.sSchStatus = "F" Then
                'It was fully scheduled so we insert a new record
                ilRevNumber = tmFsf.iRevNo + 1
                llPrevFsfCode = tmFsf.lCode  'tmFsf.lPrevFsfCode
                ilRet = mInsertFsf(iIsDelete, ilLoop, iFeedLoop, ilRevNumber, llPrevFsfCode)
            Else
                'It was not fully scheduled so we update the old record
                ilRevNumber = tmFsf.iRevNo
                llPrevFsfCode = tmFsf.lCode    'tmFsf.lPrevFsfCode
                llRecCode = tmFsf.lCode
                ilRet = mUpdateFsf(iIsDelete, ilLoop, iFeedLoop, ilRevNumber, llPrevFsfCode, tmFsf.sSchStatus, llRecCode)
            End If
        Else
            'If it's a delete record and not found then ignore it else insert it
            If iIsDelete = False Then
                ilRevNumber = 0
                llPrevFsfCode = 0
                ilRet = mInsertFsf(iIsDelete, ilLoop, iFeedLoop, ilRevNumber, llPrevFsfCode)
            End If
        End If
    Next ilLoop
    ilRet = btrClose(hmFsf)
    btrDestroy hmFsf

    If ilRet = True Then
        mBuildFsfFile = True
    Else
        mBuildFsfFile = False
    End If

End Function

'***************************************************************************************
'*
'* Procedure Name: mInsertFsf
'*
'* Created: August, 2004    By: D. Smith
'*
'* Modified:                By:
'*
'* Comments: Insert new tmFsf record
'*
'***************************************************************************************

Private Function mInsertFsf(iIsDelete As Integer, iWebLoop As Integer, iFeedLoop As Integer, iRevNumber As Integer, lPrevFsfCode As Long) As Integer

    Dim ilIdx As Integer
    Dim ilRet As Integer

    mInsertFsf = False
    tmFsf.iAdfCode = mGetAdvertiserCode(tgWebSpots(iWebLoop).sAdvt)

    For ilIdx = 0 To 6 Step 1
        tmFsf.iDays(ilIdx) = 0
    Next ilIdx

    For ilIdx = gWeekDayStr(tgWebSpots(iWebLoop).sPledgeStartDate) To gWeekDayStr(tgWebSpots(iWebLoop).sPledgeEndDate) Step 1
        tmFsf.iDays(ilIdx) = 1
    Next ilIdx

    tmFsf.iFnfCode = tgWebFeed(iFeedLoop).iCode
    gPackDate tgWebSpots(iWebLoop).sPledgeEndDate, tmFsf.iEndDate(0), tmFsf.iEndDate(1)
    gPackTime tgWebSpots(iWebLoop).sPledgeEndTime, tmFsf.iEndTime(0), tmFsf.iEndTime(1)
    gPackTime Format(gNow(), "hh:mm:ss"), tmFsf.iEnterTime(0), tmFsf.iEnterTime(1)
    gPackDate Format(gNow(), "yyyy-mm-dd"), tmFsf.iEnterDate(0), tmFsf.iEnterDate(1)
    tmFsf.iFnfCode = tgWebFeed(iFeedLoop).iCode
    tmFsf.iLen = tgWebSpots(iWebLoop).iSpotLen
    tmFsf.iNoSpots = 1
    tmFsf.iRevNo = iRevNumber
    tmFsf.iRunEvery = 0
    gPackTime tgWebSpots(iWebLoop).sPledgeStartTime, tmFsf.iStartTime(0), tmFsf.iStartTime(1)
    gPackDate tgWebSpots(iWebLoop).sPledgeStartDate, tmFsf.iStartDate(0), tmFsf.iStartDate(1)
    tmFsf.iUrfCode = tgUrf(0).iCode
    tmFsf.iVefCode = mGetVehicleCode(tgWebSpots(iWebLoop).sStationName)
    tmFsf.lAstCode = tgWebSpots(iWebLoop).lAstCode
    tmFsf.lCifCode = gGetFeedCopy(hmCif, hmCpf, tgWebFeed(iFeedLoop).iMcfCode, tmFsf.iAdfCode, tmFsf.iLen, tgWebSpots(iWebLoop).sPledgeEndDate, tgWebSpots(iWebLoop).sProd, tgWebSpots(iWebLoop).sISCI, tgWebSpots(iWebLoop).sCreativeTitle)
    tmFsf.lCode = 0
    tmFsf.lPrevFsfCode = lPrevFsfCode
    tmFsf.sDyWk = "W"
    tmFsf.sRefID = ""
    tmFsf.sSchStatus = "N"
    ilRet = mGetAdvtProdCode(tmFsf.iAdfCode, tgWebSpots(iWebLoop).sProd)
    ilRet = btrInsert(hmFsf, tmFsf, imFsfRecLen, INDEXKEY5)
    If ilRet = BTRV_ERR_NONE Then
        mInsertFsf = True
    End If
End Function

'***************************************************************************************
'*
'* Procedure Name: mUpdateFsf
'*
'* Created: August, 2004    By: D. Smith
'*
'* Modified:                By:
'*
'* Comments: Update existing tmFsf record
'*
'***************************************************************************************

Private Function mUpdateFsf(iIsDelete As Integer, iWebLoop As Integer, iFeedLoop As Integer, iRevNumber As Integer, lPrevFsfCode As Long, sSchdStatus As String, lRecCode As Long) As Integer

    Dim ilIdx As Integer
    Dim ilRet As Integer
    Dim tlFsf As FSF

    mUpdateFsf = False
    tmFsf.iAdfCode = mGetAdvertiserCode(tgWebSpots(iWebLoop).sAdvt)

    For ilIdx = 0 To 6 Step 1
        tmFsf.iDays(ilIdx) = 0
    Next ilIdx

    For ilIdx = gWeekDayStr(tgWebSpots(iWebLoop).sPledgeStartDate) To gWeekDayStr(tgWebSpots(iWebLoop).sPledgeEndDate) Step 1
        tmFsf.iDays(ilIdx) = 1
    Next ilIdx

    tmFsf.iFnfCode = tgWebFeed(iFeedLoop).iCode
    gPackDate tgWebSpots(iWebLoop).sPledgeEndDate, tmFsf.iEndDate(0), tmFsf.iEndDate(1)
    gPackTime tgWebSpots(iWebLoop).sPledgeEndTime, tmFsf.iEndTime(0), tmFsf.iEndTime(1)
    gPackTime Format(gNow(), "hh:mm:ss"), tmFsf.iEnterTime(0), tmFsf.iEnterTime(1)
    gPackDate Format(gNow(), "yyyy-mm-dd"), tmFsf.iEnterDate(0), tmFsf.iEnterDate(1)
    tmFsf.iFnfCode = tgWebFeed(iFeedLoop).iCode
    tmFsf.iLen = tgWebSpots(iWebLoop).iSpotLen
    tmFsf.iNoSpots = 1
    tmFsf.iRevNo = iRevNumber
    tmFsf.iRunEvery = 0
    gPackTime tgWebSpots(iWebLoop).sPledgeStartTime, tmFsf.iStartTime(0), tmFsf.iStartTime(1)
    gPackDate tgWebSpots(iWebLoop).sPledgeStartDate, tmFsf.iStartDate(0), tmFsf.iStartDate(1)
    tmFsf.iUrfCode = tgUrf(0).iCode
    tmFsf.iVefCode = mGetVehicleCode(tgWebSpots(iWebLoop).sStationName)
    tmFsf.lAstCode = tgWebSpots(iWebLoop).lAstCode
    tmFsf.lCifCode = 0
    tmFsf.lCode = lRecCode
    tmFsf.lPrevFsfCode = lPrevFsfCode
    tmFsf.sDyWk = "W"
    tmFsf.sRefID = ""
    If iIsDelete = True Then
        tmFsf.sSchStatus = "N"
    Else
        tmFsf.sSchStatus = tmFsf.sSchStatus
    End If
    ilRet = mGetAdvtProdCode(tmFsf.iAdfCode, tgWebSpots(iWebLoop).sProd)
    ilRet = btrGetEqual(hmFsf, tlFsf, imFsfRecLen, lRecCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        mUpdateFsf = False
        Exit Function
    End If
    ilRet = btrUpdate(hmFsf, tmFsf, imFsfRecLen)
    If ilRet = BTRV_ERR_NONE Then
        mUpdateFsf = True
    Else
        mUpdateFsf = False
    End If

End Function

'***************************************************************************************
'*
'*      Procedure Name:mGetCompletedSpots
'*
'*             Created:10/17/93      By:D. LeVine
'*            Modified:09/09/04      By:J. Dutschke
'*
'* Comments:
'*
'***************************************************************************************
Private Function mGetCompletedSpots(hlLcf As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                        tlLcfSrchKey                                            *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim ilRecLen As Integer
    Dim tlLcf As LCF
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    ReDim tmLcfArray(0 To 0) As Ext_LCF

    ilIdx = 0

    ilExtLen = Len(tlLcf)
    ilRet = btrGetFirst(hlLcf, tlLcf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, 0)

    ' Prepare to execute an extended operation.
    btrExtClear hlLcf   'Clear any previous extend operation

    llNoRec = gExtNoRec(ilExtLen)
    Call btrExtSetBounds(hlLcf, llNoRec, -1, "UC", "LCF", "") '"EG") 'Set extract limits (all records)

    ilOffSet = gFieldOffset("lcf", "lcfType")
    tlCharTypeBuff.sType = "O"
    ilRet = btrExtAddLogicConst(hlLcf, BTRV_KT_STRING, ilOffSet, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

    ilOffSet = gFieldOffset("lcf", "lcfStatus")
    tlCharTypeBuff.sType = "C"
    ilRet = btrExtAddLogicConst(hlLcf, BTRV_KT_STRING, ilOffSet, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

    ilOffSet = gFieldOffset("lcf", "lcfAffPost")
    tlCharTypeBuff.sType = "C"
    ilRet = btrExtAddLogicConst(hlLcf, BTRV_KT_STRING, ilOffSet, Len(tlCharTypeBuff.sType), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)

    ilRet = btrExtAddField(hlLcf, 0, ilExtLen) 'Extract the whole record
    ilRet = btrExtGetNext(hlLcf, tlLcf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tlLcf)  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlLcf, tlLcf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            LSet tmLcfArray(ilIdx) = tlLcf
            tmLcfArray(ilIdx).lRecPos = llRecPos    ' Save this so we can update this record later.
            ilIdx = ilIdx + 1
            ReDim Preserve tmLcfArray(0 To ilIdx) As Ext_LCF
            ilRet = btrExtGetNext(hlLcf, tlLcf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlLcf, tlLcf, ilExtLen, llRecPos)
            Loop
        Loop
    End If

End Function

'***************************************************************************************
'*
'*      Procedure Name:mGetCompletedSpots
'*
'*             Created:10/17/93      By:D. LeVine
'*            Modified:09/09/04      By:J. Dutschke
'*
'* Comments:
'*
'***************************************************************************************
Private Function mSetDaysToComplete(hlLcf As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim ilRecLen As Integer
    Dim tlLcf As LCF

    ilRecLen = Len(tlLcf)
    For ilIdx = LBound(tmLcfArray) To UBound(tmLcfArray) - 1
        If tmLcfArray(ilIdx).iSeqNo = 999 Then
            ' This record has been marked as completed.
            ilRet = btrGetDirect(hlLcf, tlLcf, ilRecLen, tmLcfArray(ilIdx).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                tlLcf.sAffPost = "S"
                ilRet = btrUpdate(hlLcf, tlLcf, ilRecLen)
            End If
        End If
    Next

End Function

'***************************************************************************************
'*
'*      Procedure Name:mGetSpots
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:J. Dutschke
'*
'*            Comments:
'*
'***************************************************************************************
Private Function mGetSpots(hlSdf, ilVefCode As Integer, llLogDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llDate                        ilIndex                       ilWkIndex                 *
'*  ilDay                         llTime                        ilTime                    *
'*  ilLen                                                                                 *
'******************************************************************************************

    Dim ilRet As Integer
    Dim slDate As String
    Dim ilIdx As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim tlSdf As SDF
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim llChfCode As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tmSdfArray(0 To 0) As SDF
    ilIdx = 0
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf)  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    slDate = Format$(llLogDate, "m/d/yy")
    gPackDate slDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf)
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)

        ' And only records where the ChfCode = 0
        llChfCode = 0
        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llChfCode, 4)

        ' And on the records where the date is equal to the passed in log date
        slDate = Format$(llLogDate, "m/d/yy")
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record
        ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tmSdfArray(ilIdx) = tlSdf
                ilIdx = ilIdx + 1
                ReDim Preserve tmSdfArray(0 To ilIdx) As SDF
                ilExtLen = Len(tlSdf)
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If

End Function

'***************************************************************************************
'*
'* Function Name: mGetASTCodeFromFSFCode
'*
'* Created: Sept, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************
Private Function mGetASTCodeFromFSFCode(FSFCode As Long) As Long
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llFsfCode                     tlFnf                                                   *
'******************************************************************************************

    Dim hlFile As Integer
    Dim ilRecLen As Integer
    Dim ilFnfCode As Integer
    Dim iLoop As Integer
    Dim tlFsf As FSF
    Dim ilRet As Integer

    mGetASTCodeFromFSFCode = -1

    ' Get the FSFCode and ASTCode from the FNF table.
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & "FSF.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_SHARE, BTRV_LOCK_NONE)
    If ilRet <> 0 Then
        Exit Function
    End If
    ilRecLen = Len(tlFsf)
    ilRet = btrGetEqual(hlFile, tlFsf, ilRecLen, FSFCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        btrClose (hlFile)
        btrDestroy (hlFile)
        Exit Function
    End If
    btrClose (hlFile)
    btrDestroy (hlFile)

    ilFnfCode = tlFsf.iFnfCode
    ' Next look up this FnfCode in the FNF Feed Names table and see if this spot is going back to the web server.
    For iLoop = LBound(tgWebFeed) To UBound(tgWebFeed)
        If tgWebFeed(iLoop).iCode = ilFnfCode Then
            ' We found the feed were looking for.
            If Len(Trim(tgWebFeed(iLoop).sIPAddress)) > 0 Then
                ' And it has a URL so we can send this spot back to the web server.
                mGetASTCodeFromFSFCode = tlFsf.lAstCode
                smURLToPostTo = Trim(tgWebFeed(iLoop).sIPAddress)
            End If
            Exit Function
        End If
    Next

End Function

'***************************************************************************************
'*
'* Function Name: gExportWebSpots
'*
'* Created: Sept, 2004  By: J. Dutschke
'*
'* Modified:              By:
'*
'* Comments:
'*
'***************************************************************************************
Public Function gExportWebSpots()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llTime                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilLcfIdx As Integer
    Dim ilSdfIdx As Integer
    Dim llFsfCode As Long
    Dim llASTCode As Long
    Dim hlLcf As Integer
    Dim hlSdf As Integer
    Dim ilStatus As Integer
    Dim slDateRan As String
    Dim slTimeRan As String
    Dim slDateTimeRan As String
    Dim slSql As String

    ' Get feed names if necessary
    If UBound(tgWebFeed) = 0 Then
        ilRet = gBuildFeedArray()
        If UBound(tgWebFeed) = 0 Then
            Exit Function
        End If
    End If

    hlSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSdf, "", sgDBPath & "SDF.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_SHARE, BTRV_LOCK_NONE)
    If ilRet <> 0 Then
        Exit Function
    End If

    hlLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLcf, "", sgDBPath & "LCF.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_SHARE, BTRV_LOCK_NONE)
    If ilRet <> 0 Then
        Exit Function
    End If

    ' Get days that need to be exported to the web site for each vehicle.
    ilRet = mGetCompletedSpots(hlLcf)
    ' Loop on each Vehicle/Date from the Lcf table.
    For ilLcfIdx = LBound(tmLcfArray) To UBound(tmLcfArray) - 1
        gUnpackDateLong tmLcfArray(ilLcfIdx).iLogDate(0), tmLcfArray(ilLcfIdx).iLogDate(1), llDate
        ' Find this spot in the SDF table. We do this so we can get back to the FSF, FNF table.
        ilRet = mGetSpots(hlSdf, tmLcfArray(ilLcfIdx).iVefCode, llDate)

        ' Loop on each spot from the Sdf table.
        For ilSdfIdx = LBound(tmSdfArray) To UBound(tmSdfArray) - 1
            llFsfCode = tmSdfArray(ilSdfIdx).lFsfCode
            llASTCode = mGetASTCodeFromFSFCode(llFsfCode)
            DoEvents
            If llASTCode <> -1 Then
                ' Send this spot back to the web server.
                ' This is the call being made from the ASP web pages. We want to call it the same way.
                ' sp_PostSpot 378, '0', '12/25/1999 7:20:00 PM'
                slDateRan = Format(llDate, "m/d/yy")
                gUnpackTime tmSdfArray(ilSdfIdx).iTime(0), tmSdfArray(ilSdfIdx).iTime(1), "A", "1", slTimeRan
                slDateTimeRan = slDateRan & " " & slTimeRan
                If tmSdfArray(ilSdfIdx).sSchStatus = "S" Or _
                    tmSdfArray(ilSdfIdx).sSchStatus = "O" Or _
                    tmSdfArray(ilSdfIdx).sSchStatus = "G" Then
                    ilStatus = 0
                Else
                    ilStatus = 1
                End If
                slSql = "sp_PostSpot " & llASTCode & ", '" & ilStatus & "', '" & slDateTimeRan & "'"
                ilRet = gWebExecSql(slSql, smURLToPostTo, "UpdateSpotsResults")
                'ilRet = gWebExecSql(slSQL, "", "UpdateSpotsResults")

                ' Mark this record as complete so mSetDaysToComplete knows to update it's status.
                tmLcfArray(ilLcfIdx).iSeqNo = 999
            End If
        Next
    Next

    Call mSetDaysToComplete(hlLcf)

    btrClose (hlLcf)
    btrDestroy (hlLcf)

    btrClose (hlSdf)
    btrDestroy (hlSdf)

End Function

Public Function gGetFeedCopy(hlCif As Integer, hlCpf As Integer, ilMcfCode As Integer, ilAdfCode As Integer, ilSpotLen As Integer, slDate As String, slProduct As String, slISCI As String, slCreativeTitle As String) As Long
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llDate                        ilCifOk                                                 *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilCpfRecLen As Integer
    Dim ilCifRecLen As Integer
    Dim ilMcfRecLen As Integer
    Dim llCifDate As Long
    Dim llCifCode As Long
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    gGetFeedCopy = 0
    If tmMcf.iCode <> ilMcfCode Then
        If ilMcfCode > 0 Then
            ilMcfRecLen = Len(tmMcf)
            hmMcf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                'MsgBox "Open Media Code Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gLogMsg "Open Media Code Error:" & str(ilRet), "FeedImport.txt", False
                Exit Function
            End If
            tmMcfSrchKey.iCode = ilMcfCode
            ilRet = btrGetEqual(hmMcf, tmMcf, ilMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmMcf)
                btrDestroy hmMcf
                Exit Function
            End If
            ilRet = btrClose(hmMcf)
            btrDestroy hmMcf
        End If
    End If
    If ilMcfCode = 0 Then
        tmMcf.iCode = 0
        If tgSpf.sUseCartNo <> "N" Then
            Exit Function
        End If
    End If
    DoEvents
    mBuildPurgeInventory hlCif, ilMcfCode, tmPurgeInv()
    ilExtLen = Len(tmCif)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hlCif   'Clear any previous extend operation
    ilCifRecLen = Len(tmCif)
    ilCpfRecLen = Len(tmCpf)
    ilRet = btrGetFirst(hlCif, tmCif, ilCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilMcfCode
        ilOffSet = gFieldOffset("Cif", "CifMcfCode")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        tlIntTypeBuff.iType = ilAdfCode
        ilOffSet = gFieldOffset("Cif", "CifAdfCode")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        tlCharTypeBuff.sType = "A"
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        ilRet = btrExtAddField(hlCif, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Function
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                tmCpfSrchKey.lCode = tmCif.lcpfCode
                ilRet = btrGetEqual(hlCpf, tmCpf, ilCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    DoEvents
                    'If (StrComp(slProduct, Trim$(tmCpf.sName), vbTextCompare) = 0) And (StrComp(slISCI, Trim$(tmCpf.sISCI), vbTextCompare) = 0) And (StrComp(slCreativeTitle, Trim$(tmCpf.sCreative), vbTextCompare) = 0) Then
                    If (StrComp(Trim$(slISCI), Trim$(tmCpf.sISCI), vbTextCompare) = 0) Then
                        'Set date
                        tmCifSrchKey.lCode = tmCif.lCode
                        ilRet = btrGetEqual(hlCif, tmCif, ilCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            gUnpackDateLong tmCif.iUsedDate(0), tmCif.iUsedDate(1), llCifDate
                            If gDateValue(slDate) > llCifDate Then
                                gPackDate slDate, tmCif.iUsedDate(0), tmCif.iUsedDate(1)
                            End If
                            gUnpackDateLong tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), llCifDate
                            If gDateValue(slDate) > llCifDate Then
                                gPackDate slDate, tmCif.iRotEndDate(0), tmCif.iRotEndDate(1)
                            End If
                            tmCif.iUrfCode = tgUrf(0).iCode
                            ilRet = btrUpdate(hlCif, tmCif, ilCifRecLen)
                        End If
                        gUnpackDateLong tmCpf.iRotEndDate(0), tmCpf.iRotEndDate(1), llCifDate
                        If gDateValue(slDate) > llCifDate Then
                            gPackDate slDate, tmCpf.iRotEndDate(0), tmCpf.iRotEndDate(1)
                            ilRet = btrUpdate(hlCpf, tmCpf, ilCifRecLen)
                        End If
                        gGetFeedCopy = tmCif.lCode
                        Exit Function
                    End If
                End If
                ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    'Assign Copy
    llCifCode = mAddCopy(hlCif, hlCpf, ilAdfCode, ilSpotLen, slDate, slProduct, slISCI, slCreativeTitle)
    gGetFeedCopy = llCifCode
    Exit Function
End Function

Private Sub mBuildPurgeInventory(hlCif As Integer, ilMcfCode As Integer, tlSortCode() As SORTCODE)
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilCifRecLen As Integer
    Dim ilLowLimit As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)

    'On Error GoTo mBuildPurgeInventoryErr
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    imPurgeInvMcf = -1
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        imPurgeInvMcf = -1
        ilLowLimit = 0
    End If
    
    If imPurgeInvMcf = ilMcfCode Then
        Exit Sub
    End If
    ReDim Preserve tlSortCode(0 To 0) As SORTCODE
    If ilMcfCode = 0 Then
        Exit Sub
    End If
    ilExtLen = Len(tmCif)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hlCif   'Clear any previous extend operation
    ilCifRecLen = Len(tmCif)
    ilRet = btrGetFirst(hlCif, tmCif, ilCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilMcfCode
        ilOffSet = gFieldOffset("Cif", "CifMcfCode")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        tlCharTypeBuff.sType = "P"
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        ilRet = btrExtAddField(hlCif, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                DoEvents
                gUnpackDateForSort tmCif.iDateEntrd(0), tmCif.iDateEntrd(1), slDate
                tlSortCode(UBound(tlSortCode)).sKey = slDate & "\" & Trim$(str$(tmCif.lCode))
                ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 1) As SORTCODE
                ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tmCif, ilExtLen, llRecPos)
                Loop
            Loop
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If
    End If
    Exit Sub
mBuildPurgeInventoryErr:
    ilRet = 1
    Resume Next
End Sub

Private Function mOpenFiles() As Integer
    Dim ilRet As Integer

    mOpenFiles = True
    hmAdf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "Open Advertiser Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gLogMsg "Open Advertiser File Error:" & str(ilRet), "FeedImport.txt", False
        mOpenFiles = False
        Exit Function
    End If
    hmPrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "Open Product Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gLogMsg "Open Product File Error:" & str(ilRet), "FeedImport.txt", False
        mOpenFiles = False
        Exit Function
    End If
    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "Open Copy Inventory Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gLogMsg "Open Copy Inventory File Error:" & str(ilRet), "FeedImport.txt", False
        mOpenFiles = False
        Exit Function
    End If
    hmCpf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'MsgBox "Open Copy Product Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gLogMsg "Open Copy Product File Error:" & str(ilRet), "FeedImport.txt", False
        mOpenFiles = False
        Exit Function
    End If

End Function

Private Sub mCloseFiles()
    Dim ilRet As Integer

    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf

End Sub

Private Function mAddProduct(ilAdfCode As Integer, slAdvtProdName As String) As Long
    Dim tlPrf As PRF
    Dim ilRet As Integer
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer
    Dim ilPrfRecLen As Integer
    Dim ilLoop As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String

    mAddProduct = 0
    ilAdfRecLen = Len(tlAdf)
    ilPrfRecLen = Len(tlPrf)
    gGetSyncDateTime slSyncDate, slSyncTime
    tlPrf.lCode = 0
    tlPrf.iAdfCode = ilAdfCode
    tlPrf.sName = Trim$(slAdvtProdName)
    tlPrf.iMnfComp(0) = 0
    tlPrf.iMnfComp(1) = 0
    tlPrf.iMnfExcl(0) = 0
    tlPrf.iMnfExcl(1) = 0
    tmAdfSrchKey.iCode = ilAdfCode
    ilRet = btrGetEqual(hmAdf, tlAdf, ilAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tlPrf.iPnfBuyer = tlAdf.iPnfBuyer
    Else
        tlPrf.iPnfBuyer = 0
    End If
    tlPrf.sCppCpm = ""
    For ilLoop = 0 To 3
        tlPrf.iMnfDemo(ilLoop) = 0
        tlPrf.lTarget(ilLoop) = 0
        tlPrf.lLastCPP(ilLoop) = 0
        tlPrf.lLastCPM(ilLoop) = 0
    Next ilLoop
    tlPrf.sState = "A"
    tlPrf.iUrfCode = 2 'Use first record retained for user
    tlPrf.iRemoteID = tgUrf(0).iRemoteUserID
    tlPrf.lAutoCode = tlPrf.lCode
    ilRet = btrInsert(hmPrf, tlPrf, ilPrfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        'MsgBox "Insert Product Error:" & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Insert Error"
        gLogMsg "Insert Product " & Trim$(slAdvtProdName) & ", Error:" & str(ilRet), "FeedImport.txt", False
        Exit Function
    Else
        tlPrf.iRemoteID = tgUrf(0).iRemoteUserID
        tlPrf.lAutoCode = tlPrf.lCode
        tlPrf.iSourceID = tgUrf(0).iRemoteUserID
        gPackDate slSyncDate, tlPrf.iSyncDate(0), tlPrf.iSyncDate(1)
        gPackTime slSyncTime, tlPrf.iSyncTime(0), tlPrf.iSyncTime(1)
        ilRet = btrUpdate(hmPrf, tlPrf, ilPrfRecLen)
    End If

End Function

Private Function mAddCopy(hlCif As Integer, hlCpf As Integer, ilAdfCode As Integer, ilSpotLen As Integer, slDate As String, slProduct As String, slISCI As String, slCreativeTitle As String) As Long
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mAddCopyErr                                                                           *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'code number
    Dim slNameCode As String
    Dim tlCif As CIF
    Dim tlCpf As CPF
    Dim ilCifRecLen As Integer
    Dim ilCpfRecLen As Integer


    mAddCopy = 0
    'Test if ISCI Used

    ilCifRecLen = Len(tlCif)
    ilCpfRecLen = Len(tlCpf)
    If tgSpf.sUseCartNo <> "N" Then
        If Trim$(tmMcf.sReuse) = "N" Then
            Exit Function
        End If
        tlCif.iMcfCode = tmMcf.iCode
        tlCif.lCode = 0
        'Get Number from Purge list
        For ilLoop = LBound(tmPurgeInv) To UBound(tmPurgeInv) - 1 Step 1
            If Trim$(tmPurgeInv(ilLoop).sKey) <> "" Then
                DoEvents
                slNameCode = tmPurgeInv(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If ilRet = CP_MSG_NONE Then
                    tmCifSrchKey.lCode = Val(slCode)
                    ilRet = btrGetEqual(hlCif, tmCif, ilCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If tmCif.sPurged = "P" Then
                            tlCif = tmCif
                            tmCif.sPurged = "H"
                            ilRet = btrUpdate(hlCif, tmCif, ilCifRecLen)
                            Exit For
                        Else
                            tmPurgeInv(ilLoop).sKey = ""
                        End If
                    Else
                        tmPurgeInv(ilLoop).sKey = ""
                    End If
                Else
                    tmPurgeInv(ilLoop).sKey = ""
                End If
            End If
        Next ilLoop
        If tlCif.lCode = 0 Then
            'MsgBox "No Copy Inventory available", vbOkOnly + vbCritical + vbApplicationModal, "Inventory"
            gLogMsg "No Copy Inventory available", "FeedImport.txt", False
            Exit Function
        End If
        tlCif.lCode = 0
    Else
        tlCif.lCode = 0
        tlCif.sName = ""
        tlCif.sCut = ""
        tlCif.iMcfCode = 0
    End If
    DoEvents
    tlCif.iEtfCode = 0
    tlCif.iEnfCode = 0
    tlCif.iAdfCode = ilAdfCode
    tlCif.sReel = ""
    tlCif.iLen = ilSpotLen
    'tlCif.lCpfCode set within save
    For ilLoop = 0 To 1 Step 1
        tlCif.iMnfComp(ilLoop) = 0
    Next ilLoop
    tlCif.sPurged = "A"
    Select Case tmMcf.sCartDisp
        Case "N"
            tlCif.sCartDisp = "N"
        Case "S"
            tlCif.sCartDisp = "S"
        Case "P"
            tlCif.sCartDisp = "P"
        Case "A"
            tlCif.sCartDisp = "A"
        Case Else
            tlCif.sCartDisp = "A"
    End Select
    tlCif.sTapeDisp = "N"
    '2/2/12: Replaced mcfTapeDisp with mcfSuppressOnExport. Use default of N
    'Select Case tmMcf.sTapeDisp
    '    Case "N"
    '        tlCif.sTapeDisp = "N"
    '    Case "R"
    '        tlCif.sTapeDisp = "R"
    '    Case "D"
    '        tlCif.sTapeDisp = "D"
    '    Case "A"
    '        tlCif.sTapeDisp = "A"
    '    Case Else
    '        tlCif.sTapeDisp = "D"
    'End Select
    tlCif.iNoTimesAir = 0
    tlCif.sHouse = "N"
    tlCif.sCleared = "N"
    tlCif.iMnfAnn = 0
    tlCif.iPurgeDate(0) = 0
    tlCif.iPurgeDate(1) = 0
    gPackDate slDate, tlCif.iUsedDate(0), tlCif.iUsedDate(1)
    gPackDate slDate, tlCif.iRotStartDate(0), tlCif.iRotStartDate(1)
    gPackDate slDate, tlCif.iRotEndDate(0), tlCif.iRotEndDate(1)
    tlCif.iUrfCode = tgUrf(0).iCode
    tlCif.sPrint = "N"

    If (Trim$(slProduct) <> "") Or (Trim$(slISCI) <> "") Or (Trim$(slCreativeTitle) <> "") Then
        tlCpf.lCode = 0
        tlCpf.sName = slProduct
        tlCpf.sISCI = slISCI
        tlCpf.sCreative = slCreativeTitle
        gPackDate slDate, tlCpf.iRotEndDate(0), tlCpf.iRotEndDate(1)
        ilRet = btrInsert(hlCpf, tlCpf, ilCpfRecLen, INDEXKEY0)
        tlCif.lcpfCode = tlCpf.lCode
    Else
        tlCif.lcpfCode = 0
    End If

    ilRet = btrInsert(hlCif, tlCif, ilCifRecLen, INDEXKEY0)
    mAddCopy = tlCif.lCode
    Exit Function
mAddCopyErr: 'VBC NR
    On Error GoTo 0
    Exit Function

End Function
