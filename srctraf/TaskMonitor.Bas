Attribute VB_Name = "TaskMonitor"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SAFREC.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSafSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule function
'
'
Option Explicit
Option Compare Text

Private smZone As String
Private tmTmf As TMF
Private imTmfRecLen As Integer
Private tmTmfSrchKey1 As TMFKEY1    'Vef key record image
'7967
'Public dgWvImportLast As Date
'Public igWVImportElapsed As Integer
'does client have a 'webserver' in traffic.ini, is there a 'web' vendor, and is the web uptodate?  Set in function gVendorToWebAllowed
Public bgVendorToWebAllowed As Boolean
'8198
Public gUsingWeb As Boolean
'8129
Public Enum VendorWvmAlert
    None = 0
    ImportTest = 1
    ImportMissed = 2
    ExportTest = 3
    ConnectionIssue = 4
    ExportRunning = 5
    WebVendorSent = 6
    ExportMissed = 7
    ExportError = 8
    NonMonitorIssue = 9
    ImportError = 10
    ManagerNotRunning = 11
    UserSetTime = 12
End Enum
Private Enum ExportInfoIndex
    'Export order: code,attcode,vendoridcode,hasbeensent errormessage exportMondayDate,ProcessedDateTime EnteredDateTime
    Code = 0
    attCode = 1
    vendoridcode = 2
    hasbeensent = 3
    result = 3
    Message = 4
    mondaydate = 5
    ProcessedDateTime = 6
    entereddatetime = 7
    Service = 7
    spotOkcount = 8
    spotErrorcount = 9
End Enum
Public Type ServiceControllerInfo
    Mode As String
    GenerateDebug As Boolean
    GenerateFile As Boolean
    ImportLast As String
    ImportFiles As String
    IsRunning As Boolean
    ImportSpan As String
End Type

Public gStopWatch As New StopWatch


Public Sub gOpenTmf()
    Dim ilRet As Integer
    
    smZone = Left$(gGetLocalTZName(), 1)
    hgTmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hgTmf, "", sgDBPath & "Tmf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open Task Monitor File, Error = " & str$(ilRet), vbOKOnly + vbInformation, "Warning"
    End If
    imTmfRecLen = Len(tmTmf)
End Sub

Public Sub gCloseTmf()
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = btrClose(hgTmf)
    btrDestroy hgTmf
End Sub

Public Sub gUpdateTaskMonitor(ilWhichTime As Integer, slTaskCode As String)
    'ilWhichTime: 0=Running; 1=Start, 2= End
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slSQLQuery As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim llDate As Long
    
    On Error GoTo ErrHand
    If smZone = "" Then
        smZone = Left$(gGetLocalTZName(), 1)
    End If
    slDateTime = Now
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
    Select Case smZone
        Case "E"
            slStr = slNowDate & " " & slNowTime
        Case "C"
            slStr = DateAdd("h", 1, slNowDate & " " & slNowTime)
        Case "M"
            slStr = DateAdd("h", 2, slNowDate & " " & slNowTime)
        Case "P"
            slStr = DateAdd("h", 3, slNowDate & " " & slNowTime)
        Case Else
            slStr = slNowDate & " " & slNowTime
    End Select
    slNowDate = Format$(slStr, "m/d/yy")
    slNowTime = Format$(slStr, "h:mm:ssAM/PM")
    imTmfRecLen = Len(tmTmf)
    tmTmfSrchKey1.sTaskCode = Trim$(slTaskCode)
    ilRet = btrGetEqual(hgTmf, tmTmf, imTmfRecLen, tmTmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        If ilWhichTime = 1 Then
            sgTmfStatus = "S"
            gPackDate slNowDate, tmTmf.iRunningDate(0), tmTmf.iRunningDate(1)
            gPackTime slNowTime, tmTmf.iRunningTime(0), tmTmf.iRunningTime(1)
            gUnpackDateLong tmTmf.i1stStartRunDate(0), tmTmf.i1stStartRunDate(1), llDate
            If gDateValue(slNowDate) <> llDate Then
                gPackDate slNowDate, tmTmf.i1stStartRunDate(0), tmTmf.i1stStartRunDate(1)
                gPackTime slNowTime, tmTmf.i1stStartRunTime(0), tmTmf.i1stStartRunTime(1)
            End If
            gPackDate slNowDate, tmTmf.iStartRunDate(0), tmTmf.iStartRunDate(1)
            gPackTime slNowTime, tmTmf.iStartRunTime(0), tmTmf.iStartRunTime(1)
            tmTmf.sStatus = "S"
        ElseIf ilWhichTime = 2 Then
            If sgTmfStatus = "S" Then
                sgTmfStatus = "C"
            End If
            gPackDate slNowDate, tmTmf.iRunningDate(0), tmTmf.iRunningDate(1)
            gPackTime slNowTime, tmTmf.iRunningTime(0), tmTmf.iRunningTime(1)
            gUnpackDateLong tmTmf.i1stEndRunDate(0), tmTmf.i1stEndRunDate(1), llDate
            If gDateValue(slNowDate) <> llDate Then
                gPackDate slNowDate, tmTmf.i1stEndRunDate(0), tmTmf.i1stEndRunDate(1)
                gPackTime slNowTime, tmTmf.i1stEndRunTime(0), tmTmf.i1stEndRunTime(1)
            End If
            gPackDate slNowDate, tmTmf.iEndRunDate(0), tmTmf.iEndRunDate(1)
            gPackTime slNowTime, tmTmf.iEndRunTime(0), tmTmf.iEndRunTime(1)
            tmTmf.sStatus = sgTmfStatus
        Else
            gPackDate slNowDate, tmTmf.iRunningDate(0), tmTmf.iRunningDate(1)
            gPackTime slNowTime, tmTmf.iRunningTime(0), tmTmf.iRunningTime(1)
        End If
        ilRet = btrUpdate(hgTmf, tmTmf, imTmfRecLen)
    End If
    Exit Sub
ErrHand:
    Exit Sub
End Sub
'7967 '8129
Public Function gMonitorVendor(tmcWebConnectIssue As Timer) As String
    'G for success, E for error, Y for other ('G'reen,I need to use E as it's what is expected,'Y'ellow)
    Dim slRet As String
    Dim slSql As String
    Dim ilFirst As Integer
    Dim dlNow As Date
    Dim llRet As Long
    Dim ilVendorIssue As VendorWvmAlert
    Dim blSetTodaysDate As Boolean
    '8133
    Dim myServiceFacts(2) As ServiceControllerInfo
    Dim c As Integer
    'fake...so stays the same as affiliate
    Dim bgVendorExportSent As Boolean
    
    bgVendorExportSent = False
On Error GoTo ERRBOX
    'return true unless error
    'when today - the elapse is greater than the last time we tested.... so when I run on 6/30 and then I test again on 6/30 - 1 day is 6/29...won't run until now is 6/31
    slRet = "G"
    blSetTodaysDate = False
    If bgVendorToWebAllowed Then
        ilFirst = 0
        If igWVImportElapsed = 0 Then
            'as minutes
            igWVImportElapsed = 30
            'testing!  Must remove!
          '  igWVImportElapsed = 2
            ilFirst = 1
        End If
        'dan is this ok?
        'dlNow = Now()
        dlNow = gNow()
        'it's noon. Last ran 11:45  11:30 > 11:45?
        If DateAdd("n", -igWVImportElapsed, dlNow) > dgWvImportLast Then
            dgWvImportLast = dlNow
            slRet = mAlertRunByOther()
            If slRet = "X" Then
                slRet = "G"
                'clears monitor issues, but not issues from vit/vet
                blSetTodaysDate = mAlertClearWVMBasic()
                'if this is false, could not connect
                If mQueryServiceController(myServiceFacts) Then
                     'are we currently exporting?
                     If Not bgVendorExportSent Then
                         'mAlerts...will write to the vet table (which gets picked up by auf)
                         'returns if in test mode or there are some waiting in queue
                         ilVendorIssue = mAlertsForWebVendors(True, myServiceFacts)
                         If ilVendorIssue <> VendorWvmAlert.None Then
                            'non monitor issue:  regular errors that were handled in function above.
                             If ilVendorIssue <> VendorWvmAlert.NonMonitorIssue Then
                                 mAlertMonitorIssue ilVendorIssue, True
                             End If
                             Select Case ilVendorIssue
                                 Case VendorWvmAlert.NonMonitorIssue, VendorWvmAlert.ExportMissed, VendorWvmAlert.ExportTest, VendorWvmAlert.ExportError
                                     slRet = "E"
                                 Case VendorWvmAlert.ConnectionIssue
                                     slRet = "Y"
                                     mVendorToWebFailed tmcWebConnectIssue
'                                 Case VendorWvmAlert.ExportError
'                                     slRet = "E"
                                 Case VendorWvmAlert.ExportRunning
                                     slRet = "Y"
                                     If mAlertsClearWebVendors(True) Then
                                        blSetTodaysDate = True
                                     End If
                             End Select
                         Else
                             If mAlertsClearWebVendors(True) Then
                                blSetTodaysDate = True
                            End If
                        End If
                    End If
                    'import
                    ilVendorIssue = mAlertsForWebVendors(False, myServiceFacts)
                    If ilVendorIssue <> VendorWvmAlert.None Then
                        If ilVendorIssue <> VendorWvmAlert.NonMonitorIssue Then
                            mAlertMonitorIssue ilVendorIssue, True
                        End If
                        Select Case ilVendorIssue
                            Case VendorWvmAlert.NonMonitorIssue, VendorWvmAlert.ImportTest, VendorWvmAlert.ImportMissed, VendorWvmAlert.ImportError
                                slRet = "E"
                            Case VendorWvmAlert.ConnectionIssue
                                slRet = "Y"
                                mVendorToWebFailed tmcWebConnectIssue
                        End Select
                    Else
                        mAlertsClearWebVendors False
                    End If
                    'now the manager!
                    For c = 0 To 2
                       If myServiceFacts(c).Mode = "M" Then
                           If Not myServiceFacts(c).IsRunning Then
                               slRet = "E"
                               mAlertMonitorIssue VendorWvmAlert.ManagerNotRunning, True
                           End If
                           Exit For
                        End If
                    Next c
                Else
                    If slRet <> "E" Then
                        slRet = "Y"
                    End If
                    mAlertMonitorIssue VendorWvmAlert.ConnectionIssue, True
                    mVendorToWebFailed tmcWebConnectIssue
                End If
                If Not bgVendorExportSent Then
                    If slRet = "G" Then
                        gUpdateTaskMonitor ilFirst, "WVM"
                        If Not blSetTodaysDate Then
                           'set a cleared date when didn't have to do anything to block other users
                           mAlertSetUserGeneric
                        End If
                   ' in general, the old issue is still set as old
                    ElseIf ilVendorIssue = VendorWvmAlert.NonMonitorIssue And Not blSetTodaysDate Then
                       mAlertSetUserGeneric
                    End If
                   Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dgWvImportLast '& " and testing import span from : " & dlNow & " and timer span: " & igWVImportElapsed
                End If
            ElseIf Not bgVendorExportSent Then
                Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dlNow & " doesn't need to run since previous user ran since " & Format(DateAdd("n", -igWVImportElapsed, dlNow), sgSQLTimeForm)
            End If 'another user hasn't run
            If bgVendorExportSent Then
                bgVendorExportSent = False
                slRet = "Y"
                mAlertsClearWebVendors True
                mAlertMonitorIssue VendorWvmAlert.WebVendorSent, True
                Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dgWvImportLast '& " and testing import span from : " & dlNow & " and timer span: " & igWVImportElapsed
            End If
        End If 'time to run
    Else
        slRet = "Y"
    End If
CLEANUP:
    gMonitorVendor = slRet
    Exit Function
ERRBOX:
    slRet = "E"
    GoTo CLEANUP
End Function
Private Sub mAlertSetUserGeneric()
    Dim llCount As Long
    Dim ilCef As Integer
    Dim slDate As String
    Dim slTime As String
    'fake to make match affiliate
    Dim igUstCode As Integer
    
    igUstCode = 0
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)

    ilCef = VendorWvmAlert.UserSetTime
    'this is a fake record. always status C,sub E, cefcode above
    SQLQuery = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & igUstCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'C' and auftype = 'V' and aufsubtype  = 'E' and aufcefCode = " & ilCef
    If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
        gHandleError "TrafficErrors.txt", "taskMonitor-mAlertSetUserGeneric"
        Exit Sub
    End If
    If llCount = 0 Then
        SQLQuery = "insert into AUF_Alert_User (aufcode,aufCreateUstCode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufcefcode,aufCleardate,aufClearTime) values (0," & igUstCode & ",'" & slDate & "','" & slTime & "','R','V','E'," & ilCef & ",'" & slDate & "','" & slTime & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            gHandleError "TrafficErrors.txt", "taskMonitor-mAlertSetUserGeneric"
        End If
    End If
End Sub

Private Function mExecWebSQLCount(sSQL As String, Optional ilRetryMax = 0) As Long
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
    Dim aDataArray() As String
    Dim slCount As String
    Dim WebCmds As New WebCommands
     '7967 to traffic
'    If igDemoMode Or Not gUsingWeb Then
'        mExecWebSQLCount = 0
'        GoTo Cleanup
'    End If
    mExecWebSQLCount = -1    ' -1 is an error condition.
    If Not gLoadOptionTrafficThenAffiliate("WebServer", "RootURL", slRootURL) Then
        gLogMsg "Error: gCheckWebSession: LoadOption RootURL Error", "TrafficErrors.Txt", False
      '  gMsgBox "Error: gCheckWebSession: LoadOption RootURL Error"
        GoTo CLEANUP
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOptionTrafficThenAffiliate("WebServer", "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To ilRetryMax
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
            ' Parse out the response we got.
            alRecordsArray = Split(slResponse, vbCrLf)
            If Not IsArray(alRecordsArray) Then
                GoTo CLEANUP
            End If
            ' We have to have back at least two records. The first one is the column headers.
            ' The rest of the entries are the data itself.
            If UBound(alRecordsArray) < 2 Then
                ' If the table is empty, we will get back at least one record containing the column
                ' definitions of the table itself, but no data records.
                mExecWebSQLCount = 0
                GoTo CLEANUP
            End If
            slCount = Replace(alRecordsArray(1), """", "")
            If IsNumeric(slCount) Then
                mExecWebSQLCount = CLng(slCount)
                GoTo CLEANUP
            Else
                mExecWebSQLCount = 0
                GoTo CLEANUP
            End If
'            ' Each record we get back is a comma delimited string. In this case were only interested
'            ' in the first record.
'            aDataArray = Split(alRecordsArray(1), ",")
'            If Not IsArray(aDataArray) Then
'                Exit Function
'            End If
'            mExecWebSQLCount = UBound(aDataArray)
'            'If mExecWebSQLCount < 1 Then
'            If mExecWebSQLCount = 1 Then
'                ' The SQL Statement asked for count
'                Exit Function
'            End If
        End If
        If ilRetryMax > 0 Then
            Call gSleep(1)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "mExecWebSQLCount, " & ilRetryMax + 1 & " tries were exceeded. sSQL = " & sSQL & " llErrorCode = " & llErrorCode, "TrafficErrors.Txt", False
CLEANUP:
    Erase alRecordsArray
    Exit Function
    
ErrHandler:
    llErrorCode = err.Number
    Resume Next
End Function
Private Function mAlertsClearWebVendors(blExport As Boolean) As Boolean
    'does not clear monitor's messages!
    Dim slSubType As String
    Dim slSql As String
    Dim blRet As Boolean
    Dim llCount As Long
    
    blRet = False
    If blExport Then
        slSubType = "'E'"
    Else
        slSubType = "'I'"
    End If
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & tgUrf(0).iCode & ", aufcleardate = '" & Format(gNow(), sgSQLDateForm) & "', aufcleartime = '" & Format(gNow(), sgSQLTimeForm) & "' where auftype = 'V' and aufsubtype = " & slSubType & " and aufcefcode = 0 and aufStatus = 'R' "
    If gSQLAndReturn(slSql, False, llCount) <> 0 Then
        gHandleError "TrafficErrors.txt", "taskMonitor-gClearVendorAlters"
        Exit Function
    ElseIf llCount > 0 Then
        blRet = True
    End If
    mAlertsClearWebVendors = blRet
    Exit Function
ErrHand:
    gHandleError "", "taskMonitor-mAlertsClearWebVendors"
End Function
Private Sub mAlertMonitorIssue(ilWvmIssue As VendorWvmAlert, blAdd As Boolean)
    Dim llCefInsert As Long
    Dim slDate As String
    Dim slMon As String
    Dim slSubType As String
    Dim blAlreadyThere As Boolean
    
    'I don't think I need this since I clear all before writing
    blAlreadyThere = mAlertExistsWvmIssue(ilWvmIssue, slSubType)
    If blAdd And Not blAlreadyThere Then
        mAlertAddWvmIssue ilWvmIssue, slSubType
    ElseIf blAlreadyThere And Not blAdd Then
        mAlertClearWvmIssue ilWvmIssue, slSubType
    End If
End Sub
Private Function mAlertExistsWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim llCefInsert As Long
    Dim auf_rst As ADODB.Recordset
    
    blRet = False
    slSubType = mAlertSubType(ilWvmIssue)
    llCefInsert = ilWvmIssue
On Error GoTo ERRBOX
    slSql = "select count(*) as Total from Auf_Alert_User where aufStatus = 'R' and auftype = 'V' and aufsubtype = '" & slSubType & "' and aufcefCode = " & llCefInsert
    'Set auf_rst = cnn.Execute(slSQL)
    Set auf_rst = gSQLSelectCall(slSql)
    If auf_rst!Total > 0 Then
        blRet = True
    End If
    mAlertExistsWvmIssue = blRet
    Exit Function
ERRBOX:
    mAlertExistsWvmIssue = False
End Function
Private Sub mAlertAddWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String)
    '8129 add direcly to auf..special messages only
    Dim slSql As String
    Dim llCefInsert As Long
    Dim slMon As String
    
    slMon = dgWvImportLast
    slMon = gObtainPrevMonday(slMon)
    slMon = Format(slMon, sgSQLDateForm)
    llCefInsert = ilWvmIssue
    slSql = "insert into AUF_Alert_User (aufcode,aufCreateUrfCode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufcefcode,aufMoWeekDate) values (0," & tgUrf(0).iCode & ",'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSubType & "'," & llCefInsert & ",'" & slMon & "')"
    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
       gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddWvmIssue"
    End If
End Sub
Private Function mAlertClearWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String) As Boolean
    '8129  returns true if any were cleared
    Dim slSql As String
    Dim llCefInsert As Long
    Dim slTime As String
    Dim llRet As Long
    Dim blRet As Boolean
    Dim slDate As String
    
    blRet = False
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)
    llCefInsert = ilWvmIssue
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUrfCode = " & tgUrf(0).iCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'R' and auftype = 'V' and aufsubtype = '" & slSubType & "' and aufcefCode = " & llCefInsert
    If gSQLAndReturn(slSql, False, llRet) <> 0 Then
        gHandleError "TrafficErrors.txt", "taskMonitor-mAlertClearWvmIssue"
    ElseIf llRet > 0 Then
        blRet = True
    End If
    mAlertClearWvmIssue = blRet
End Function
Private Function mAlertClearWVMBasic() As Boolean
    '8129 returns true if any cleared
    Dim slSql As String
    Dim slTime As String
    Dim slDate As String
    Dim slSubType As String
    Dim llRet As Long
    Dim blRet As Boolean
    
    blRet = False
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUrfCode = " & tgUrf(0).iCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'R' and auftype = 'V' and aufsubtype in ('I','E') and aufcefCode > 0"
    If gSQLAndReturn(slSql, False, llRet) <> 0 Then
        gHandleError "TrafficErrors.txt", "taskMonitor-mAlertClearWVMBasic"
    ElseIf llRet > 0 Then
        blRet = True
    End If
    mAlertClearWVMBasic = blRet
End Function

Private Function mAlertSubType(ilWvm As VendorWvmAlert) As String
    Dim slRet As String
    
    'note: VendorWvmAlert.ConnectionIssue shows as export
    slRet = ""
    Select Case ilWvm
        Case VendorWvmAlert.ImportTest, VendorWvmAlert.ImportMissed
            slRet = "I"
        Case Else
            slRet = "E"
    End Select
    mAlertSubType = slRet
End Function
'8133
Private Function mAlertAddVendorIssue(blIsExport As Boolean, slData As String) As String
    '8133 added vendor id and result to auf
    'add new only if it's an error.  I don't have to update because I already cleared them all!
    'EXPORT code,attcode,vendoridcode,hasbeensent errormessage exportMondayDate,ProcessedDateTime EnteredDateTime
    Dim slSql As String
    Dim slSub As String
    Dim slCode As String
    Dim slMondayDate As String
    Dim ilresult As Integer
    Dim ilVendorId As Integer
    Dim llAtt As Long
    Dim slRet As String
    
    slRet = ""
    If blIsExport Then
        If mParseWebVendorExport(slData, slCode, ilresult, slMondayDate, ilVendorId, llAtt) Then
            slRet = slCode
            slSub = "E"
        On Error GoTo ERRORBOX
            slSql = "insert into AUF_Alert_User (aufcode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufchfcode,aufMoWeekDate,aufVefCode,aufCountdown,aufUlfCode) values (0,'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSub & "'," & slCode & ",'" & slMondayDate & " '," & ilVendorId & "," & ilresult & "," & llAtt & ")"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ERRORBOX:
                gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddVendorIssue"
                Exit Function
            End If
        End If
    Else
        slSub = "I"
        If mParseWebVendorImport(slData, slMondayDate, ilVendorId) Then
            slSql = "insert into AUF_Alert_User (aufcode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufMoWeekDate,aufVefCode,aufCefCode) values (0,'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSub & "','" & Format(dgWvImportLast, sgSQLDateForm) & " '," & ilVendorId & "," & VendorWvmAlert.ImportMissed & ")"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ERRORBOX:
                gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddVendorIssue"
                Exit Function
            End If
        
        End If
    End If
    mAlertAddVendorIssue = slRet
    Exit Function
ERRORBOX:
    gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddVendorIssue"
    mAlertAddVendorIssue = ""
End Function
Private Function mParseWebVendorExport(slData As String, slCode As String, ilresult As Integer, slMondayDate As String, ilVendorId As Integer, llAtt As Long) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim slVendor As String
    
    blRet = True
    ilresult = 0
    slCode = ""
    slMondayDate = ""
    ilVendorId = 0
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 8 Then
        slCode = Trim$(Replace(slValues(ExportInfoIndex.Code), """", ""))
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.hasbeensent), """", ""))
        If IsNumeric(slVendor) Then
            ilresult = CInt(slVendor)
        Else
            blRet = False
        End If
        slMondayDate = Trim$(Replace(slValues(ExportInfoIndex.mondaydate), """", ""))
        slMondayDate = Format(slMondayDate, sgSQLDateForm)
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.vendoridcode), """", ""))
        If IsNumeric(slVendor) Then
            ilVendorId = CInt(slVendor)
        Else
            blRet = False
        End If
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.attCode), """", ""))
        If IsNumeric(slVendor) Then
            llAtt = CLng(slVendor)
        Else
            blRet = False
        End If
    Else
        blRet = False
    End If
CLEANUP:
    Erase slValues
    mParseWebVendorExport = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "TrafficErrors.txt", "taskMonitor-mParseWebVendorExport"
    GoTo CLEANUP
End Function
Private Function mParseWebVendorImport(slData As String, slMondayDate As String, ilVendorId As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim slVendor As String
    
    blRet = True
    slMondayDate = ""
    ilVendorId = 0
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 1 Then
        slVendor = Trim$(Replace(slValues(0), """", ""))
        If IsNumeric(slVendor) Then
            ilVendorId = CInt(slVendor)
        Else
            blRet = False
        End If
        slMondayDate = Trim$(Replace(slValues(1), """", ""))
        slMondayDate = Format(slMondayDate, sgSQLDateForm)
    Else
        blRet = False
    End If
CLEANUP:
    Erase slValues
    mParseWebVendorImport = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "TrafficErrors.txt", "taskMonitor-mParseWebVendorImport"
    GoTo CLEANUP
End Function

'Private Sub mAlertAddVendorIssue(blIsExport As Boolean, slMondayDate As String, slCode As String, slResult As String)
'    'add new only if it's an error.  I don't have to update because I already cleared them all!
'    'EXPORT code,attcode,vendoridcode,hasbeensent errormessage exportMondayDate,ProcessedDateTime EnteredDateTime
'    Dim ilresult As Integer
'    Dim slSql As String
'    Dim slSub As String
'
'    If blIsExport Then
'        slSub = "E"
'    Else
'        slSub = "I"
'    End If
'On Error GoTo ERRORBOX
'    If IsNumeric(slResult) Then
'        ilresult = CInt(slResult)
'        'insert only if an error (or didn't get ok yet)
'        If ilresult < -1 Or ilresult = 6 Then
'            slSql = "insert into AUF_Alert_User (aufcode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufchfcode,aufMoWeekDate) values (0,'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSub & "'," & slCode & ",'" & slMondayDate & " ')"
'            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
'                '6/13/16: Replaced GoSub
'                'GoSub ERRORBOX:
'                gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddVendorIssue"
'                Exit Sub
'            End If
'        End If
'    End If
'    Exit Sub
'ERRORBOX:
'    gHandleError "TrafficErrors.txt", "taskMonitor-mAlertAddVendorIssue"
'End Sub

'Private Function mAlertsForWebVendors(blExport As Boolean) As VendorWvmAlert
'    'return none if no issue, exportmissed,exporttest, connectionIssue, or exportrunning-which now means there was an error in the code!
'    'gExecWebSqlForVendor returns 0 if 'demo' mode.
'    Dim ilRet As VendorWvmAlert
'    Dim slDateRealMonday As String
'    Dim slDateMonday As String
'    Dim slSql As String
'    Dim slData() As String
'    Dim c As Long
'    Dim llCount As Long
'    Dim slCode As String
'    Dim slResult As String
'    Dim slMondayOfRecord As String
'    Dim myRst As ADODB.Recordset
'On Error GoTo errbox
'
'    ilRet = VendorWvmAlert.None
'    slDateMonday = dgWvImportLast
'    slDateMonday = gObtainPrevMonday(slDateMonday)
'    slDateMonday = Format(slDateMonday, sgSQLDateForm) & " " & Format(slDateMonday, sgSQLTimeForm)
'    slDateRealMonday = gObtainPrevMonday(gNow())
'    slDateRealMonday = Format(slDateRealMonday, sgSQLDateForm) & " " & Format(slDateRealMonday, sgSQLTimeForm)
'    If blExport Then
'        slSql = "Select count(*) as amount from vendorservicecontroller where Mode = 'E' and isrunning = 'Y'"
'        llCount = mExecWebSQLCount(slSql)
'        If llCount < 0 Then
'            ilRet = VendorWvmAlert.ConnectionIssue
'        ElseIf llCount > 0 Then
'            ilRet = VendorWvmAlert.ExportRunning
'        Else
'            slSql = "Select count(*) as amount from vendorservicecontroller where Mode = 'E' and GenerateFile = 'Y'"
'            llCount = mExecWebSQLCount(slSql)
'            If llCount < 0 Then
'                ilRet = VendorWvmAlert.ConnectionIssue
'            ElseIf llCount > 0 Then
'                ilRet = VendorWvmAlert.ExportTest
'            Else
'                'didn't go out? Note I don't check senttoaff..there shoulnd't be any = 0
'                slSql = "select count(*) as amount from webvendorexport where hasbeensent = 0 and exportMondayDate >='" & slDateRealMonday & "'"
'                llCount = mExecWebSQLCount(slSql)
'                If llCount > 0 Then
'                   ilRet = VendorWvmAlert.ExportMissed
'                ElseIf llCount < 0 Then
'                    ilRet = VendorWvmAlert.ConnectionIssue
'                Else
'                    'errors shown in export queue? First let's see if we can clear old
'                    slSql = "select count(*) as amount from webvendorexport where (hasbeensent =6 or hasbeensent < -1) and exportMondayDate >='" & slDateRealMonday & "'  AND sentToAff = 'Y' "
'                    llCount = mExecWebSQLCount(slSql)
'                    If llCount > 0 Then
'                        'We have errors up on web from before. we haven't taken care of old errors
'                        ilRet = VendorWvmAlert.NonMonitorIssue
'                        slSql = "select count(*) as Amount FROM AUF_ALERT_USER WHERE aufType = 'V'  AND aufStatus = 'R' AND aufSubType = 'E' AND aufchfcode <> 0 and aufMoWeekDate >='" & slDateMonday & "' "
'                        'Set myRst = cnn.Execute(slSQL)
'                        Set myRst = gSQLSelectCall(slSql)
'                        'all the same, which would be normal (haven't done export yet) doesn't need to go through each one
'                        If myRst!amount <> llCount Then
'                            mAlertsClearWebVendors True
'                            slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff = 'Y' "
'                            slSql = slSql & " AND exportMondayDate >= '" & slDateMonday & "'"
'                            llCount = gExecWebSQLForVendor(slData, slSql, True)
'                            If llCount > 1 Then
'                                mAlertsClearWebVendors True
'                                '0 is header
'                                For c = 1 To llCount - 1
'                                    If mRetrieveFromVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                                        mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                                    Else
'                                        ilRet = VendorWvmAlert.ExportError
'                                        GoTo Cleanup
'                                    End If
'                                Next c
'                            ElseIf llCount < 0 Then
'                                ilRet = VendorWvmAlert.ConnectionIssue
'                                GoTo Cleanup
'                            End If
'                        End If
'                    ElseIf llCount < 0 Then
'                        ilRet = VendorWvmAlert.ConnectionIssue
'                        GoTo Cleanup
'                    Else
'                        'no old errors. Clear.
'                        mAlertsClearWebVendors True
'                    End If
'                    slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff <> 'Y' "
'                    slSql = slSql & " AND exportMondayDate >= '" & slDateRealMonday & "'"
'                    llCount = gExecWebSQLForVendor(slData, slSql, True)
'                    If llCount > 1 Then
'                        ilRet = VendorWvmAlert.NonMonitorIssue
'                        '0 is header
'                        For c = 1 To llCount - 1
'                            'already senttoaff? return no slResult and false but now we only get !Y, so it's moot
'                            If mInsertToVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                                mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                            ElseIf Len(slResult) > 0 Then
'                                ilRet = VendorWvmAlert.ExportError
'                                GoTo Cleanup
'                            End If
'                        Next c
'                    ElseIf llCount < 0 Then
'                        ilRet = VendorWvmAlert.ConnectionIssue
'                    End If
'                End If
'            End If ' test mode
'        End If 'is running?
'    Else
'    '7967 end turn on/off
'
''    '6/7/16 Dan don't allow imports until I figure out how to clear them.
''    'now imports  skip -3 where astcode doesn't exist.
''    slSql = "Select code,attcode,vendoridcode,result,message,importMondayDate,ProcessedDateTime,service,spotcount,warningerrorcount from webimportarchive WHERE result < 0  AND sentToAff <> 'Y' "
''    slSql = slSql & " AND importMondayDate >= '" & slDate & "'"
''    llCount = gExecWebSQLForVendor(slData, slSql, True)
''    If llCount > 1 Then
''        '0 is header
''        For c = 1 To llCount - 1
''            If mInsertToVit(slData(c), slCode, slResult, slMondayOfRecord) Then
''                'don't add id doesn't exist to alert.  Don't know how to ever clear
''                If slResult <> "-3" Then
''                    mAlertAddVendorIssue False, slMondayOfRecord, slCode, slResult
''                End If
''            Else
''                ilRet = False
''            End If
''        Next c
''    End If
''    mAlertsClearWebVendors False, True
'    End If
'Cleanup:
'    If Not myRst Is Nothing Then
'        If (myRst.State And adStateOpen) <> 0 Then
'            myRst.Close
'        End If
'        Set myRst = Nothing
'    End If
'    mAlertsForWebVendors = ilRet
'    Exit Function
'errbox:
'    ilRet = VendorWvmAlert.ExportError
'    gHandleError "", "mAlertsForWebVendors"
'    GoTo Cleanup
'End Function
'8133
Private Function mAlertsForWebVendors(blExport As Boolean, myServiceFacts() As ServiceControllerInfo) As VendorWvmAlert
    'return none if no issue, exportmissed,exporttest, connectionIssue, or exportrunning-which now means there was an error in the code!
    'gExecWebSqlForVendor returns 0 if 'demo' mode.
    Dim ilRet As VendorWvmAlert
    Dim slDateRealMonday As String
    Dim slDateMonday As String
    Dim slDateShouldImport As String
    Dim slSql As String
    Dim slData() As String
    Dim c As Long
    Dim llCount As Long
    Dim slCode As String
    Dim myRst As ADODB.Recordset
    Dim blContinue As Boolean
    '8133
    Dim slWVECode As String
    Dim slControllerLastImported As String
On Error GoTo ERRBOX
    
    ilRet = VendorWvmAlert.None
    blContinue = True
    If blExport Then
        slDateMonday = dgWvImportLast
        slDateMonday = gObtainPrevMonday(slDateMonday)
        slDateMonday = Format(slDateMonday, sgSQLDateForm) & " " & Format(slDateMonday, sgSQLTimeForm)
        slDateRealMonday = gObtainPrevMonday(gNow())
        slDateRealMonday = Format(slDateRealMonday, sgSQLDateForm) & " " & Format(slDateRealMonday, sgSQLTimeForm)
        For c = 0 To 2
           If myServiceFacts(c).Mode = "E" Then
               If myServiceFacts(c).IsRunning Then
                   ilRet = VendorWvmAlert.ExportRunning
                   blContinue = False
                   Exit For
                ElseIf myServiceFacts(c).GenerateFile Then
                    ilRet = VendorWvmAlert.ExportTest
                    blContinue = False
               End If
               Exit For
           End If
        Next c
        If blContinue Then
                'didn't go out? Note I don't check senttoaff..there shoulnd't be any = 0
            slSql = "select count(*) as amount from webvendorexport where hasbeensent = 0 and exportMondayDate >='" & slDateRealMonday & "'"
            llCount = mExecWebSQLCount(slSql)
            If llCount > 0 Then
               ilRet = VendorWvmAlert.ExportMissed
            ElseIf llCount < 0 Then
                ilRet = VendorWvmAlert.ConnectionIssue
            Else
                'errors shown in export queue? First let's see if we can clear old
                slSql = "select count(*) as amount from webvendorexport where (hasbeensent =6 or hasbeensent < -1) and exportMondayDate >='" & slDateRealMonday & "'  AND sentToAff = 'Y' "
                llCount = mExecWebSQLCount(slSql)
                If llCount > 0 Then
                    'We have errors up on web from before. we haven't taken care of old errors
                    ilRet = VendorWvmAlert.NonMonitorIssue
                    slSql = "select count(*) as Amount FROM AUF_ALERT_USER WHERE aufType = 'V'  AND aufStatus = 'R' AND aufSubType = 'E' AND aufchfcode <> 0 and aufMoWeekDate >='" & slDateMonday & "' "
                    Set myRst = gSQLSelectCall(slSql)
                    'all the same, which would be normal (haven't done export yet) doesn't need to go through each one
                    If myRst!amount <> llCount Then
                        mAlertsClearWebVendors True
                        slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff = 'Y' "
                        slSql = slSql & " AND exportMondayDate >= '" & slDateMonday & "'"
                        llCount = gExecWebSQLForVendor(slData, slSql, True)
                        If llCount > 1 Then
                            mAlertsClearWebVendors True
                            '0 is header
                            For c = 1 To llCount - 1
                                '8133
                                If Len(mAlertAddVendorIssue(True, slData(c))) = 0 Then
                                    ilRet = VendorWvmAlert.ExportError
                                    GoTo CLEANUP
                                End If
'                                If mRetrieveFromVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                                    mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                                Else
'                                    ilRet = VendorWvmAlert.ExportError
'                                    GoTo Cleanup
'                                End If
                            Next c
                        ElseIf llCount < 0 Then
                            ilRet = VendorWvmAlert.ConnectionIssue
                            GoTo CLEANUP
                        End If
                    End If
                ElseIf llCount < 0 Then
                    ilRet = VendorWvmAlert.ConnectionIssue
                    GoTo CLEANUP
                Else
                    'no old errors. Clear.
                    mAlertsClearWebVendors True
                End If
                slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff <> 'Y' "
                slSql = slSql & " AND exportMondayDate >= '" & slDateRealMonday & "'"
                llCount = gExecWebSQLForVendor(slData, slSql, True)
                If llCount > 1 Then
                    ilRet = VendorWvmAlert.NonMonitorIssue
                    '0 is header
                    For c = 1 To llCount - 1
                        '8133
'                        'already senttoaff? return no slResult and false but now we only get !Y, so it's moot
'                        If mInsertToVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                            mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                        ElseIf Len(slResult) > 0 Then
'                            ilRet = VendorWvmAlert.ExportError
'                            GoTo Cleanup
'                        End If
                        slWVECode = mAlertAddVendorIssue(True, slData(c))
                        If Len(slWVECode) > 0 Then
                            slSql = "update webvendorexport set senttoaff = 'Y' where code = " & slWVECode
                            If mExecWebSQLWithRowsVendors(slSql) <> 1 Then
                                ilRet = VendorWvmAlert.ExportError
                                GoTo CLEANUP
                            End If
                        Else
                            ilRet = VendorWvmAlert.ExportError
                            GoTo CLEANUP
                        End If
                    Next c
                ElseIf llCount < 0 Then
                    ilRet = VendorWvmAlert.ConnectionIssue
                End If
            End If 'web export has been sent = 0
        End If 'not test or 'running': continue
    'import
    Else
        blContinue = True
        For c = 0 To 2
           If myServiceFacts(c).Mode = "I" Then
                If myServiceFacts(c).ImportFiles <> "" Then
                    ilRet = VendorWvmAlert.ImportTest
                    blContinue = False
                Else
                    If IsDate(myServiceFacts(c).ImportLast) And IsDate(myServiceFacts(c).ImportSpan) Then
                        slControllerLastImported = myServiceFacts(c).ImportLast
                        'this is how I store this info.  if date = 1/1/1970 then use hour; otherwise, take date and - 1 1/2/1970 = 1 day
                        If Day(myServiceFacts(c).ImportSpan) > 1 Then
                            slDateShouldImport = DateAdd("d", Day(myServiceFacts(c).ImportSpan) - 1, myServiceFacts(c).ImportLast)
                        Else
                            slDateShouldImport = DateAdd("h", Day(myServiceFacts(c).ImportSpan), myServiceFacts(c).ImportLast)
                        End If
                        'if span + last is out in future from now, we are ok.  2nd > 1st is positive number
                        If DateDiff("n", dgWvImportLast, slDateShouldImport) < 0 Then
                            ilRet = VendorWvmAlert.ImportMissed
                            blContinue = False
                        End If
                    Else
                        ilRet = VendorWvmAlert.ImportMissed
                        blContinue = False
                    End If
                End If
                Exit For
           End If
        Next c
        If blContinue Then
            '8133 test each webvendor to make sure updated recently
            If Len(slControllerLastImported) > 0 Then
                slSql = "select idcode,lastimportdate from webvendors WHERE importMethod > 0 and lastimportdate < '" & Format(slControllerLastImported, sgSQLDateForm) & "' "
                llCount = gExecWebSQLForVendor(slData, slSql, True)
                If llCount > 1 Then
                    mAlertAddVendorIssue False, slData(c)
                    ilRet = VendorWvmAlert.NonMonitorIssue
                End If
            Else
                'didn't find 'I' controller!
                ilRet = VendorWvmAlert.ImportError
            End If
        End If
    End If
CLEANUP:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAlertsForWebVendors = ilRet
    Exit Function
ERRBOX:
    ilRet = VendorWvmAlert.ExportError
    gHandleError "", "mAlertsForWebVendors"
    GoTo CLEANUP
End Function

Private Function mRetrieveFromVet(slData As String, slCode As String, slResult As String, slMondayDate As String) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim c As Integer
    Dim slAdd As String
    Dim slTemp As String
    Dim slProcessedDate As String
    Dim slProcessedTime As String
    Dim slMessage As String
    Dim llCount As Long
    
    blRet = True
    slResult = ""
    slCode = ""
    slMondayDate = ""
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 8 Then
        slCode = Trim$(Replace(slValues(ExportInfoIndex.Code), """", ""))
        slResult = Trim$(Replace(slValues(ExportInfoIndex.hasbeensent), """", ""))
        slMondayDate = Trim$(Replace(slValues(ExportInfoIndex.mondaydate), """", ""))
        slMondayDate = Format(slMondayDate, sgSQLDateForm)
    End If
CLEANUP:
    Erase slValues
    mRetrieveFromVet = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "TrafficErrors.txt", "taskMonitor-mRetrieveFromVet"
    GoTo CLEANUP
End Function

Private Function mInsertToVet(slData As String, slCode As String, slResult As String, slMondayDate As String) As Boolean
    'O: code, result (or hasbeen sent) monday export date
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim c As Integer
    Dim slAdd As String
    Dim slTemp As String
    Dim slProcessedDate As String
    Dim slProcessedTime As String
    Dim slMessage As String
    Dim llCount As Long
        
    blRet = True
    slResult = ""
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 8 Then
        slTemp = Trim$(Replace(slValues(8), """", ""))
        'has been sent already?
        If slTemp <> "Y" Then
            For c = 0 To 7
                slTemp = Trim$(Replace(slValues(c), """", ""))
                Select Case c
                    Case ExportInfoIndex.Code, ExportInfoIndex.attCode, ExportInfoIndex.vendoridcode, ExportInfoIndex.hasbeensent
                        slAdd = slAdd & slTemp & ","
                        If c = ExportInfoIndex.Code Then
                            slCode = slTemp
                        ElseIf c = ExportInfoIndex.hasbeensent Then
                            slResult = slTemp
                        End If
                    Case ExportInfoIndex.Message
                         slAdd = slAdd & "'" & slTemp & "',"
                         slMessage = slTemp
                    Case ExportInfoIndex.mondaydate, ExportInfoIndex.ProcessedDateTime
                        slAdd = slAdd & "'" & Format(slTemp, sgSQLDateForm) & "',"
                        If c = ExportInfoIndex.mondaydate Then
                            slMondayDate = Format(slTemp, sgSQLDateForm)
                        Else
                            slProcessedDate = Format(slTemp, sgSQLDateForm)
                            slProcessedTime = Format(slTemp, sgSQLTimeForm)
                            slAdd = slAdd & "'" & slProcessedTime & "',"
                        End If
                    Case ExportInfoIndex.entereddatetime
                        slAdd = slAdd & "'" & Format(slTemp, sgSQLDateForm) & "','" & Format(slTemp, sgSQLTimeForm) & "'"
                End Select
            Next c
            slAdd = "(" & slAdd & " )"
            slSql = "update VET_Vendor_Export set vetHasBeenSent =" & slResult & ", vetErrorMessage = '" & slMessage & "', vetProcessedDate = '" & slProcessedDate & "', vetprocessedtime = '" & slProcessedTime & "', vetExportMondayDate = '" & slMondayDate & "' where vetcode =" & slCode
            If gSQLAndReturn(slSql, False, llCount) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVet"
                mInsertToVet = False
                Erase slValues
                Exit Function
            End If
            If llCount = 0 Then
                slSql = "insert into VET_Vendor_Export (vetcode,vetattcode,vetVendorId,vetHasBeenSent,vetErrorMessage,vetExportMondayDate,vetProcessedDate,vetProcessedTime,vetEnteredDate,vetEnteredTime) values " & slAdd
                If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVet"
                    mInsertToVet = False
                    Erase slValues
                    Exit Function
                End If
            End If
    '        'we should only be here if no error.
            slSql = "update webvendorexport set senttoaff = 'Y' where code = " & slCode
            If mExecWebSQLWithRowsVendors(slSql) <> 1 Then
                blRet = False
            End If
        Else
            blRet = False
        End If  ' block if already sent
    End If
CLEANUP:
    Erase slValues
    mInsertToVet = blRet
    Exit Function
ErrHand:
    blRet = False
    slResult = "ouch"
    gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVet"
    GoTo CLEANUP
End Function
Private Function mInsertToVit(slData As String, slCode As String, slResult As String, slMondayDate As String) As Boolean
    'O: code, result (or hasbeen sent) monday export date
    '(1, 0, 21, -3,' IDDOESNOTEXIST','2010-06-28','2016-03-01',' S', 0, 3 )
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim c As Integer
    Dim slAdd As String
    Dim slTemp As String
    Dim slProcessedDate As String
    Dim slProcessedTime As String
    Dim slMessage As String
    Dim llCount As Long
    
    blRet = True
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 9 Then
        For c = 0 To 9
            slTemp = Trim$(Replace(slValues(c), """", ""))
            Select Case c
                Case ExportInfoIndex.Code, ExportInfoIndex.attCode, ExportInfoIndex.vendoridcode, ExportInfoIndex.result, ExportInfoIndex.spotOkcount, ExportInfoIndex.spotErrorcount
                    slAdd = slAdd & slTemp & ","
                    If c = ExportInfoIndex.Code Then
                        slCode = slTemp
                    ElseIf c = ExportInfoIndex.result Then
                        slResult = slTemp
                    End If
                Case ExportInfoIndex.Message
                     slAdd = slAdd & "'" & slTemp & "',"
                     slMessage = slTemp
                Case ExportInfoIndex.mondaydate, ExportInfoIndex.ProcessedDateTime
                    slAdd = slAdd & "'" & Format(slTemp, sgSQLDateForm) & "',"
                    If c = ExportInfoIndex.mondaydate Then
                        slMondayDate = Format(slTemp, sgSQLDateForm)
                    Else
                        slProcessedDate = Format(slTemp, sgSQLDateForm)
                        slProcessedTime = Format(slTemp, sgSQLTimeForm)
                        slAdd = slAdd & "'" & slProcessedTime & "',"
                    End If
                Case ExportInfoIndex.Service
                    slAdd = slAdd & "'" & slTemp & "',"
            End Select
        Next c
        slAdd = "(" & gLoseLastLetterIfComma(slAdd) & " )"
        'service, spot counts
        slSql = "update VIT_Vendor_Import set vitResult =" & slResult & ", vitMessage = '" & slMessage & "', vitProcessedDate = '" & slProcessedDate & "', vitprocessedtime = '" & slProcessedTime & "', vitImportMondayDate = '" & slMondayDate & "' where vitcode =" & slCode
        If gSQLAndReturn(slSql, False, llCount) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVit"
            mInsertToVit = False
            Erase slValues
            Exit Function
        End If
        If llCount = 0 Then
            'insert into VIT_Vendor_Import (vitcode,vitattcode,vitVendorId,vitResult,vitMessage,vitImportMondayDate,vitProcessedDate,vitProcessedTime,vitService,vitSpotCount,vitSpotErrorCount) values (1,0,21,-3,'IDDOESNOTEXIST','2010-06-28','2016-03-01','03:16:35','S',0,3 )
            slSql = "insert into VIT_Vendor_Import (vitcode,vitattcode,vitVendorId,vitResult,vitMessage,vitImportMondayDate,vitProcessedDate,vitProcessedTime,vitService,vitSpotCount,vitSpotErrorCount) values " & slAdd
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVit"
                mInsertToVit = False
                Erase slValues
                Exit Function
            End If
        End If
'        'we should only be here if no error.
        slSql = "update webImportArchive set senttoaff = 'Y' where code = " & slCode
        If mExecWebSQLWithRowsVendors(slSql) <> 1 Then
            blRet = False
        End If
    End If
CLEANUP:
    Erase slValues
    mInsertToVit = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "TrafficErrors.txt", "taskMonitor-mInsertToVit"
    GoTo CLEANUP
End Function

Private Function gExecWebSQLForVendor(aDataArray() As String, sSQL As String, blWantData As Boolean, Optional ilRetryMax As Integer = 0) As Long
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
    '7967 to traffic
'    If igDemoMode Then
'        gExecWebSQLForVendor = 0
'        Exit Function
'    End If
    gExecWebSQLForVendor = -1    ' -1 is an error condition.
    If Not gLoadOptionTrafficThenAffiliate("WebServer", "RootURL", slRootURL) Then
        gLogMsg "Error: gExecWebSQLForVendor: LoadOption RootURL Error", "TrafficErrors.Txt", False
'        gMsgBox "Error: gCheckWebSession: LoadOption RootURL Error"
        Exit Function
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOptionTrafficThenAffiliate("WebServer", "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To ilRetryMax
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
            If Not blWantData Then
                ' Caller does not want any data returned.
                gExecWebSQLForVendor = 0
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
                gExecWebSQLForVendor = 0
                Exit Function
            End If
        
            ' Each record we get back is a comma delimited string. In this case were only interested
            ' in the first record.
          '  aDataArray = Split(alRecordsArray(1), ",")
            aDataArray = alRecordsArray
            If Not IsArray(aDataArray) Then
                Exit Function
            End If
            gExecWebSQLForVendor = UBound(aDataArray)
            'If gExecWebSQLForVendor < 1 Then
            If gExecWebSQLForVendor > 1 Then
                ' The SQL Statement asked for two fields. DTStamp and PCName
                Exit Function
            End If
        End If
        If ilRetryMax > 0 Then
            Call gSleep(1)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "gExecWebSQLForVendor, " & ilRetryMax + 1 & " tries were exceeded. sSQL = " & sSQL & " llErrorCode = " & llErrorCode, "TrafficErrors.Txt", False
    Exit Function
    
ErrHandler:
    llErrorCode = err.Number
    Resume Next
End Function
Private Function gGetDataNoQuotes(sDataStr As String) As String
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
Private Function gUrlEncoding(sStr As String) As String

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
Private Function gSetURLPathEndSlash(ByVal slInPath As String, ilAdjDrivePath As Integer) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If right$(slPath, 1) <> "/" Then
        slPath = slPath + "/"
    End If
    If ilAdjDrivePath Then
        slPath = gAdjustDrivePath(slPath)
    End If
    gSetURLPathEndSlash = slPath
End Function
Public Function gVendorToWebAllowed(Optional tmcWebConnectIssue As Timer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slRootURL As String
    Dim llRet As Long
    
On Error GoTo ERRBOX
    blRet = False
    If gLoadOptionTrafficThenAffiliate("WebServer", "RootURL", slRootURL) Then
        If mWebVendorIsUsed() Then
            'doesn't exist will return 0!
            slSql = "Select COL_LENGTH('webVendors_Header','attCode') as amount"
            llRet = mExecWebSQLCount(slSql)
            If llRet > 0 Then
                blRet = True
            ElseIf llRet < 0 Then
                mVendorToWebFailed tmcWebConnectIssue
            Else
                gLogMsg "Call Counterpoint. Web is out of sync: web vendor tables/fields missing.", "TrafficErrors.Txt", False
            End If
        End If
    End If
CLEANUP:
    bgVendorToWebAllowed = blRet
    gVendorToWebAllowed = blRet
    Exit Function
ERRBOX:
    blRet = False
    GoTo CLEANUP
End Function
Private Function mWebVendorIsUsed() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim rst As ADODB.Recordset
On Error GoTo ERRBOX
    blRet = False
    slSql = "select wvtVendorId from WVT_Vendor_Table " & mGetWebVendorSqlWhere()
    'Set rst = cnn.Execute(slSQL)
    Set rst = gSQLSelectCall(slSql)
    If Not rst.EOF Then
        blRet = True
    End If
    mWebVendorIsUsed = blRet
    Exit Function
ERRBOX:
    mWebVendorIsUsed = False
End Function
Private Function mGetWebVendorSqlWhere() As String
    'So getting the web vendors remains consistent
    'test appropriate username pasword (and address if needed)
    'if using for vat, add wvt to From and add AND (wvtVendorID = vatWvtVendorID)
    Dim slSql As String
    
    slSql = " where wvtApprovalPassword <> '' AND (((WvtExportmethod = 1 AND WvtStationuserName <> '' AND WvtStationPassword <> '') OR (wvtExportMethod = 2 AND WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '') OR (wvtExportMethod = 3 And WvtStationuserName <> '' AND WvtStationPassword <> '' AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '')) "
    slSql = slSql & " OR  ((WvtImportmethod = 1 AND WvtStationuserName <> '' AND WvtStationPassword <> '') OR (wvtImportMethod = 2 AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '') OR (wvtImportMethod = 3 And WvtStationuserName <> '' AND WvtStationPassword <> '' AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '')))  "
    
    mGetWebVendorSqlWhere = slSql
End Function
Private Sub mVendorToWebFailed(tmcWebConnectIssue As Timer)
    bgVendorToWebAllowed = False
    If Not tmcWebConnectIssue Is Nothing Then
       tmcWebConnectIssue.Enabled = True
    End If
End Sub
Private Function mAlertRunByOther() As String
    'O: G E Y or X, meaning need to run
    Dim blRet As Boolean
    Dim slRet As String
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim dlDate As Date
    Dim ilAlert As Integer
    
    blRet = False
    slRet = "X"
On Error GoTo ERRBOX
'test will get first and then will not after the # of the last number
'igWVImportElapsed = igWVImportElapsed + (igWVImportElapsed * 1)
    dlDate = DateAdd("n", -igWVImportElapsed, dgWvImportLast)
'reset test
'igWVImportElapsed = igWVImportElapsed - (igWVImportElapsed * 1)
    slSql = "Select count(*) as amount from auf_alert_user where auftype = 'V' and aufstatus = 'C' and aufcleardate >= '" & Format(dlDate, sgSQLDateForm) & "' AND aufcleartime > '" & Format(dlDate, sgSQLTimeForm) & "'"
    'Set myRst = cnn.Execute(slSQL)
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        If myRst!amount > 0 Then
            slSql = "select aufcefcode from auf_alert_user where auftype = 'V' and aufstatus = 'R'"
            'Set myRst = cnn.Execute(slSQL)
            Set myRst = gSQLSelectCall(slSql)
            If Not myRst.EOF Then
                Do While Not myRst.EOF
                    ilAlert = myRst!aufcefcode
                    If mParseMonitorIssueIsYellow(ilAlert) Then
                        slRet = "Y"
                        Exit Do
                    Else
                        slRet = "E"
                    End If
                    myRst.MoveNext
                Loop
            Else
                slRet = "G"
            End If
        Else
            slSql = "Select count(*) as amount from auf_alert_user where auftype = 'V' and aufstatus = 'R' and aufEntereddate >= '" & Format(dlDate, sgSQLDateForm) & "' AND aufEnteredtime > '" & Format(dlDate, sgSQLTimeForm) & "'"
            'Set myRst = cnn.Execute(slSQL)
            Set myRst = gSQLSelectCall(slSql)
            'not find?  then must run.  but slret already = X
            If Not myRst.EOF Then
                If myRst!amount > 0 Then
                    slSql = "select aufcefcode from auf_alert_user where auftype = 'V' and aufstatus = 'R'"
                    'Set myRst = cnn.Execute(slSQL)
                    Set myRst = gSQLSelectCall(slSql)
                    If Not myRst.EOF Then
                        Do While Not myRst.EOF
                            ilAlert = myRst!aufcefcode
                            If mParseMonitorIssueIsYellow(ilAlert) Then
                                slRet = "Y"
                                Exit Do
                            Else
                                slRet = "E"
                            End If
                            myRst.MoveNext
                        Loop
                    Else
                        slRet = "G"
                    End If
                End If
            End If
        End If
    End If
CLEANUP:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAlertRunByOther = slRet
    Exit Function
ERRBOX:
    slRet = "X"
    GoTo CLEANUP
End Function
Private Function mParseMonitorIssueIsYellow(ilAlert As Integer) As Boolean
    Dim blRet As Boolean
    
    blRet = False
    'none means regular alert issue--red!
    If ilAlert <> VendorWvmAlert.None Then
        Select Case ilAlert
            'these are yellow
            Case VendorWvmAlert.ConnectionIssue, VendorWvmAlert.ExportRunning, VendorWvmAlert.WebVendorSent
                blRet = True
        End Select
    End If
    mParseMonitorIssueIsYellow = blRet
End Function


Private Function mExecWebSQLWithRowsVendors(sSQL As String) As Long

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
'    If igDemoMode Then
'        mExecWebSQLWithRowsVendors = 0
'        Exit Function
'    End If

    mExecWebSQLWithRowsVendors = -1    ' -1 is an error condition.

    'D.S. 02/03/09
    'First make sure that we can establish a connection
    ilMaxRetries = 4
    ilStepNo = 2

' Dan only here if bgVendorAllowed
'    ilRet = gHasWebAccess()
'    ilStepNo = 3
'    If Not ilRet Then
'        For ilRetries = 0 To ilMaxRetries Step 1
'            Call Sleep(3000)  'Sleep 3 seconds
'            ilRet = gHasWebAccess()
'            If ilRet = True Then
'                'We got access
'                Exit For
'            End If
'        Next ilRetries
'
'        If Not ilRet And ilRetries = ilMaxRetries Then
'            gLogMsg "Error: mExecWebSQLWithRowsVendors was called and gHasWebAccess returned False", "AffWebErrorLog.txt", False
'            Exit Function
'        End If
'
'    End If
    ilStepNo = 4
    If Not gLoadOptionTrafficThenAffiliate("WebServer", "RootURL", slRootURL) Then
        gLogMsg "Error: mExecWebSQLWithRowsVendors: LoadOption RootURL Error", "AffWebErrorLog.Txt", False
        gMsgBox "Error: mExecWebSQLWithRowsVendors: LoadOption RootURL Error"
        Exit Function
    End If
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)

    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.

    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOptionTrafficThenAffiliate("WebServer", "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If

    For ilRetries = 0 To 14
        ilStepNo = 4
        llErrorCode = 0
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(sSQL)
            If Not Left(slResponse, 5) = "ERROR" Then
                llReturn = 200
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            ilStepNo = 5
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            ilStepNo = 6
            objXMLHTTP.Send
            ilStepNo = 7
            llReturn = objXMLHTTP.Status
            ilStepNo = 8
            slResponse = Trim(objXMLHTTP.responseText)
            ilStepNo = 9
            Set objXMLHTTP = Nothing
        End If
        If Left(slResponse, 5) = "ERROR" Then
            gLogMsg "Error: mExecWebSQLWithRowsVendors - " & slResponse, "AffWebErrorLog.Txt", False
            Exit Function
        End If

        If llReturn = 200 Then
            ilIdx = 1
            llReturn = 0
            While ilIdx < Len(slResponse) - 1 And Mid(slResponse, ilIdx, 1) <> " "
                ilIdx = ilIdx + 1
            Wend
            llReturn = Left(slResponse, ilIdx)
            mExecWebSQLWithRowsVendors = llReturn
            Exit Function
        Else
            'gLogMsg "mExecWebSQLWithRowsVendors is retrying", "WebSubsRetryLog.Txt", False
            Call gSleep(2)  ' Delay for 2 seconds between requests.
        End If
    Next

    ' We were never successful if we make it to here.
    gLogMsg "mExecWebSQLWithRowsVendors, retries were exceeded. slResponse = " & slResponse & " StepNo = " & ilStepNo, "AffWebErrorLog.txt", False

    Exit Function

ErrHandler:
    'llErrorCode = Err.Number
    'gMsg = "A general error has occured in modWebSubs-mExecWebSQLWithRowsVendors: Retries = " & ilRetries & " Step Num = " & ilStepNo & " slResponse = " & slResponse & " llReturn = " & llReturn
    'gLogMsg gMsg & Err.Description & " Error #" & Err.Number, "AffWebErrorLog.txt", False
    'gLogMsg "    SQLQuery = " & sSQL, "AffWebErrorLog.txt", False
    Resume Next
End Function
Public Function gLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    gLoseLastLetterIfComma = slNewString
End Function
'Public Function gHasWebAccess() As Boolean
'    gHasWebAccess = False
'    If Not gUsingWeb Then
'        ' This site does not have the Using Web Server turned on in the Site Options.
'        Exit Function
'    End If
'    If Not gWebAccessTestedOk Then
'        ' The function gTestAccessToWebServer was not able to access the web server and therfore
'        ' this session should not attempt to make any web calls.
'        Exit Function
'    End If
'    gHasWebAccess = True
'End Function
Private Function mQueryServiceController(myServiceFacts() As ServiceControllerInfo) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slData() As String
    Dim slValues() As String
    Dim llCount As Long
    Dim c As Integer
    Dim ilCurrent As Integer
    'service controller is always 3 records
    ilCurrent = 0
    blRet = True
    slSql = "Select Mode,GenerateDebug,GenerateFile,ImportLast,ImportFiles,isRunning,ImportSpan from vendorservicecontroller"
    llCount = gExecWebSQLForVendor(slData, slSql, True)
    If llCount > 0 Then
        For c = 1 To llCount - 1
            slValues = Split(slData(c), ",")
            If UBound(slValues) >= 6 Then
                ilCurrent = c - 1
                If ilCurrent < 3 Then
                    With myServiceFacts(ilCurrent)
                        .Mode = Trim$(Replace(slValues(0), """", ""))
                        If Trim$(Replace(slValues(1), """", "")) = "Y" Then
                            .GenerateDebug = True
                        End If
                        If Trim$(Replace(slValues(2), """", "")) = "Y" Then
                            .GenerateFile = True
                        End If
                        .ImportLast = Trim$(Replace(slValues(3), """", ""))
                        .ImportFiles = Trim$(Replace(slValues(4), """", ""))
                        If Trim$(Replace(slValues(5), """", "")) = "Y" Then
                            .IsRunning = True
                        End If
                         .ImportSpan = Trim$(Replace(slValues(6), """", ""))
                    End With
                End If
            End If
        Next c
    ElseIf llCount < 0 Then
        blRet = False
    End If
    Erase slData
    Erase slValues
    mQueryServiceController = blRet
End Function

Public Function gGetUniqueFileName() As String
    
    Dim sUKey As String
    
    'Resolution appears to be about 3 microseconds in a tight loop
    sUKey = Replace(CStr(gStopWatch.ElapsedSeconds), ".", "_")
    gGetUniqueFileName = CurDir & "\TempFile_" & sgUserName & "_" & sUKey & ".txt"

End Function

