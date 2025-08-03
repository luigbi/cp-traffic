Attribute VB_Name = "modExport"
'******************************************************
'*  modExport - various global declarations for importing
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'
Public lgLastServiceDate As Long
Public lgLastServiceTime As Long
Public igCountTimeNotChanged As Long


Public sgDSN As String

'Export parameters
Public igExportSource As Integer    '1=From Custom button on Export screen; 2=From Export Queue
Public igExportTypeNumber As Integer
Public sgExportTypeChar As String
Public sgWebExport As String    'W=CSI Web; C=Cumulus; B=Both
Public sgExportName As String
Public lgExportEhtCode As Long
Public lgExportEhtInfoIndex As Long
Public igExportReturn As Integer   '0=Cancelled; 1=Ok; 2=Error
Public sgExporStartDate As String
Public sgExportEndDate As String
Public igExportDays As String
Public lgExportEqtCode As Long
Public sgExportResultName As String

Public igVehicleSpecChgFlag As Integer   'Set in VehicleSpec

Public sgXDSSection() As String
'9256 moved here
Public Const STATIONXMLRECEIVERID As String = "ReceiverIDSource"

Type EHTSTDCOLOR
    lEhtCode As Long
    sLogStatus As String * 1
    sCopyStatus As String * 1
    sGenFont As String * 1
    sGen As String * 1
End Type

Type EHTINFO
    lEhtCode As Long
    lFirstEvt As Long
    lFirstEct As Long
    lStdEhtCode As Long
    iRefRowNo As Integer
    blRemoved As Boolean
    sLogStatus As String * 1
    sCopyStatus As String * 1
    sGenFont As String * 1
    sGen As String * 3
End Type
Public tgEhtInfo() As EHTINFO

Type EVTINFO
    iVefCode As Integer
    lNextEvt As Long
End Type
Public tgEvtInfo() As EVTINFO

Type ECTINFO
    sLogType As String * 1
    sFieldType As String * 1
    sFieldName As String * 30
    lFieldValue As Long
    sFieldString As String * 250
    lNextEct As Long
End Type
Public tgEctInfo() As ECTINFO

Type SPECINFO
    sName As String * 10
    sType As String * 1
    sFullName As String * 30
    sCheckDateSpan As String * 1
End Type

Public tgSpecInfo() As SPECINFO


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
    'Remove " and '
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
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
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
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        'ilPos = InStr(1, slName, """", 1)
        'If ilPos > 0 Then
        '    Mid$(slName, ilPos, 1) = "'"
        '    ilFound = True
        'End If
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
'

Private Sub gExportVefCode(ilVefCode As Integer, ilToBeProcessed As Integer, ilBeenProcessed As Integer)
    On Error GoTo ErrHand
    SQLQuery = "UPDATE eqt_Export_Queue SET "
    SQLQuery = SQLQuery & "eqtProcesingVefCode = " & ilVefCode & ", "
    SQLQuery = SQLQuery & "eqtToBeProcessed = " & ilToBeProcessed & ", "
    SQLQuery = SQLQuery & "eqtBeenProcessed = " & ilBeenProcessed & " "
    SQLQuery = SQLQuery & "WHERE eqtCode = " & lgExportEqtCode
    'cnn.Execute slSQL_AlertClear, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "ExportQueueLog.txt", "modExport-gExportVefCode"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-gExportVefCode"
    Exit Sub
'ErrHand1:
'    gHandleError "ExportQueueLog.txt", "ExportQueue-gExportVefCode"
'    Return
End Sub

Public Function gCustomStartStatus(slExportType As String, slExportName As String, slEqtType As String, slStartDate As String, slCycle As String, ilVefCode() As Integer, ilShttCode() As Integer) As Long
    Dim ilRet As Integer
    Dim llEhtCode As Long
    Dim llEctCode As Long
    Dim llEvtCode As Long
    Dim llEstCode As Long
    Dim llEqtCode As Long
    Dim Ctrl As control
    Dim blVisible As Boolean
    Dim blSetValue As Boolean
    Dim slIndex As String
    Dim slControlName As String
    Dim slFieldType As String
    Dim slFieldValue As String
    Dim llFieldValue As Long
    Dim ilNoVehicles As Integer
    Dim ilCycle As Integer
    Dim llLoop As Integer
    Dim llEct As Long
    Dim llEht As Integer

    If igExportSource = 2 Then
        gCustomStartStatus = -1
        Exit Function
    End If
    ReDim tgEhtInfo(0 To 1) As EHTINFO
    ReDim tgEvtInfo(0 To 0) As EVTINFO
    ReDim tgEctInfo(0 To 0) As ECTINFO
    lgExportEhtInfoIndex = 0
    tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1

    llEhtCode = mAddEht(slExportType, slExportName)
    If llEhtCode > 0 Then

        llEht = lgExportEhtInfoIndex
        llEct = tgEhtInfo(llEht).lFirstEct
        Do While llEct <> -1
            llEctCode = mAddEct(llEhtCode, tgEctInfo(llEct).sLogType, tgEctInfo(llEct).sFieldType, tgEctInfo(llEct).sFieldName, tgEctInfo(llEct).lFieldValue, tgEctInfo(llEct).sFieldString)
            llEct = tgEctInfo(llEct).lNextEct
        Loop

        ilNoVehicles = 0
        ilCycle = Val(slCycle)
        For llLoop = 0 To UBound(ilVefCode) - 1 Step 1
            llEvtCode = mAddEvt(llEhtCode, ilVefCode(llLoop))
            If llEvtCode > 0 Then
                ilNoVehicles = ilNoVehicles + 1
            End If
        Next llLoop
        If ilNoVehicles = 1 Then
            For llLoop = 0 To UBound(ilShttCode) - 1 Step 1
                llEstCode = mAddEst(llEhtCode, llEvtCode, ilShttCode(llLoop))
            Next llLoop
        End If
        llEqtCode = mAddEqt(llEhtCode, slEqtType, ilNoVehicles, slStartDate, ilCycle)
        gCustomStartStatus = llEqtCode
    Else
        gCustomStartStatus = 0
    End If
    Exit Function
IndexErr:
    ilRet = 1
    Resume Next
mSaveCtrlErr:
    blVisible = False
    Resume Next
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-gCustomStartStatus"
    gCustomStartStatus = 0

End Function

Private Function mAddEht(slExportType As String, slExportName As String) As Long
    Dim llEhtCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into eht_Export_Header ( "
    SQLQuery = SQLQuery & "ehtCode, "
    SQLQuery = SQLQuery & "ehtExportType, "
    SQLQuery = SQLQuery & "ehtSubType, "
    SQLQuery = SQLQuery & "ehtStandardEhtCode, "
    SQLQuery = SQLQuery & "ehtExportName, "
    SQLQuery = SQLQuery & "ehtUstCode, "
    SQLQuery = SQLQuery & "ehtLDE, "
    SQLQuery = SQLQuery & "ehtLeadTime, "
    SQLQuery = SQLQuery & "ehtCycle, "
    SQLQuery = SQLQuery & "ehtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slExportType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("C") & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slExportName) & "', "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEhtCode = gInsertAndReturnCode(SQLQuery, "eht_Export_Header", "ehtCode", "Replace")
    mAddEht = llEhtCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-mAddEht"
    mAddEht = 0
End Function


Private Function mAddEct(llEhtCode As Long, slLogType As String, slFieldType As String, slFieldName As String, llFieldValue As Long, slFieldString As String) As Long
    Dim llEctCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into ect_Export_Criteria ( "
    SQLQuery = SQLQuery & "ectCode, "
    SQLQuery = SQLQuery & "ectEhtCode, "
    SQLQuery = SQLQuery & "ectLogType, "
    SQLQuery = SQLQuery & "ectFieldType, "
    SQLQuery = SQLQuery & "ectFieldName, "
    SQLQuery = SQLQuery & "ectFieldValue, "
    SQLQuery = SQLQuery & "ectFieldString, "
    SQLQuery = SQLQuery & "ectUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slLogType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldName) & "', "
    SQLQuery = SQLQuery & llFieldValue & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slFieldString) & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    
    llEctCode = gInsertAndReturnCode(SQLQuery, "ect_Export_Criteria", "ectCode", "Replace")

    mAddEct = llEctCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-mAddEct"
    mAddEct = 0

End Function

Private Function mAddEvt(llEhtCode As Long, ilVefCode As Integer) As Long
    Dim llEvtCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into evt_Export_Vehicles ( "
    SQLQuery = SQLQuery & "evtCode, "
    SQLQuery = SQLQuery & "evtEhtCode, "
    SQLQuery = SQLQuery & "evtVefCode, "
    SQLQuery = SQLQuery & "evtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEvtCode = gInsertAndReturnCode(SQLQuery, "evt_Export_Vehicles", "evtCode", "Replace")
    mAddEvt = llEvtCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-mAddEvt"
    mAddEvt = 0

End Function

Private Function mAddEst(llEhtCode As Long, llEvtCode As Long, ilShttCode As Integer) As Long
    Dim llEstCode As Long
    
    On Error GoTo ErrHand
    SQLQuery = "Insert Into est_Export_Station ( "
    SQLQuery = SQLQuery & "estCode, "
    SQLQuery = SQLQuery & "estEhtCode, "
    SQLQuery = SQLQuery & "estEvtCode, "
    SQLQuery = SQLQuery & "estShttCode, "
    SQLQuery = SQLQuery & "estUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & llEvtCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEstCode = gInsertAndReturnCode(SQLQuery, "est_Export_Station", "estCode", "Replace")
    mAddEst = llEstCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-mAddEst"
    mAddEst = 0

End Function



Public Function gCustomEndStatus(llEqtCode As Long, ilReturn As Integer, slResultName As String) As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    
    On Error GoTo ErrHand
    If igExportSource = 2 Then
        gCustomEndStatus = True
        Exit Function
    End If
    If llEqtCode <= 0 Then
        gCustomEndStatus = True
        Exit Function
    End If
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
    SQLQuery = "UPDATE eqt_Export_Queue SET "
    If ilReturn = 2 Then
        SQLQuery = SQLQuery & "eqtStatus = 'E'" & ", "
    Else
        SQLQuery = SQLQuery & "eqtStatus = 'C'" & ", "
    End If
    SQLQuery = SQLQuery & "eqtResultFile = '" & slResultName & "',"
    SQLQuery = SQLQuery & "eqtDateCompleted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "eqtTimeCompleted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
    SQLQuery = SQLQuery & "WHERE eqtCode = " & llEqtCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffErrorLog.txt", "modExport-gCustomEndStatus"
        gCustomEndStatus = False
        Exit Function
    End If
    llEqtCode = -1
    gCustomEndStatus = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Update-gCustomEndStatus"
    gCustomEndStatus = False
    Exit Function
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Custom Update-gCustomEndStatus"
'    gCustomEndStatus = False
'    Exit Function
End Function

Private Function mAddEqt(llEhtCode As Long, slExportType As String, ilNoVehicles As Integer, slStartDate As String, ilCycle As Integer) As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llEqtCode As Long
    
    On Error GoTo ErrHand
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
    SQLQuery = "Insert Into eqt_Export_Queue ( "
    SQLQuery = SQLQuery & "eqtCode, "
    SQLQuery = SQLQuery & "eqtEhtCode, "
    SQLQuery = SQLQuery & "eqtPriority, "
    SQLQuery = SQLQuery & "eqtDateEntered, "
    SQLQuery = SQLQuery & "eqtTimeEntered, "
    SQLQuery = SQLQuery & "eqtStatus, "
    SQLQuery = SQLQuery & "eqtDateStarted, "
    SQLQuery = SQLQuery & "eqtTimeStarted, "
    SQLQuery = SQLQuery & "eqtDateCompleted, "
    SQLQuery = SQLQuery & "eqtTimeCompleted, "
    SQLQuery = SQLQuery & "eqtUstCode, "
    SQLQuery = SQLQuery & "eqtResultFile, "
    SQLQuery = SQLQuery & "eqtType, "
    SQLQuery = SQLQuery & "eqtStartDate, "
    SQLQuery = SQLQuery & "eqtNumberDays, "
    SQLQuery = SQLQuery & "eqtEndDate, "
    SQLQuery = SQLQuery & "eqtProcesingVefCode, "
    SQLQuery = SQLQuery & "eqtToBeProcessed, "
    SQLQuery = SQLQuery & "eqtBeenProcessed, "
    SQLQuery = SQLQuery & "eqtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & llEhtCode & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & Format$(slNowDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slNowTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & "P" & "', "
    SQLQuery = SQLQuery & "'" & Format$(slNowDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slNowTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & slExportType & "', "
    SQLQuery = SQLQuery & "'" & Format$(slStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & ilCycle & ", "
    SQLQuery = SQLQuery & "'" & Format$(DateAdd("d", ilCycle - 1, slStartDate), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & ilNoVehicles & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llEqtCode = gInsertAndReturnCode(SQLQuery, "eqt_export_queue", "eqtCode", "Replace")
    mAddEqt = llEqtCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Custom Insert Status-mAddEqt"
    mAddEqt = 0

End Function

Public Sub gEraseEventDate()
    'Purpose:  Remove events that have had a date change from the affiliate and web system
    '          ECF contains the events that have been changed
    Dim rst_ecf As ADODB.Recordset
    Dim rst_att As ADODB.Recordset
    Dim rst_event As ADODB.Recordset
    Dim rst_Ast As ADODB.Recordset
    Dim rst_Lst As ADODB.Recordset
    Dim llAttCount As Long
    Dim llLstCount As Long
    Dim slEventDate As String
    Dim slMoDate As String
    Dim slSuDate As String
    Dim blRetainCPTT As Boolean
    Dim llRet As Long
    Dim slStr As String
    Dim slQuery As String
    
    On Error GoTo ErrHand
    
     SQLQuery = "select * from ecf_Event_Chg_Date" & _
    " WHERE ecfWebCleared <> 'Y' " & _
    " OR ecfMarketronCleared <> 'Y' " & _
    " OR ecfUnivisionCleared <> 'Y' " & _
    " ORDER BY ecfGsfCode, ecfEnteredDate, ecfEnteredTime"
    Set rst_ecf = gSQLSelectCall(SQLQuery)
    Do While Not rst_ecf.EOF
        llAttCount = 0
        llLstCount = 0
        slEventDate = Format(rst_ecf!ecfFromDate, sgShowDateForm)
        slMoDate = gObtainPrevMonday(slEventDate)
        slSuDate = DateAdd("d", 6, slMoDate)
        SQLQuery = "SELECT vefName, vefCode, gsfGameNo, ghfSeasonName FROM gsf_Game_Schd LEFT OUTER JOIN vef_Vehicles ON gsfVefCode = vefCode " & _
        "LEFT OUTER JOIN ghf_Game_Header ON gsfGhfCode = ghfCode " & _
        "WHERE gsfCode = " & rst_ecf!ecfGsfCode
        Set rst_event = gSQLSelectCall(SQLQuery)
        If Not rst_event.EOF Then
            gLogMsg "Deleting ALL Event Spots for : " & rst_event!vefName & " Season " & rst_event!ghfSeasonName & " Event # " & rst_event!gsfGameNo, "ClearEvents.Txt", False
            
            SQLQuery = "Select COUNT(lstCode) from LST"
            SQLQuery = SQLQuery + " WHERE"
            SQLQuery = SQLQuery + " lstLogVefCode = " & rst_event!vefCode
            SQLQuery = SQLQuery + " AND lstLogDate >= " & "'" & Format(slMoDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery + " AND lstLogDate <= " & "'" & Format(slSuDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery + " AND lstGsfCode <> " & rst_ecf!ecfGsfCode
            Set rst = gSQLSelectCall(SQLQuery)
            'If rst.EOF Then
            If rst(0).Value = 0 Then
                blRetainCPTT = False
            Else
                blRetainCPTT = True
            End If
            
            'Remove spots from ast, lst, and web spots
            SQLQuery = "SELECT attCode FROM att WHERE attVefCode = " & rst_event!vefCode
            Set rst_att = gSQLSelectCall(SQLQuery)
            Do While Not rst_att.EOF
                'Remove spots from Web
                If rst_ecf!ecfWebCleared <> "Y" Then
                    slStr = "Delete From Spots Where attCode = " & rst_att!attCode
                    slStr = slStr & " AND gsfCode = " & rst_ecf!ecfGsfCode
                    slStr = slStr & " AND FeedDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    llRet = gExecWebSQLWithRowsEffected(slStr)
                    slStr = "Delete From SpotRevisions Where attCode = " & rst_att!attCode
                    slStr = slStr & " AND gsfCode = " & rst_ecf!ecfGsfCode
                    slStr = slStr & " AND FeedDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    llRet = gExecWebSQLWithRowsEffected(slStr)
                    'D.S. 02/05/13 Added call to delete from the Event Info table
                    slStr = "Delete From GameInfo Where attCode = " & rst_att!attCode
                    slStr = slStr & " AND Code = " & rst_ecf!ecfGsfCode
                    llRet = gExecWebSQLWithRowsEffected(slStr)
                End If

                'Dan: Remove spots from Marketron
                If rst_ecf!ecfMarketronCleared <> "Y" Then
                
                End If
                
                'Dan: Remove spots from Univision
                If rst_ecf!ecfUnivisionCleared <> "Y" Then
                
                End If
                
                'Remove AST if exist
                SQLQuery = "Select COUNT(astCode) from AST"
                SQLQuery = SQLQuery + " LEFT OUTER JOIN lst On astlsfCode = lstCode"
                SQLQuery = SQLQuery + " WHERE"
                SQLQuery = SQLQuery + " astAtfCode = " & rst_att!attCode
                SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_ecf!ecfGsfCode
                SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                Set rst = gSQLSelectCall(SQLQuery)
                
                'If Not rst.EOF Then
                If rst(0).Value <> 0 Then
                    llAttCount = llAttCount + rst(0).Value
                    'D.S. 02/12/13 - I don't think you can use a left outer join on a delete statement
                    'SQLQuery = "DELETE FROM ast"
                    'SQLQuery = SQLQuery + " LEFT OUTER JOIN lst On astlsfCode = lstCode"
                    'SQLQuery = SQLQuery + " WHERE"
                    'SQLQuery = SQLQuery + " astAtfCode = " & rst_att!attCode
                    'SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_ecf!ecfGsfCode
                    'SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    'SQLQuery = "DELETE FROM ast where astCode in"
                    'SQLQuery = SQLQuery + " (Select astCode from ast LEFT OUTER JOIN lst On astlsfCode = lstCode"
                    'SQLQuery = SQLQuery + " AND astAtfCode = " & rst_att!attCode
                    'SQLQuery = SQLQuery + " AND astFeedDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    'SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_ecf!ecfGsfCode
                    'SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "')"
                    
                    SQLQuery = "Select astCode from ast LEFT OUTER JOIN lst On astlsfCode = lstCode"
                    SQLQuery = SQLQuery + " WHERE astAtfCode = " & rst_att!attCode
                    SQLQuery = SQLQuery + " AND astFeedDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery + " AND lstGsfCode = " & rst_ecf!ecfGsfCode
                    SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                    Set rst_Ast = gSQLSelectCall(SQLQuery)
                    Do While Not rst_Ast.EOF
                        SQLQuery = "DELETE from ast where astCode = " & rst_Ast!astCode
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand1:
                            gHandleError "AffErrorLog.txt", "ModExport-gEraseEventDate"
                            Exit Sub
                        End If
                        rst_Ast.MoveNext
                    Loop
                End If
                
                'Remove CPTT if no lst exist for the week exspect for the event that was changed
                If Not blRetainCPTT Then
                    SQLQuery = "DELETE FROM cptt"
                    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & rst_att!attCode
                    SQLQuery = SQLQuery + " AND cpttStartdate = " & "'" & Format(slMoDate, sgSQLDateForm) & "'"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand1:
                        gHandleError "AffErrorLog.txt", "ModExport-gEraseEventDate"
                        Exit Sub
                    End If
                End If
                
                rst_att.MoveNext
            Loop
            SQLQuery = "Select COUNT(lstCode) from LST"
            SQLQuery = SQLQuery + " WHERE"
            SQLQuery = SQLQuery + " lstGsfCode = " & rst_ecf!ecfGsfCode
            'd.s. 02/05/13 added date below
            SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
            Set rst = gSQLSelectCall(SQLQuery)
            
            'If Not rst.EOF Then
            If rst(0).Value > 0 Then
                llLstCount = llLstCount + rst(0).Value
                SQLQuery = "DELETE FROM lst"
                SQLQuery = SQLQuery + " WHERE lstGsfCode = " & rst_ecf!ecfGsfCode
                SQLQuery = SQLQuery + " AND lstLogDate = " & "'" & Format(slEventDate, sgSQLDateForm) & "'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gHandleError "AffErrorLog.txt", "ModExport-gEraseEventDate"
                    Exit Sub
                End If
            End If
            gLogMsg "Deleted " & llAttCount & " Affiliate Spots for : " & rst_event!vefName & " Season " & rst_event!ghfSeasonName & " Event # " & rst_event!gsfGameNo, "ClearEvents.Txt", False
            gLogMsg "Deleted " & llLstCount & " Log Spots for : " & rst_event!vefName & " Season " & rst_event!ghfSeasonName & " Event # " & rst_event!gsfGameNo, "ClearEvents.Txt", False
        End If
        SQLQuery = "UPDATE ecf_Event_Chg_Date SET "
        SQLQuery = SQLQuery & "ecfWebCleared = 'Y'" & ","
        SQLQuery = SQLQuery & "ecfMarketronCleared = 'Y'" & ","
        SQLQuery = SQLQuery & "ecfUnivisionCleared = 'Y'"
        SQLQuery = SQLQuery & "WHERE ecfCode = " & rst_ecf!ecfCode
        'cnn.Execute slSQL_AlertClear, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "AffErrorLog.txt", "ModExport-gEraseEventDate"
            Exit Sub
        End If
        
        rst_ecf.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gEraseEventDate"
    Exit Sub
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "gEraseEventDate"
'    Return
End Sub
'6082
Public Function gAstFileWrite(frmForm As Form, rsAstFiles As ADODB.Recordset, smDate As String, smExportDirectory As String, bmAstFileError As Boolean) As Boolean
    'true means error. False, all ok
    'bmAstFileError if true, problem writing
    Dim blRet As Boolean
    Dim slProblemFiles As String
    Dim blWritten As Boolean
    '6680 move here
    Dim olDmas As Dictionary
    
    blWritten = False
    slProblemFiles = ""
    blRet = False
    Set olDmas = New Dictionary
On Error GoTo ERRORBOX
    If Not rsAstFiles Is Nothing Then
        With rsAstFiles
            If (.State And adStateOpen) <> 0 Then
                If .RecordCount > 0 Then
                    .Sort = "Contract"
                    If Not mAstFileGetContracts(rsAstFiles) Then
                        bmAstFileError = True
                        GoTo Cleanup
                    End If
                    '7161
                    .Sort = "FeedDate"
                    If Not mAstFileWriteUnits(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Units"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWriteStations(rsAstFiles, smExportDirectory, olDmas) Then
                        blRet = True
                        slProblemFiles = "Stations"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    '6680 write dmas here instead of inside 'WriteStations' above
                    If Not mAstFileWriteDmas(rsAstFiles, smExportDirectory, olDmas) Then
                        blRet = True
                        slProblemFiles = "Dmas"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWriteContracts(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Contracts"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWriteAgencies(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Agencies"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWriteAdvertisers(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Advertisters"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWriteISCI(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Iscis"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If
                    If Not mAstFileWritePrograms(rsAstFiles, smExportDirectory) Then
                        blRet = True
                        slProblemFiles = "Programs"
                        GoTo Cleanup
                    Else
                        blWritten = True
                    End If

               End If
            Else
                blRet = True
            End If
        End With
    End If
Cleanup:
    If blRet Then
        frmForm.mSetResults "Problem with Transparency Files:  Could not write " & slProblemFiles, RGB(255, 0, 0)
    ElseIf bmAstFileError Then
        blRet = True
        frmForm.mSetResults "Problem with Transparency Files: Some spots couldn't be collected.", RGB(255, 0, 0)
    ElseIf blWritten Then
        frmForm.mSetResults "Transparency Files created", RGB(0, 0, 255)
    End If
    Set olDmas = Nothing
    gAstFileWrite = blRet
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileWrite.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = True
    GoTo Cleanup
End Function
Private Function mAstFileGetContracts(myRst As ADODB.Recordset) As Boolean
    Dim blRet As Boolean
    Dim llPreviousContract As Long
    Dim slAdv As String
    Dim slAgy As String
    Dim slBuyer As String
    Dim llContractId As Long
    Dim slStart As String
    Dim slEnd As String
    Dim llIndex As Long
    Dim ilAgy As Integer
    Dim ilAdv As Integer
    
On Error GoTo ERRORBOX
    blRet = True
    slAdv = ""
    slAgy = ""
    slBuyer = ""
    slStart = ""
    slEnd = ""
    llContractId = 0
    ilAgy = 0
    ilAdv = 0
    myRst.MoveFirst
    Do While Not myRst.EOF
        If myRst!Contract <> 0 Then
            If myRst!Contract <> llPreviousContract Then
                llPreviousContract = myRst!Contract
                mAstFileContractInfo llPreviousContract, slBuyer, llContractId, slStart, slEnd, ilAgy
                If myRst!advid > 0 Then
                    llIndex = gBinarySearchAdf(myRst!advid)
                    If llIndex <> -1 Then
                        slAdv = Left(Trim$(tgAdvtInfo(llIndex).sAdvtName), 30)
                    End If
                Else
                    slAdv = ""
                End If
                'got an agency from above?
                If ilAgy > 0 Then
                    llIndex = gBinarySearchAgency(CLng(ilAgy))
                    If llIndex <> -1 Then
                        slAgy = Left(Trim$(tgAgencyInfo(llIndex).sAgencyName), 30)
                    End If
                Else
                    slAgy = ""
                    ilAgy = 0
                End If
            'else?  use values from previous mAstFileContractInfo!
            End If
        'blackout
        Else
            slStart = ""
            slEnd = ""
            slBuyer = ""
            ilAgy = 0
            slAgy = ""
            llContractId = 0
            'don't return buyer:  espn doesn't store this info with adv.
            If myRst!advid > 0 Then
                llIndex = gBinarySearchAdf(myRst!advid)
                If llIndex <> -1 Then
                    slAdv = Left(Trim$(tgAdvtInfo(llIndex).sAdvtName), 30)
                End If
            Else
                slAdv = ""
            End If
        End If
        myRst!agencyid = ilAgy
        myRst!Agency = slAgy
        myRst!Client = slAdv
        myRst!Buyer = slBuyer
        myRst!CntrStart = slStart
        myRst!cntrEnd = slEnd
        myRst!ContractId = llContractId
        myRst.Update
        myRst.MoveNext
    Loop
Cleanup:
    mAstFileGetContracts = blRet
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileGetContracts.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup
End Function


Private Sub mAstFileContractInfo(llContractNo As Long, slBuyer As String, llContract As Long, slStart As String, slEnd As String, ilAgy As Integer)
    'Out: all but llPreviousContract
    Dim slSql As String

    slBuyer = ""
    slStart = ""
    slEnd = ""
    llContract = 0
    ilAgy = 0
    If Len(llContractNo) > 0 Then
On Error GoTo ErrHand
        'get agency as it exists nowhere else!
         slSql = " Select chfCode,  chfCntrNo,chfBuyer, chfStartDate,chfEndDate,chfAgfCode from CHF_Contract_Header where (chfSchStatus = 'F' or chfschstatus = 'M') AND chfDelete = 'N' AND  chfCntrNo = " & llContractNo
        Set rst = gSQLSelectCall(slSql)
        If Not rst.EOF Then
            slBuyer = rst!chfBuyer
            llContract = rst!chfCode
            slStart = rst!chfStartDate
            slEnd = rst!chfEndDate
            ilAgy = rst!chfAgfCode
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "XDigitalExportLog.txt", "Export XDS-mAstFileContractInfo"
End Sub

Private Function mAstFileWriteUnits(myRst As ADODB.Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim slSendDate As String
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        .MoveFirst
        Do While Not .EOF
            '7161 pledge changed to feed
            If .Fields("FeedDate") <> slPreviousDate Then
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Units_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    '6539
                    If myTransFile.isNew Then
                    '6679 added to header
                        myTransFile.WriteFacts "UnitID,UnitContractID,UnitISCIID,ProgramID,UnitStationID, XDSiteID,UnitScheduledDate,AdvID,AgencyID"
                    End If
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If blContinue Then
                '7161 field name changed
                slSendDate = Format(.Fields("FeedDate"), "YYYY-MM-DD") & " " & .Fields("FeedTime")
                '6679 added advid and agency id
                slLine = myTransFile.CsvSafe(.Fields("UnitID")) & "," & .Fields("ContractID") & "," & .Fields("ISCIID") & "," & .Fields("ProgramID") & "," & .Fields("StationID") & "," & .Fields("XDSiteId") & "," & slSendDate & "," & .Fields("AdvID") & "," & .Fields("AgencyID")
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
    End With
Cleanup:
    mAstFileWriteUnits = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileWriteUnits.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup

End Function
Private Function mAstFileWriteStations(myRst As ADODB.Recordset, smExportDirectory As String, olDmas As Dictionary) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousStation As Integer
    Dim slStationAddress As String
    Dim slStationGroup As String
    Dim llStationGroupId As Long
    Dim ilDmaId As Integer
   ' Dim slDMA As String
    Dim olStationOwners As Dictionary
   ' Dim olDmas As Dictionary
    
    llStationGroupId = 0
    ilDmaId = 0
    slStationAddress = ""
   ' slDMA = ""
   'dan 03/10/14 this was just wrong
   ' ilPreviousStation = 2
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    Set olStationOwners = New Dictionary
    'Set olDmas = New Dictionary
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '7161
        .Sort = "FeedDate, StationID"
        .MoveFirst
        Do While Not .EOF
            '7161
            If .Fields("FeedDate") <> slPreviousDate Then
                'reset. must write station
                ilPreviousStation = 0
                'station groups
                If Not mAstFileWriteStationGroups(olStationOwners, smExportDirectory, slPreviousDate) Then
                    blRet = False
                End If
                olStationOwners.RemoveAll
                '6680 remove from here
'                If Not mAstFileWriteDmas(olDmas, smExportDirectory, slPreviousDate) Then
'                    blRet = False
'                End If
'                olDmas.RemoveAll
                '7161
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Stations_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
               ' myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                   ' myTransFile.WriteFacts "StationID,GroupID,StationName,StationDMAID,SiteID,StationAddress"
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "StationID,GroupID,StationName,StationDMAID,SiteID,StationAddress"
                    End If
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousStation <> .Fields("StationId") Then
                ilPreviousStation = .Fields("StationID")
                blContinue = True
            Else
                blContinue = False
            End If
            If blContinue Then
                slStationAddress = mAstFileGetStation(.Fields("StationId"), ilDmaId, llStationGroupId)
                slLine = .Fields("StationId") & "," & llStationGroupId & "," & myTransFile.CsvSafe(.Fields("CallLetters")) & "," & ilDmaId & "," & .Fields("XdSiteId") & "," & myTransFile.CsvSafe(slStationAddress)
                If Not olStationOwners.Exists(llStationGroupId) Then
                    slStationGroup = mAstFileGetOwner(llStationGroupId)
                    olStationOwners.Add llStationGroupId, slStationGroup
                End If
                '6680 change dictionary.  Now relate station to dmaID, and pass out of function
                If Not olDmas.Exists(ilPreviousStation) Then
                    olDmas.Add ilPreviousStation, ilDmaId
                End If
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
     End With
    'for last file
     If Not mAstFileWriteStationGroups(olStationOwners, smExportDirectory, slPreviousDate) Then
         blRet = False
     End If
     '6680 remove from here
'     If Not mAstFileWriteDmas(olDmas, smExportDirectory, slPreviousDate) Then
'         blRet = False
'     End If
Cleanup:
    mAstFileWriteStations = blRet
    Set myTransFile = Nothing
    Set olStationOwners = Nothing
   ' Set olDmas = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileWriteStations.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup

End Function
Private Function mAstFileGetStation(ilShtt As Integer, ilDma As Integer, llArtt As Long) As String
    '0 is station address
    ' also, ilDma, ilStationGroup,
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slRet As String
    Dim slDivider As String
    
    slDivider = " "
    slRet = ""
    If Len(ilShtt) > 0 Then
On Error GoTo ErrHand
        slSql = "Select shttOwnerArttCode,shttMktCode,shttAddress1,shttAddress2,shttcity,shttState,shttCountry,shttZip from shtt where shttcode = " & ilShtt
        Set myRst = gSQLSelectCall(slSql)
        If Not myRst.EOF Then
            llArtt = myRst!shttOwnerArttCode
            ilDma = myRst!shttMktCode
            slRet = Trim$(myRst!shttAddress1) & slDivider & Trim$(myRst!shttAddress2) & slDivider & Trim$(myRst!shttCity) & slDivider & Trim$(myRst!shttState) & slDivider & Trim$(myRst!shttCountry) & slDivider & Trim$(myRst!shttZip)
        End If
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAstFileGetStation = slRet
    Exit Function
ErrHand:
    gHandleError "XDigitalExportLog.txt", "Export XDS-mAstFileGetStation"
    GoTo Cleanup
End Function
Private Function mAstFileGetOwner(llArtt As Long) As String
    '0 is station owner
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slRet As String

On Error GoTo ErrHand
    slRet = ""
    If llArtt > 0 Then
        slSql = "select (coalesce(rtrim(arttFirstName),' ') + ' ' + coalesce(rtrim(arttLastName),' ')) as Name from artt where arttCode = " & llArtt
        Set myRst = gSQLSelectCall(slSql)
        If Not myRst.EOF Then
            slRet = myRst!Name
        End If
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAstFileGetOwner = slRet
    Exit Function
ErrHand:
    gHandleError "XDigitalExportLog.txt", "Export XDS-mAstFileGetOwner"
    GoTo Cleanup
End Function
Private Function mAstFileGetDma(ilMkt As Integer, ilRank) As String
    '6680 add rank
    '0 is station owner.  also, ilRank
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slRet As String

On Error GoTo ErrHand
    slRet = ""
    If ilMkt > 0 Then
        slSql = "select mktName as Name,mktRank as Rank from mkt where mktCode = " & ilMkt
        Set myRst = gSQLSelectCall(slSql)
        If Not myRst.EOF Then
            slRet = myRst!Name
            ilRank = myRst!Rank
        End If
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAstFileGetDma = slRet
    Exit Function
ErrHand:
    gHandleError "XDigitalExportLog.txt", "Export XDS-mAstFileGetDma"
    GoTo Cleanup
End Function
Private Function mAstFileWriteStationGroups(StationGroups As Dictionary, smExportDirectory As String, slDate As String) As Boolean
    'returns true if empty
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim slLine As String
    Dim vArray As Variant
    Dim c As Integer
    
    blRet = True
On Error GoTo ERRORBOX
    With StationGroups
        If .Count > 0 Then
            Set myTransFile = New CLogger
            slFileName = "Transparency_StationGroups_" & Format(slDate, "yy-mm-dd") & ".csv"
            '6539
            'myTransFile.CleanFile smExportDirectory & slFileName, 0
            myTransFile.LogPath = smExportDirectory & slFileName
            If myTransFile.isLog Then
                'myTransFile.WriteFacts "StationGroupID,StationGroupName"
                '6539
                If myTransFile.isNew Then
                    myTransFile.WriteFacts "StationGroupID,StationGroupName"
                End If
                vArray = StationGroups.Items
                For c = 0 To .Count - 1
                    slLine = .Keys(c) & "," & myTransFile.CsvSafe(CStr(vArray(c)))
                    myTransFile.WriteFacts slLine
                Next c
            Else
                blRet = False
            End If
        End If
    End With
Cleanup:
    mAstFileWriteStationGroups = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileWriteStationGroups.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup

End Function
Private Function mAstFileWriteDmas(myRst As ADODB.Recordset, smExportDirectory As String, olDmas As Dictionary) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousStation As Integer
    Dim ilDmaId As Integer
    Dim slName As String
    Dim ilRank As Integer
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        .MoveFirst
        Do While Not .EOF
            '7161
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousStation = 0
                '7161
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_DMAs_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "DmaID,DmaName,Ranking,UnitID"
                    End If
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If blContinue Then
                If ilPreviousStation <> .Fields("StationId") Then
                    ilPreviousStation = .Fields("StationID")
                    If olDmas.Exists(ilPreviousStation) Then
                        ilDmaId = olDmas.Item(ilPreviousStation)
                        slName = mAstFileGetDma(ilDmaId, ilRank)
                    Else
                        ilDmaId = 0
                        slName = ""
                        ilRank = 0
                    End If
                End If
                If ilDmaId > 0 Then
                    slLine = ilDmaId & "," & myTransFile.CsvSafe(slName) & "," & ilRank & "," & .Fields("UnitID")
                    myTransFile.WriteFacts slLine
                End If
            End If
            .MoveNext
        Loop
     End With
Cleanup:
    mAstFileWriteDmas = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error in mAstFileWriteDmas.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup

End Function
Private Function mAstFileWriteContracts(myRst As ADODB.Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousContract As Integer
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '7161
        .Sort = "FeedDate, ContractID"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousContract = 0
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Contracts_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    'myTransFile.WriteFacts "ContractID,ContractNumber,AgencyID,AdvertiserID,Buyer,StartDate,EndDate"
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "ContractID,ContractNumber,AgencyID,AdvertiserID,Buyer,StartDate,EndDate"
                    End If
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousContract <> .Fields("ContractId") Then
                ilPreviousContract = .Fields("ContractID")
                blContinue = True
            Else
                blContinue = False
            End If
            If blContinue Then
                slLine = .Fields("ContractID") & "," & .Fields("Contract") & "," & .Fields("AgencyId") & "," & .Fields("AdvId") & "," & myTransFile.CsvSafe(.Fields("Buyer")) & "," & Format(Trim(.Fields("cntrStart")), "YYYY-MM-DD") & "," & Format(Trim(.Fields("cntrEnd")), "YYYY-MM-DD")
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
    End With
Cleanup:
   mAstFileWriteContracts = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error inmAstFileWriteContracts.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup

End Function
Private Function mAstFileWriteAgencies(myRst As ADODB.Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousAgency As Integer
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '7161
        .Sort = "FeedDate, AgencyID"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousAgency = 0
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Agencies_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                 '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "AgencyID,AgencyName"
                    End If
                   ' myTransFile.WriteFacts "AgencyID,AgencyName"
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousAgency <> .Fields("AgencyId") Then
                ilPreviousAgency = .Fields("AgencyID")
                blContinue = True
            Else
                blContinue = False
            End If
            If blContinue Then
                slLine = .Fields("AgencyID") & "," & myTransFile.CsvSafe(.Fields("Agency"))
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
    End With
Cleanup:
   mAstFileWriteAgencies = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error inmAstFileWriteAgencies.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup
End Function
Private Function mAstFileWriteAdvertisers(myRst As ADODB.Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousAdvertisers As Integer
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '7161
        .Sort = "FeedDate, AdvID"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousAdvertisers = 0
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Advertisers_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "AdvertisersID,AdvertisersName"
                    End If
                    'myTransFile.WriteFacts "AdvertisersID,AdvertisersName"
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousAdvertisers <> .Fields("AdvId") Then
                ilPreviousAdvertisers = .Fields("AdvID")
                blContinue = True
            Else
                blContinue = False
            End If
            If blContinue Then
                slLine = .Fields("AdvID") & "," & myTransFile.CsvSafe(.Fields("Client"))
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
    End With
Cleanup:
   mAstFileWriteAdvertisers = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error inmAstFileWriteAdvertisers.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup
End Function
Private Function mAstFileWriteISCI(myRst As ADODB.Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousISCI As Integer
    Dim slAbbr As String
    Dim slAbbr_ISCI As String
    Dim llAdf As Long
    Dim slAddAdvtToISCI As String
    Dim ilValue10 As Integer
    '6676
    Dim llPreviousContract As Long
    Dim ilPos As Integer
    Dim ilEndPos As Integer
    
    '9/25/13: Add advertiser abbreviation to ISCI name
    slAddAdvtToISCI = "N"
    SQLQuery = "Select spfUsingFeatures10 From SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue10 = Asc(rst!spfUsingFeatures10)
        If (ilValue10 And ADDADVTTOISCI) = ADDADVTTOISCI Then
            slAddAdvtToISCI = "Y"
        End If
    End If
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '6676 added contract 7161
        .Sort = "FeedDate, ISCIID, Contract"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousISCI = 0
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_ISCIs_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "ISCIID,AdvertiserId,AdvAbbr,ISCIName,ContractID"
                    End If
                    'myTransFile.WriteFacts "ISCIID,AdvertiserId,ISCIName"
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousISCI <> .Fields("isciId") Then
                ilPreviousISCI = .Fields("isciID")
                llPreviousContract = -1
                blContinue = True
            Else
                blContinue = False
            End If
            '6676 contract has changed, though isci hasn't
            If llPreviousContract <> .Fields("Contract") Then
                llPreviousContract = .Fields("Contract")
                blContinue = True
            End If
            If blContinue Then
                '6676 get abbr.
                slAbbr = ""
                llAdf = gBinarySearchAdf(CLng(.Fields("AdvId")))
                If llAdf <> -1 Then
                    slAbbr = Trim$(Left(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6))
                    If slAbbr = "" Then
                        slAbbr = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
                    End If
                End If
                slAbbr_ISCI = .Fields("Isci")
                ilPos = InStr(1, slAbbr_ISCI, "(")
                ilEndPos = InStrRev(slAbbr_ISCI, ")")
                If ilPos > 0 And ilEndPos > 0 And ilEndPos > ilPos Then
                    slAbbr_ISCI = Mid(slAbbr_ISCI, ilPos + 1, ilEndPos - ilPos - 1)
                End If
'                If slAddAdvtToISCI = "N" Then
'                    'llAdf = gBinarySearchAdf(CLng(.Fields("AdvId")))
'                    'If llAdf <> -1 Then
'                    '    slAbbr = Trim$(Left(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr), 6))
'                    '    If slAbbr = "" Then
'                    '        slAbbr = Trim$(Left(tgAdvtInfo(llAdf).sAdvtName, 6))
'                    '    End If
'                    '    slAbbr_ISCI = slAbbr & "(" & Trim$(.Fields("Isci")) & ")"
'                    'Else
'                    '    blContinue = False
'                    'End If
'                    slAbbr_ISCI = .Fields("Isci")
'                Else
'                    slAbbr_ISCI = .Fields("Isci")
'                End If
            End If
            If blContinue Then
                'slLine = .Fields("IsciID") & "," & myTransFile.CsvSafe(.Fields("AdvId")) & "," & myTransFile.CsvSafe(.Fields("Isci"))
                slLine = .Fields("IsciID") & "," & myTransFile.CsvSafe(.Fields("AdvId")) & "," & myTransFile.CsvSafe(slAbbr) & "," & myTransFile.CsvSafe(slAbbr_ISCI) & "," & llPreviousContract
                myTransFile.WriteFacts slLine
            End If
            .MoveNext
        Loop
    End With
Cleanup:
   mAstFileWriteISCI = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error inmAstFileWriteISCI.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup
End Function
Private Function mAstFileWritePrograms(myRst As Recordset, smExportDirectory As String) As Boolean
    Dim blRet As Boolean
    Dim myTransFile As CLogger
    Dim slFileName As String
    Dim blContinue As Boolean
    Dim slLine As String
    Dim slPreviousDate As String
    Dim ilPreviousProgram As Integer
    Dim llVef As Long
    Dim slName As String
    
    blContinue = False
    slPreviousDate = ""
    Set myTransFile = New CLogger
    blRet = True
On Error GoTo ERRORBOX
    With myRst
        '7161
        .Sort = "FeedDate, ProgramID"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FeedDate") <> slPreviousDate Then
                ilPreviousProgram = 0
                slPreviousDate = .Fields("FeedDate")
                slFileName = "Transparency_Programs_" & Format(slPreviousDate, "yy-mm-dd") & ".csv"
                '6539
                'myTransFile.CleanFile smExportDirectory & slFileName, 0
                myTransFile.LogPath = smExportDirectory & slFileName
                If myTransFile.isLog Then
                    blContinue = True
                    '6539
                    If myTransFile.isNew Then
                        myTransFile.WriteFacts "ProgramID,ProgramName"
                    End If
                    'myTransFile.WriteFacts "ProgramID,ProgramName"
                Else
                    blContinue = False
                    blRet = False
                End If
            End If
            If ilPreviousProgram <> .Fields("ProgramId") Then
                ilPreviousProgram = .Fields("ProgramID")
                blContinue = True
            Else
                blContinue = False
            End If
            If blContinue Then
                '9/25/13: Exclude station and show chile vehicle name instead of the parent name
                'slLine = .Fields("ProgramID") & "," & .Fields("StationId") & "," & myTransFile.CsvSafe(.Fields("ProgramName"))
                llVef = gBinarySearchVef(CLng(.Fields("ProgramID")))
                If llVef <> -1 Then
                    slName = Trim$(tgVehicleInfo(llVef).sVehicle)
                    slLine = .Fields("ProgramID") & "," & myTransFile.CsvSafe(slName)
                    myTransFile.WriteFacts slLine
                End If
            End If
            .MoveNext
        Loop
    End With
Cleanup:
   mAstFileWritePrograms = blRet
    Set myTransFile = Nothing
    Exit Function
ERRORBOX:
    gLogMsg "Error inmAstFileWritePrograms.  Error: " & Err.Description, "XDigitalExportLog.txt", False
    blRet = False
    GoTo Cleanup
End Function
Public Function gXDSShortTitle(llAdfSearch As Long, slProduct As String, blNoAbbreviation As Boolean, blAddProduct As Boolean) As String
    '7219
    'blAddProduct only if blNoAbbreviation is false (blNoAbbreviation is for 'CU')
    Dim slShortTitle As String
    Dim llAdf As Long
    
    '7557 match what is in xds
    slShortTitle = "ADV_MISSING-" & llAdfSearch
    'slShortTitle = Trim$(slProduct)
    llAdf = gBinarySearchAdf(llAdfSearch)
    '7429
    If llAdf = -1 Then
        gPopAdvertisers
        llAdf = gBinarySearchAdf(llAdfSearch)
    End If
    If llAdf <> -1 Then
        If blNoAbbreviation Then
            slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtName)
        Else
            slShortTitle = Trim$(Left$(tgAdvtInfo(llAdf).sAdvtAbbr, 6))
            If slShortTitle = "" Then
                slShortTitle = Trim$(Left$(tgAdvtInfo(llAdf).sAdvtName, 6))
            End If
            If blAddProduct Then
                '7557
                slShortTitle = slShortTitle & "," & Trim$(slProduct)
'                slShortTitle = slShortTitle & ", " & Trim$(slProduct)
            End If
        End If
    End If
    gXDSShortTitle = UCase$(gFileNameFilter(slShortTitle))
End Function
'9256 moved
Public Function gStationXmlReceiverChoices(slSection As String, slXMLINIInputFile As String, blStation As Boolean, blAgreement As Boolean) As Boolean
    'returns true if field exists, otherwise false
    'O: blstation and blagreement
    Dim blRet As Boolean
    Dim slRet As String
    
    slRet = ""
    blRet = False
    blStation = False
    blAgreement = False
    gLoadFromIni slSection, STATIONXMLRECEIVERID, slXMLINIInputFile, slRet
    If slRet <> "Not Found" Then
        blRet = True
        Select Case slRet
            Case "A"
                blAgreement = True
            Case "B"
                blAgreement = True
                blStation = True
            Case "S"
                blStation = True
        End Select
    End If
    gStationXmlReceiverChoices = blRet
End Function

'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
Public Sub gBrowseForFolder(CommonDialog As Object, pathTextbox As Object)
    Dim slCurDir As String
    On Error GoTo ErrHandler
    
    slCurDir = CurDir
    CommonDialog.DialogTitle = "Select a directory" 'titlebar
    CommonDialog.InitDir = pathTextbox.Text
    CommonDialog.fileName = "Select a Directory"  'Something in filenamebox
    CommonDialog.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog.Filter = "Directories|*.~~~" 'set files-filter to show dirs only
    CommonDialog.CancelError = True 'allow escape key/cancel
    CommonDialog.ShowSave   'show the dialog screen
    
    If Err <> 32755 Then    ' User didn't chose Cancel.
        pathTextbox.Text = CurDir
        If right(pathTextbox.Text, 1) <> "\" Then
            pathTextbox.Text = pathTextbox.Text & "\"
        End If
    End If
    ChDir slCurDir
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub


