VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIdc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'6411
Dim myRsTimeFacts As ADODB.Recordset
'6247
Dim myRsSchedule As ADODB.Recordset
Dim myService As CWebService
Public GenericISCI As String
Public Notes As String
Public ScheduleInfoMore As String
Private lmTriggerId As Long
Private smErrorMessage As String
Private bmIsConnected As Boolean
Private smUrl As String
Private smRet As String
Private smMajorSplit As String
Private smMinorSplit As String
Dim tmRegionalInfo() As CcsiToIdcRegionalInfo
Private imTestTrigger As Integer
Private smDateStart As String
Public imNumberDays As Integer
'Private dmScheduleDates() As String
'6342
Dim tmRegionalTime() As CcsiToIdcRegionalInfo
Dim tmAllTemp() As CcsiToIdcRegionalInfo
'set to true in mCreateTriggerTime and false in Clear.  So timed triggers must come last.
Private bmUseRegionalTime As Boolean

'properties that 'pass-through'
'NameSpace
'LogPath
'isLog
'isTest
Public Property Let isTest(blValue As Boolean)
    myService.isTest = blValue
End Property
Public Property Get isTest() As Boolean
    isTest = myService.isTest
End Property
Public Property Get TriggerDate() As String
    TriggerDate = smDateStart
End Property
Public Property Let TriggerDate(slDate As String)
    smDateStart = myService.FormatDate(slDate)
End Property
Public Property Let NumberDays(ilDays As Integer)
    imNumberDays = ilDays - 1
End Property
Public Property Get NumberDays() As Integer
    NumberDays = imNumberDays + 1
End Property
Public Property Get IsNoTriggers() As Boolean
    '6508
    Dim blRet As Boolean
    
    If myRsSchedule.EOF Then
        blRet = True
    Else
        blRet = False
    End If
    IsNoTriggers = blRet
End Property
'Public Property Let ScheduleDate(dlDate As Date)
'    Dim c As Integer
'    Dim ilUbound As Integer
'    Dim blAdd As Boolean
'
'    blAdd = True
'    ilUbound = UBound(dmScheduleDates)
'    If ilUbound > 0 Then
'        For c = 0 To ilUbound - 1
'            If DateDiff("d", dlDate, dmScheduleDates(c)) = 0 Then
'                blAdd = False
'                Exit For
'            End If
'        Next c
'    End If
'    If blAdd Then
'        dmScheduleDates(ilUbound) = dlDate
'        ReDim Preserve dmScheduleDates(ilUbound + 1)
'    End If
'End Property
Public Property Let RegionalInfoTime(myValue As CcsiToIdcRegionalInfo)
    Dim ilCurrent As Integer
    
    If Len(myValue.Sites) > 0 Then
        ilCurrent = UBound(tmRegionalTime)
        Set tmRegionalTime(ilCurrent) = New CcsiToIdcRegionalInfo
        With tmRegionalTime(ilCurrent)
            .Definition = myValue.Definition
            .endDate = myValue.endDate
            .RegionalISCIs = myValue.RegionalISCIs
            .Sites = myValue.Sites
            .startDate = myValue.startDate
            .Rotation = myValue.Rotation
            .startTime = myValue.startTime
            .endTime = myValue.endTime
        End With
        ReDim Preserve tmRegionalTime(ilCurrent + 1)
    End If
End Property
Public Property Let RegionalInfo(myValue As CcsiToIdcRegionalInfo)
    Dim ilCurrent As Integer
    
    If Len(myValue.Sites) > 0 Then
        ilCurrent = UBound(tmRegionalInfo)
        Set tmRegionalInfo(ilCurrent) = New CcsiToIdcRegionalInfo
        With tmRegionalInfo(ilCurrent)
            .Definition = myValue.Definition
            .endDate = myValue.endDate
            .RegionalISCIs = myValue.RegionalISCIs
            .Sites = myValue.Sites
            .startDate = myValue.startDate
            .Rotation = myValue.Rotation
            .startTime = myValue.startTime
            .endTime = myValue.endTime
        End With
        ReDim Preserve tmRegionalInfo(ilCurrent + 1)
    End If
End Property
Public Property Get ReturnValue() As String
    ReturnValue = smRet
End Property
Public Property Get SplitMajor() As String
    SplitMajor = smMajorSplit
End Property
Public Property Get SplitMinor() As String
    SplitMinor = smMinorSplit
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = smErrorMessage
End Property
Public Property Get NameSpace() As String
    NameSpace = myService.NameSpace
End Property
Public Property Let LogPath(slvalue As String)
    myService.LogPath = slvalue
End Property
Public Property Get isLog() As Boolean
    isLog = myService.isLog
End Property
Public Property Get SoapUrl() As String
    SoapUrl = smUrl
End Property
Public Property Let SoapUrl(slUrl As String)
    smErrorMessage = ""
    If mUrlSafe(slUrl) Then
        smUrl = slUrl
        mSetSoap
    Else
        smErrorMessage = "The url is not valid.  Export not sent."
        smUrl = ""
    End If
End Property
Public Property Get IsMasterState() As Boolean
    IsMasterState = mIsStateMaster()
End Property
Public Property Get IsConnected() As Boolean
    IsConnected = mTestConnection()
End Property
Public Property Get TriggerId() As Long
    TriggerId = lmTriggerId
End Property
Private Sub Class_Initialize()
    Clear False
    Set myService = New CWebService
    With myService
        .NameSpace = "datacast.com/PMWSDL"
    End With
    Set myRsSchedule = mPrepRecordsetSchedule()
End Sub
Public Function CreateLogName(slPathAndName As String)
    CreateLogName = myService.CreateLogName(slPathAndName)
End Function
Public Sub Clear(blDayOnly As Boolean)
    If Not blDayOnly Then
        GenericISCI = ""
        Notes = ""
        lmTriggerId = 0
        smMajorSplit = ";"
        smMinorSplit = "~"
        smErrorMessage = ""
        ScheduleInfoMore = ""
        imTestTrigger = 0
        smDateStart = ""
        imNumberDays = -1
   End If
   smDateStart = ""
   bmUseRegionalTime = False
   ReDim tmRegionalInfo(0)
   ReDim tmRegionalTime(0)
   ReDim tmAllTemp(0)
   ReDim dmScheduleDates(0)
   'can I do this faster?
   Set myRsTimeFacts = mPrepRecordsetTimeFacts()
End Sub
Public Sub LogStart()
    'simple message, and tests to see if need new file
    myService.LogStart
End Sub
Public Sub LogEnd()
    myService.LogEnd
End Sub
Public Function IsTriggerOnIDC(llTriggerId As Long) As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    On Error GoTo ERRORBOX
    smRet = ""
    With myService
        .MethodName = "GetTrigger"
        .Body = .BuildXmlElement("TriggerId", CStr(llTriggerId))
        mSend "Deleted"
        If smRet = "True" Or Len(smErrorMessage) > 0 Then
            blOk = False
        End If
    End With
    IsTriggerOnIDC = blOk
    Exit Function
ERRORBOX:
    IsTriggerOnIDC = False
    smErrorMessage = "That trigger does not exist."
End Function
Public Function CreateTriggerTime() As String
    ' O: string of triggers
'    Dim dlStart As Date
'    Dim dlEnd As Date
    Dim slTriggers As String
    Dim rsTime As ADODB.Recordset
    Dim c As Integer
    Dim ilUpper As Integer
    Dim slSafe As String
    Dim rsAllForTrigger As ADODB.Recordset
    Dim dlStart As Date
    Dim dlEnd As Date
    'for finding schedules for week
    Dim slDatesNeeded As String
    Dim ilDateDiff As Integer
    Dim ilCount As Integer
    Dim slDateToAdd As String
    Dim slDates() As String

    bmUseRegionalTime = True
    slTriggers = ""
 On Error GoTo ERRORBOX
    If mIsComplete() Then
        'load up the times
       ' smDateStart = Format(dlDate, "yyyy-mm-dd")
        Set rsTime = mPrepRecordsetTime()
        ilUpper = UBound(tmRegionalTime) - 1
        For c = 0 To ilUpper
            rsTime.Filter = "Time = #" & tmRegionalTime(c).startTime & "#"
            If rsTime.EOF Then
                rsTime.AddNew Array("Time"), Array(tmRegionalTime(c).startTime)
            End If
        Next c
        For c = 0 To ilUpper
            '6368
            If DateDiff("s", tmRegionalTime(c).endTime, "00:00:00") = 0 Then
                tmRegionalTime(c).endTime = "23:59:58"
            End If
            rsTime.Filter = "Time = #" & tmRegionalTime(c).endTime & "#"
            If rsTime.EOF Then
                rsTime.AddNew Array("Time"), Array(tmRegionalTime(c).endTime)
            End If
        Next c
        rsTime.Filter = adFilterNone
        If rsTime.RecordCount > 0 Then
            rsTime.MoveFirst
            dlStart = Format(rsTime!TIME, "Short Time")
            rsTime.MoveNext
            Do While Not rsTime.EOF
                'reset for each set of times
                Set rsAllForTrigger = mPrepRecordsetAll()
                ReDim tmAllTemp(0) As CcsiToIdcRegionalInfo
                dlEnd = Format(rsTime!TIME, "hh:nn:ss")
                For c = 0 To ilUpper
                    If DateDiff("s", tmRegionalTime(c).startTime, dlStart) >= 0 And DateDiff("s", dlEnd, tmRegionalTime(c).endTime) >= 0 Then
                    'If tmRegionalTime(c).StartTime <= dlStart And tmRegionalTime(c).ENDTIME >= dlEnd Then
                        'add to recordset that will order with the non timed so we can work on station overlap
                        rsAllForTrigger.AddNew Array("Index", "Rotation", "isTimed"), Array(c, tmRegionalTime(c).Rotation, True)
                        'this is the set of timed regional for one trigger (one discreet set of time).  I need to know the days valid
                        ilDateDiff = DateDiff("d", tmRegionalTime(c).startDate, tmRegionalTime(c).endDate)
                        For ilCount = 0 To ilDateDiff
                            slDateToAdd = DateAdd("d", ilCount, tmRegionalTime(c).startDate)
                            If InStr(1, slDatesNeeded, slDateToAdd) = 0 Then
                                slDatesNeeded = slDatesNeeded & slDateToAdd & smMajorSplit
                            End If
                        Next ilCount
                        'Dan 3/13/15: Move after loop
                        'If Len(slDatesNeeded) > 1 Then
                        '    '6737 ";" must be last letter!
                        '    If InStrRev(slDatesNeeded, smMajorSplit) = Len(slDatesNeeded) Then
                        '        slDatesNeeded = mLoseLastLetter(slDatesNeeded)
                        '    End If
                        'End If
                    End If
                Next c
                'Dan 3/13/15
                If Len(slDatesNeeded) > 1 Then
                    '6737 ";" must be last letter!
                    If InStrRev(slDatesNeeded, smMajorSplit) = Len(slDatesNeeded) Then
                        slDatesNeeded = mLoseLastLetter(slDatesNeeded)
                    End If
                End If
                For c = 0 To UBound(tmRegionalInfo) - 1
                    'add to recordset that will order the non timed
                    rsAllForTrigger.AddNew Array("Index", "Rotation", "isTimed"), Array(c, tmRegionalInfo(c).Rotation, False)
                Next c
                'now we can test stations. If pass, add to array
                If rsAllForTrigger.RecordCount > 0 Then
                    rsAllForTrigger.MoveFirst
                    Do While Not rsAllForTrigger.EOF
                        If rsAllForTrigger!istimed Then
                            isSafeStationTime tmAllTemp, tmRegionalTime(rsAllForTrigger!Index).Sites, slSafe
                        Else
                            isSafeStationTime tmAllTemp, tmRegionalInfo(rsAllForTrigger!Index).Sites, slSafe
                        End If
                        If Len(slSafe) > 2 Then
                            If rsAllForTrigger!istimed Then
                                mTempRegionalInfoTime tmRegionalTime(rsAllForTrigger!Index), tmAllTemp, slSafe
                            Else
                                mTempRegionalInfoTime tmRegionalInfo(rsAllForTrigger!Index), tmAllTemp, slSafe
                            End If
                        End If
                        rsAllForTrigger.MoveNext
                    Loop
                    'Now I've got all regional for one period
                    'creates lmTriggerId.
                    mCreateTrigger
                    mCreateTimeFacts lmTriggerId
                    If lmTriggerId > 0 Then
                        slTriggers = slTriggers & CStr(lmTriggerId) & smMajorSplit
                    End If
                    '6252
                    dlStart = Format(dlStart, "hh:nn:ss")
                    If DateDiff("s", dlStart, "12:00:00 am") = 0 Then
                        dlStart = DateAdd("s", 1, dlStart)
                    End If
                    dlEnd = Format(dlEnd, "hh:nn:ss")
                    If DateDiff("s", dlEnd, "12:00:00 am") = 0 Then
                        dlEnd = DateAdd("s", -1, dlEnd)
                    End If
                    If lmTriggerId > 0 Then
                        slDates = Split(slDatesNeeded, smMajorSplit)
                        For c = 0 To UBound(slDates) ' - 1
                            myRsSchedule.AddNew Array("Trigger", "Generic", "Day", "StartTime", "EndTime"), Array(lmTriggerId, GenericISCI, slDates(c), dlStart, dlEnd)
                        Next c
                    End If
                    slDatesNeeded = ""
                   ' myRsSchedule.AddNew Array("Trigger", "Generic", "Day", "StartTime", "EndTime"), Array(lmTriggerId, GenericISCI, dlDate, dlStart, dlEnd)
                End If  ' found regionals to sort by rotation
                'clear the previous time stuff!
                dlStart = dlEnd
                rsTime.MoveNext
           Loop
        End If 'rsTime ok
    End If 'have generic
    slTriggers = mLoseLastLetter(slTriggers)
cleanup:
    CreateTriggerTime = slTriggers
    If Not rsTime Is Nothing Then
        If (rsTime.State And adStateOpen) <> 0 Then
            rsTime.Close
        End If
        Set rsTime = Nothing
    End If
    If Not rsAllForTrigger Is Nothing Then
        If (rsAllForTrigger.State And adStateOpen) <> 0 Then
            rsAllForTrigger.Close
        End If
        Set rsAllForTrigger = Nothing
    End If
    Erase slDates
    Exit Function
ERRORBOX:
    lmTriggerId = 0
    slTriggers = ""
    smErrorMessage = smErrorMessage & " " & Err.Description
    GoTo cleanup
End Function
Public Function CreateTrigger() As Long
    ' O:Trigger  Prep schedule
    Dim dlStart As Date
    Dim dlEnd As Date
    Dim c As Integer
    Dim dlDate As Date
    
 On Error GoTo ERRORBOX
    If mIsComplete() Then
        'creates lmTriggerId.
        mCreateTrigger
        '6252
        dlStart = Format(tmRegionalInfo(0).startTime, "short time")
        If DateDiff("s", dlStart, "12:00:00 am") = 0 Then
            dlStart = DateAdd("s", 1, dlStart)
        End If
        dlEnd = Format(tmRegionalInfo(0).endTime, "short time")
        If DateDiff("s", dlEnd, "12:00:00 am") = 0 Then
            dlEnd = DateAdd("s", -1, dlEnd)
        End If
        If lmTriggerId > 0 Then
            For c = 0 To imNumberDays
                dlDate = DateAdd("d", c, smDateStart)
                myRsSchedule.AddNew Array("Trigger", "Generic", "Day", "StartTime", "EndTime"), Array(lmTriggerId, GenericISCI, dlDate, dlStart, dlEnd)
            Next c
        End If
    End If
    CreateTrigger = lmTriggerId
    Exit Function
ERRORBOX:
    lmTriggerId = 0
    CreateTrigger = 0
    smErrorMessage = smErrorMessage & Err.Description
End Function

Private Sub mCreateTrigger()
    Dim llNewTrigger As Long
    Dim slDefault As String
    
    slDefault = "false"
    llNewTrigger = 0
    lmTriggerId = 0
    With myService
        .MethodName = "CreateTrigger"
        .Body = .BuildXmlElement("ISCI", Me.GenericISCI) & .BuildXmlElement("Notes", Me.Notes) & .BuildXmlElement("VisibleInReceiver", "true") & .BuildXmlElement("SpliceID", 0) _
        & .BuildXmlElement("ActivateImmediately", "false") & .BuildXmlElement("DefaultTrig", slDefault) & .BuildXmlElement("Deleted", "false") & .BuildXmlElement("Actions", mCreateActions())
        .Body = "<NewTrigger><TriggerID>0</TriggerID>" & .Body & "</NewTrigger>"
    End With
    mSend "CreateTriggerResult"
    If Not Me.isTest Then
        If IsNumeric(smRet) Then
            llNewTrigger = CLng(smRet)
        End If
    Else
        imTestTrigger = imTestTrigger + 1
        llNewTrigger = imTestTrigger
    End If
    lmTriggerId = llNewTrigger
End Sub
Public Function CreateSchedule() As Boolean
    Dim llNewSchedule As Long
    Dim llPreviousDate As Date
    Dim blRet As Boolean
    Dim myClone As ADODB.Recordset
    
    blRet = True
 On Error GoTo ERRORBOX
    llPreviousDate = #1/1/1970#
    If Not myRsSchedule.EOF Then
        Set myClone = myRsSchedule.Clone
        myRsSchedule.MoveFirst
        Do While Not myRsSchedule.EOF
            If llPreviousDate <> myRsSchedule!Day Then
                llPreviousDate = myRsSchedule!Day
                myClone.Filter = "Day = " & llPreviousDate
                If myClone.RecordCount > 0 Then
                    If mCreateSchedule(myClone) = 0 And Not Me.isTest Then
                        blRet = False
                    End If
                End If
            End If
            myRsSchedule.MoveNext
        Loop
    Else
        blRet = False
    End If
    CreateSchedule = blRet
    Exit Function
ERRORBOX:
    CreateSchedule = False
    smErrorMessage = smErrorMessage & " " & Err.Description
End Function
Private Function mCreateSchedule(myClone As ADODB.Recordset) As Long
    Dim llRet As Long
    Dim slWindows As String
    Dim slWindowInfo As String
    Dim slDay As String
    
    llRet = 0
 On Error GoTo ERRORBOX
    With myService
        slWindows = .BuildXmlElement("Windows", mCreateWindows(myClone, slWindowInfo))
        .MethodName = "CreateSchedule"
        .Body = .BuildXmlElement("ScheduleID", "0") & .BuildXmlElement("StartDate", Format(myClone!Day, "yyyy-mm-dd")) & .BuildXmlElement("EndDate", Format(myClone!Day, "yyyy-mm-dd")) & slWindows
        .Body = "<NewSchedule>" & .Body & "</NewSchedule>"
    End With
    mSend "CreateScheduleResult"
    If Len(smRet) > 0 Then
        If IsNumeric(smRet) Then
            llRet = CLng(smRet)
        End If
        myClone.MoveFirst
        slDay = Format(myClone!Day, "mm-dd-yy")
        Do While Not myClone.EOF
            myClone!Schedule = llRet
            myClone.MoveNext
        Loop
        If llRet = 0 Then
            smErrorMessage = "  " & slDay & ". " & smErrorMessage
        End If
    Else
        myClone.MoveFirst
        slDay = Format(myClone!Day, "mm-dd-yy")
        smErrorMessage = "  " & slDay & ". " & smErrorMessage
    End If
    ScheduleInfoMore = ScheduleInfoMore & vbCrLf & "Schedule for " & slDay & vbCrLf & slWindowInfo
    mCreateSchedule = llRet
    Exit Function
ERRORBOX:
    smErrorMessage = smErrorMessage & " " & Err.Description
    mCreateSchedule = 0
End Function
Private Function mCreateWindows(myClone As ADODB.Recordset, slWindowInfo As String) As String
    Dim slRet As String
    Dim slDayFilter As String
    Dim slTimeFilter As String
    Dim rsTime As Recordset
    Dim dlPrevious As Date
    Dim dlStart As Date
    Dim dlEnd As Date
    Dim dlEndToWrite As Date
    Dim slActs As String
    Dim slWindow As String
    Dim slActiveInfo As String
    
    slActiveInfo = ""
    slWindowInfo = ""
    Set rsTime = mPrepRecordsetTime()
    slDayFilter = myClone.Filter
    myClone.Sort = "StartTime"
    myClone.MoveFirst
    dlPrevious = Format("11:59:59", "short time")
    Do While Not myClone.EOF
        If myClone!startTime <> dlPrevious Then
            dlPrevious = myClone!startTime
            rsTime.Filter = "Time = #" & dlPrevious & "#"
            If rsTime.EOF Then
                rsTime.AddNew Array("Time"), Array(dlPrevious)
            End If
        End If
        myClone.MoveNext
    Loop
    myClone.Sort = "EndTime"
    dlPrevious = Format("00:00:01", "short time")
    myClone.MoveFirst
     Do While Not myClone.EOF
        If myClone!endTime <> dlPrevious Then
            dlPrevious = myClone!endTime
            rsTime.Filter = "Time =  #" & dlPrevious & "#"
            If rsTime.EOF Then
                rsTime.AddNew Array("Time"), Array(dlPrevious)
            End If
        End If
        myClone.MoveNext
    Loop
    rsTime.Filter = adFilterNone
    ' I want the later start times first, so I can 'pull out' the same generics and only write the ones that should be written
    myClone.Sort = "StartTime desc"
    If rsTime.RecordCount <> 0 Then
        rsTime.MoveFirst
        dlStart = rsTime!TIME
        rsTime.MoveNext
        Do While Not rsTime.EOF
            dlEnd = rsTime!TIME
        If DateDiff("s", Format(dlEnd, "ss"), "59") = 0 Then
                dlEndToWrite = dlEnd
            Else
                dlEndToWrite = DateAdd("s", -1, dlEnd)
            End If
            myClone.Filter = slDayFilter & " AND StartTime <= #" & dlStart & "# AND EndTime >= #" & dlEnd & "#"
            With myService
                slActs = .BuildXmlElement("Activations", mCreateActivation(myClone, slActiveInfo))
                slWindow = .BuildXmlElement("WindowId", "0") & .BuildXmlElement("ScheduleID", "0") & .BuildXmlElement("StartTime", Format(dlStart, "hh:mm:ss")) & .BuildXmlElement("EndTime", Format(dlEndToWrite, "hh:mm:ss")) & slActs
            End With
            slRet = slRet & "<Window>" & slWindow & "</Window>"
            slWindowInfo = slWindowInfo & "    Window " & Format(dlStart, "hh:mm:ss") & "-" & Format(dlEndToWrite, "hh:mm:ss") & vbCrLf & slActiveInfo & vbCrLf
            dlStart = dlEnd
            rsTime.MoveNext
        Loop
    End If
    If Not rsTime Is Nothing Then
        If (rsTime.State And adStateOpen) <> 0 Then
            rsTime.Close
        End If
        Set rsTime = Nothing
    End If
    myClone.Filter = slDayFilter
    mCreateWindows = slRet
End Function

Private Function mCreateActivation(myClone As ADODB.Recordset, slActiveInfo As String) As String
    Dim slRet As String
    Dim slPart As String
    Dim slWritten() As String
    Dim ilUbound As Integer
    Dim c As Integer
    Dim blFound As Boolean
    
    slRet = ""
    slActiveInfo = ""
    ReDim slWritten(0) As String
    With myService
        Do While Not myClone.EOF
            'edit out already entered generic!
            blFound = False
            ilUbound = UBound(slWritten)
            For c = 0 To ilUbound
                If slWritten(c) = myClone!Generic Then
                    blFound = True
                    Exit For
                End If
            Next c
            If Not blFound Then
                slWritten(ilUbound) = myClone!Generic
                ReDim Preserve slWritten(ilUbound + 1)
                slPart = .BuildXmlElement("ISCI", Trim$(myClone!Generic)) & .BuildXmlElement("Rotate", "true") & .BuildXmlElement("TriggerID", CStr(myClone!trigger))
                slRet = slRet & "<Activation>" & slPart & "</Activation>"
                slActiveInfo = slActiveInfo & CStr(myClone!trigger) & ","
            End If
            myClone.MoveNext
        Loop
    End With
    Erase slWritten
    slActiveInfo = mLoseLastLetter(slActiveInfo)
    slActiveInfo = "        Triggers: " & slActiveInfo
    mCreateActivation = slRet
End Function
Public Function DeleteFromIDC(ilTrigger As Long) As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    On Error GoTo ERRORBOX
    With myService
        .MethodName = "DeleteTrigger"
        .Body = .BuildXmlElement("TriggerId", CStr(ilTrigger))
    End With
    mSend "faultstring"
    If Len(smRet) > 0 Then
        blOk = False
        smErrorMessage = smRet
    End If
    DeleteFromIDC = blOk
    Exit Function
ERRORBOX:
    blOk = False
    smErrorMessage = "Cannot delete trigger. " & Err.Description
    DeleteFromIDC = blOk
End Function
Public Function DeleteFromIDCSchedule(llSchedule As Long) As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    On Error GoTo ERRORBOX
    With myService
        .MethodName = "DeleteSchedule"
        .Body = .BuildXmlElement("ScheduleID", CStr(llSchedule))
    End With
    mSend "faultstring"
    If Len(smRet) > 0 Then
        blOk = False
        smErrorMessage = smRet
    End If
    DeleteFromIDCSchedule = blOk
    Exit Function
ERRORBOX:
    blOk = False
    smErrorMessage = "Cannot delete schedule #" & llSchedule & ". " & Err.Description
    DeleteFromIDCSchedule = blOk
End Function
Public Function DeleteMultiples(slTriggers As String) As Boolean
    Dim slTriggersArray() As String
    Dim c As Integer
    Dim slBody As String
    Dim blOk As Boolean
    
    blOk = True
    slBody = ""
On Error GoTo ERRORBOX
    slTriggersArray = Split(slTriggers, smMajorSplit)
    With myService
        For c = 0 To UBound(slTriggersArray)
            slBody = slBody & .BuildXmlElement("int", slTriggersArray(c))
        Next c
        .MethodName = "DeleteTriggers"
        .Body = .BuildXmlElement("TriggerID", slBody)
    End With
    mSend "DeleteTriggersResponse"
    If Not (smErrorMessage = "" Or myService.isTest) Then
        blOk = False
    End If
cleanup:
    Erase slTriggersArray
    DeleteMultiples = blOk
    Exit Function
ERRORBOX:
    blOk = False
    smErrorMessage = "Cannot delete trigger. " & Err.Description
    GoTo cleanup
End Function
Public Function WriteRegionalInfo(llTrigger As Long) As String
    'llTrigger is for timeSensitive
    Dim slRet As String
    Dim c As Integer
    Dim ilIndex As Integer
    Dim slSites() As String
    Dim slName As String
    Dim slId As String
    Dim slstations As String
    Dim slFiles As String
    Dim ilPos As Integer
    Dim slTypes() As String
    Dim slFile As String
    Dim slOccurrence As String
    '6174
    Dim slCompare As String
 On Error GoTo ERRORBOX
    slRet = ""
    slCompare = smMinorSplit
    If bmUseRegionalTime Then
        '6411
        If Not myRsTimeFacts Is Nothing Then
            myRsTimeFacts.Filter = "Index = " & llTrigger
            Do While Not myRsTimeFacts.EOF
                If myRsTimeFacts!Rotation > 0 Then
                    slSites = Split(myRsTimeFacts!Stations, smMajorSplit)
                    slstations = ""
                    slFiles = ""
                    For c = 0 To UBound(slSites)
                        If GetInfo(slSites(c), slName, slId, slCompare) Then
                            slstations = slstations & Trim(slName) & " id=" & Trim(slId) & "   "
                        End If
                    Next c
                    slTypes = Split(myRsTimeFacts!Files, smMajorSplit)
                    For c = 0 To UBound(slTypes)
                        If GetInfo(slTypes(c), slFile, slOccurrence) Then
                            slFiles = slFiles & Trim(slFile) & " %=" & Trim(slOccurrence) & "    "
                        End If
                    Next c
                    slRet = slRet & "Region:" & Trim(myRsTimeFacts!Definition) & "  Rotation #:" & myRsTimeFacts!Rotation & vbCrLf & " Stations: " & slstations & vbCrLf & " Files:" & slFiles & vbCrLf
                End If
                myRsTimeFacts.MoveNext
            Loop
        Else
            slRet = "Problem writing facts: no time sensitive facts were recorded."
        End If
    Else
        For ilIndex = UBound(tmRegionalInfo) - 1 To 0 Step -1
            With tmRegionalInfo(ilIndex)
                '6175
                If .Rotation > 0 Then
                    slSites = Split(.Sites, smMajorSplit)
                    slstations = ""
                    slFiles = ""
                    For c = 0 To UBound(slSites)
                        If GetInfo(slSites(c), slName, slId, slCompare) Then
                            slstations = slstations & slName & " id=" & slId & "   "
                        End If
                    Next c
                    slTypes = Split(.RegionalISCIs, smMajorSplit)
                    For c = 0 To UBound(slTypes)
                        If GetInfo(slTypes(c), slFile, slOccurrence) Then
                            slFiles = slFiles & slFile & " %=" & slOccurrence & "    "
                        End If
                    Next c
                    slRet = slRet & "Region:" & .Definition & "  Rotation #:" & .Rotation & vbCrLf & " Stations: " & slstations & vbCrLf & " Files:" & slFiles & vbCrLf
                End If
                slCompare = smMinorSplit
            End With
        Next ilIndex
    End If
    Erase slSites
    Erase slTypes
    WriteRegionalInfo = slRet
    Exit Function
ERRORBOX:
    WriteRegionalInfo = ""
End Function
Public Function isAtLeastOneRegionalTime() As Boolean
    Dim blRet As Boolean
 On Error GoTo ERRORBOX
    blRet = True
    If UBound(tmRegionalTime) = 0 Then
       ' smErrorMessage = smErrorMessage & " This trigger has no regions to send."
        blRet = False
    End If
    isAtLeastOneRegionalTime = blRet
    Exit Function
ERRORBOX:
    isAtLeastOneRegionalTime = False
End Function
Public Function isAtLeastOneRegional() As Boolean
    Dim blRet As Boolean
 On Error GoTo ERRORBOX
    blRet = True
    If UBound(tmRegionalInfo) = 0 Then
        blRet = False
'    ElseIf Len(tmRegionalInfo(0).RegionalISCIs) = 0 Then
'        blRet = False
    End If
    isAtLeastOneRegional = blRet
    Exit Function
ERRORBOX:
    isAtLeastOneRegional = False
End Function
Private Function mCreateActions() As String
    Dim c As Integer
    Dim slString As String
    Dim slMinorString As String
    Dim slEndDate As String
    
    slEndDate = DateAdd("d", imNumberDays, smDateStart)
    slEndDate = myService.FormatDate(slEndDate)
    slMinorString = ""
    slString = ""
    '6342 workaround:  Use a different array of regionals if this is 'time sensitive' trigger.
    If bmUseRegionalTime Then
         For c = UBound(tmAllTemp) - 1 To 0 Step -1
            '6175
            If tmAllTemp(c).Rotation > 0 Then
                With myService
                    slMinorString = .BuildXmlElement("ActionID", 0) & .BuildXmlElement("Description", tmAllTemp(c).Definition) & .BuildXmlElement("Delete", "false") _
                    & .BuildXmlElement("DoesNotExpire", "false") & .BuildXmlElement("EarliestDate", smDateStart) & .BuildXmlElement("ExpiryDate", slEndDate) _
                    & .BuildXmlElement("OccurrenceType", "PercentageSplit") & .BuildXmlElement("Sites", mCreateSites(tmAllTemp(c).Sites)) & .BuildXmlElement("Types", mCreateTypes(tmAllTemp(c).RegionalISCIs))
                End With
            End If
            If Len(slMinorString) > 0 Then
                slString = slString & "<Action>" & slMinorString & "</Action>"
            End If
        Next c
    Else
        For c = UBound(tmRegionalInfo) - 1 To 0 Step -1
            '6175
            If tmRegionalInfo(c).Rotation > 0 Then
                With myService
                    slMinorString = .BuildXmlElement("ActionID", 0) & .BuildXmlElement("Description", tmRegionalInfo(c).Definition) & .BuildXmlElement("Delete", "false") _
                    & .BuildXmlElement("DoesNotExpire", "false") & .BuildXmlElement("EarliestDate", smDateStart) & .BuildXmlElement("ExpiryDate", slEndDate) _
                    & .BuildXmlElement("OccurrenceType", "PercentageSplit") & .BuildXmlElement("Sites", mCreateSites(tmRegionalInfo(c).Sites)) & .BuildXmlElement("Types", mCreateTypes(tmRegionalInfo(c).RegionalISCIs))
                End With
            End If
            If Len(slMinorString) > 0 Then
                slString = slString & "<Action>" & slMinorString & "</Action>"
            End If
        Next c
    End If
    mCreateActions = slString
End Function
Private Function mCreateSites(slSiteInput As String) As String
    Dim slSites() As String
    Dim c As Integer
    Dim ilPos As Integer
    Dim slName As String
    Dim slId As String
    Dim slMinorString As String
    Dim slString As String
    '6174
    Dim slCompare As String
    
    slCompare = smMinorSplit
    slSites = Split(slSiteInput, smMajorSplit)
    For c = 0 To UBound(slSites)
        If GetInfo(slSites(c), slName, slId, slCompare) Then
            With myService
                slMinorString = "<Site>" & .BuildXmlElement("SiteName", slName) & .BuildXmlElement("SiteID", slId) & "</Site>"
            End With
            slString = slString & slMinorString
        End If
    Next c
    Erase slSites
    mCreateSites = slString
End Function
Private Function mCreateTypes(slTypeInput As String) As String
    Dim slTypes() As String
    Dim c As Integer
    Dim ilPos As Integer
    Dim slFile As String
    Dim slOccurrence As String
    Dim slMinorString As String
    Dim slRet As String
    
    slTypes = Split(slTypeInput, smMajorSplit)
    For c = 0 To UBound(slTypes)
        If GetInfo(slTypes(c), slFile, slOccurrence) Then
            With myService
                slMinorString = "<ActionType>" & .BuildXmlElement("Filename", slFile) & .BuildXmlElement("Type", "PlayFile") & .BuildXmlElement("Occurrence", slOccurrence) & .BuildXmlElement("TypeID", "0") & "</ActionType>"
            End With
            slRet = slRet & slMinorString
        End If
    Next c
    mCreateTypes = slRet
End Function

Private Function mIsComplete() As Boolean
    Dim blOk As Boolean

    blOk = True
    smErrorMessage = ""
    If Len(smUrl) = 0 Then
        smErrorMessage = "Soap Url cannot be empty."
        blOk = False
    End If
    If Len(GenericISCI) = 0 Then
        smErrorMessage = smErrorMessage & " Generic ISCI cannot be empty."
        blOk = False
    End If
    If Len(smDateStart) = 0 Then
        smErrorMessage = smErrorMessage & " Trigger Date cannot be empty."
        blOk = False
    End If
    If imNumberDays < 0 Or imNumberDays > 6 Then
        smErrorMessage = smErrorMessage & " Number of days must be between 1 and 7."
        blOk = False
    End If
'    If UBound(dmScheduleDates) = 0 Then
'        smErrorMessage = smErrorMessage & " Schedule Date cannot be empty."
'        blOk = False
'    End If
    mIsComplete = blOk
End Function
'Private Function mIsComplete(blRegionalTime As Boolean) As Boolean
'    Dim blOk As Boolean
'
'    blOk = True
'    smErrorMessage = ""
'    If Len(smUrl) = 0 Then
'        smErrorMessage = "Soap Url cannot be empty."
'        blOk = False
'    End If
'    If Len(GenericISCI) = 0 Then
'        smErrorMessage = smErrorMessage & " Generic ISCI cannot be empty."
'        blOk = False
'    End If
'    If blRegionalTime Then
'        If Not isAtLeastOneRegionalTime() Then
'            blOk = False
'        End If
'    Else
'        If Not isAtLeastOneRegional() Then
'            blOk = False
'        End If
'    End If
'    mIsComplete = blOk
'End Function
Private Function mUrlSafe(slUrl As String) As Boolean
    Dim blRet As Boolean
    
    blRet = False
    If Len(slUrl) > 0 Then
        If InStr(1, slUrl, "ttp:", vbTextCompare) > 0 And InStr(1, slUrl, "/PMWebService/PMWSDL.asmx") > 0 Then
            blRet = True
        ElseIf UCase(slUrl) = "TEST" Then
            blRet = True
        ElseIf InStr(1, slUrl, "XD", vbTextCompare) > 0 Then
            blRet = True
        End If
    End If
    mUrlSafe = blRet
End Function
Private Sub mSetSoap()
    myService.WebServiceAddress = smUrl
End Sub
Private Function mIsStateMaster() As Boolean
    Dim blRet As Boolean
    
    If Len(smUrl) > 0 Then
        With myService
            .MethodName = "GetStatus"
        End With
        mSend "RedundancyState"
        If Len(smErrorMessage) = 0 Then
            If smRet = "Master" Then
                blRet = True
            ElseIf myService.isTest Then
                blRet = True
            End If
        End If
    Else
        smErrorMessage = "Url not set!"
    End If
    mIsStateMaster = blRet
End Function
Private Function mTestConnection() As Boolean
    Dim blRet As Boolean
    
    If Len(smUrl) > 0 Then
        With myService
            .MethodName = "GetStatus"
        End With
        mSend "DatabaseStatus"
        If Len(smErrorMessage) = 0 Then
            If smRet = "Connected" Then
                blRet = True
            ElseIf myService.isTest Then
                blRet = True
            End If
        End If
    Else
        smErrorMessage = "Url not set!"
    End If
    mTestConnection = blRet
End Function
Private Sub mSend(slReturnField As String)
    Dim slvalue As String
    
    With myService
        If .Send() Then
            If .isTest Then
                slvalue = .ReturnValue
            Else
                slvalue = .Read(slReturnField)
            End If
            If Len(slvalue) = 0 Then
                smErrorMessage = .ErrorMessage
                smRet = ""
            Else
                smRet = slvalue
            End If
        Else
            smErrorMessage = .ErrorMessage
        End If
    End With
End Sub
Private Sub mSendXD(slReturnField As String)
    Dim slvalue As String
    
    With myService
        If .SendSpecial1() Then
            If .isTest Then
                slvalue = .ReturnValue
            Else
                slvalue = .Read(slReturnField)
            End If
            If Len(slvalue) = 0 Then
                smErrorMessage = .ErrorMessage
                smRet = ""
            Else
                smRet = slvalue
            End If
        Else
            smErrorMessage = .ErrorMessage
        End If
    End With
End Sub
Public Function isSafeStationTime(slTimeTemp() As CcsiToIdcRegionalInfo, slSiteInput As String, slSafe As String) As String
    Dim slName As String
    Dim slId As String
    Dim c As Integer
    Dim slTestLine As String
    Dim slSites() As String
    Dim slNewLine As String
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim slUnsafe As String
    Dim ilPos As Integer
    Dim ilLength As Integer
    '6345
    Dim blSafe As Boolean
    
On Error GoTo ERRORBOX
    slUnsafe = ""
    slSafe = ""
    slNewLine = ""
    If UBound(slTimeTemp) > 0 Then
        slSites = Split(slSiteInput, smMajorSplit)
        For ilIndex = 0 To UBound(slSites)
            '10/28/13
            blSafe = True
            GetInfo slSites(ilIndex), slName, slId
            If slId > 0 Then
                For c = 0 To UBound(slTimeTemp) - 1 Step 1
                    slTestLine = slTimeTemp(c).Sites & smMajorSplit
                    ilPos = InStr(1, slTestLine, smMinorSplit & slId & smMajorSplit)
                    If ilPos > 0 Then
                        blSafe = False
                    End If
                Next c
            End If
            '6345
            If blSafe Then
                If InStr(1, slSafe, smMinorSplit & slId & smMajorSplit, vbTextCompare) = 0 Then
                    slSafe = slSafe & slName & smMinorSplit & slId & smMajorSplit
                End If
            End If
        Next ilIndex
    Else
        slSafe = slSiteInput & smMajorSplit
    End If
    If Len(slSafe) > 0 Then
        slSafe = mLoseLastLetter(slSafe)
    End If
cleanup:
    isSafeStationTime = slUnsafe
    Erase slSites
    Exit Function
ERRORBOX:
    GoTo cleanup
End Function
Public Function IsSafeStation(slSiteInput As String, slSafe As String) As String
    Dim slName As String
    Dim slId As String
    Dim c As Integer
    Dim slTestLine As String
    Dim slSites() As String
    Dim slNewLine As String
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    Dim slUnsafe As String
    Dim ilPos As Integer
    Dim ilLength As Integer
    '6345
    Dim blSafe As Boolean
    
On Error GoTo ERRORBOX
    slUnsafe = ""
    slSafe = ""
    slNewLine = ""
   ' blSafe = True
    If UBound(tmRegionalInfo) > 0 Then
        slSites = Split(slSiteInput, smMajorSplit)
        For ilIndex = 0 To UBound(slSites)
            '10/28/13 Dan
            blSafe = True
            GetInfo slSites(ilIndex), slName, slId
            If slId > 0 Then
                For c = 0 To UBound(tmRegionalInfo) - 1 Step 1
                    slTestLine = tmRegionalInfo(c).Sites & smMajorSplit
                    ilPos = InStr(1, slTestLine, smMinorSplit & slId & smMajorSplit)
                    If ilPos > 0 Then
                        blSafe = False
                    End If
                Next c
            End If
            '6345
            If blSafe Then
                If InStr(1, slSafe, smMinorSplit & slId & smMajorSplit, vbTextCompare) = 0 Then
                    slSafe = slSafe & slName & smMinorSplit & slId & smMajorSplit
                End If
            End If
        Next ilIndex
    Else
        slSafe = slSiteInput & smMajorSplit
    End If
cleanup:
    IsSafeStation = slUnsafe
    Erase slSites
    Exit Function
ERRORBOX:
    GoTo cleanup
End Function

Public Function BuildLine(slLineToAddTo As String, slFirstValue As String, slSecondValue As String) As String
    Dim slNewLine As String
    
    slNewLine = slLineToAddTo
    If slNewLine = "" Then
        slNewLine = Trim$(slFirstValue) & smMinorSplit & Trim$(slSecondValue)
    Else
        slNewLine = slNewLine & smMajorSplit & Trim$(slFirstValue) & smMinorSplit & Trim$(slSecondValue)
    End If
    BuildLine = slNewLine
End Function
Public Function GetInfo(slValues As String, slFirstValue As String, slSecondValue As String, Optional slCompareString As String = "") As Boolean
    Dim ilPos As Integer
    Dim blRet As Boolean
    '6174 slcomparestring.  only return true if not already there. Action will only have site once
    blRet = False
    slSecondValue = ""
    slFirstValue = ""
    ilPos = InStr(1, slValues, smMinorSplit, vbBinaryCompare)
    If ilPos > 0 And ilPos < Len(slValues) Then
        slFirstValue = Mid(slValues, 1, ilPos - 1)
        slSecondValue = Mid(slValues, ilPos + 1)
        blRet = True
    End If
    If Len(slCompareString) > 0 Then
        If InStr(1, slCompareString, smMinorSplit & slSecondValue & smMinorSplit) > 0 Then
            blRet = False
        Else
            slCompareString = slCompareString & slSecondValue & smMinorSplit
        End If
    End If
    GetInfo = blRet
End Function
Public Function GetScheduleInfo(blIsStart As Boolean, llTrigger As Long, llSchedule As Long, slDate As String) As Boolean
    'first time in, use blIsStart.  Always check is true before using values.
    'O trigger, schedule and date
    llTrigger = 0
    llSchedule = 0
    slDate = ""
    If myRsSchedule Is Nothing Then
        GetScheduleInfo = False
        Exit Function
    End If
    If blIsStart Then
        myRsSchedule.Filter = adFilterNone
        If myRsSchedule.RecordCount = 0 Then
            GetScheduleInfo = False
            Exit Function
        End If
        myRsSchedule.MoveFirst
    ElseIf myRsSchedule.EOF Then
        GetScheduleInfo = False
        Exit Function
    Else
        myRsSchedule.MoveNext
    End If
    If myRsSchedule.EOF Then
        GetScheduleInfo = False
        Exit Function
    End If
    llTrigger = myRsSchedule!trigger
    llSchedule = myRsSchedule!Schedule
    slDate = Format(myRsSchedule!Day, "yyyy-mm-dd")
    GetScheduleInfo = True
End Function
Public Function GetSiteId(slSiteId As String) As Boolean
    Dim blOk As Boolean
    
    blOk = True
    smErrorMessage = ""
    On Error GoTo ERRORBOX
    smRet = ""
    With myService
       ' .NameSpace = "datacast.com"
        .MethodName = "dat:getSiteInformationBySiteID"
        .Body = "<param-1 xsi:type=""xsd:int"">" & slSiteId & "</param-1>" '.BuildXmlElement("param-1", slSiteId)
        mSendXD "SiteID"
        If smRet = "False" Or Len(smErrorMessage) > 0 Then
            blOk = False
        End If
    End With
    GetSiteId = blOk
   ' myService.NameSpace = "datacast.com/PMWSDL"
    Exit Function
ERRORBOX:
    GetSiteId = False
    smErrorMessage = "That trigger does not exist."
End Function
Private Function mPrepRecordsetSchedule() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Day", adDate
            .Append "StartTime", adDBTime
            .Append "EndTime", adDBTime
            .Append "Trigger", adInteger
            .Append "Generic", adChar, 15
            .Append "Schedule", adInteger
        End With
    myRs.Open
    myRs!Day.Properties("optimize") = True
    myRs.Sort = "Day"
    Set mPrepRecordsetSchedule = myRs
End Function
Private Function mPrepRecordsetTime() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Time", adDBTime ' adChar, 10
        End With
    myRs.Open
    myRs!TIME.Properties("optimize") = True
    myRs.Sort = "Time"
    Set mPrepRecordsetTime = myRs
End Function
Private Function mPrepRecordsetAll() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Rotation", adInteger
            .Append "Index", adInteger
            .Append "IsTimed", adBoolean
        End With
    myRs.Open
    myRs!Rotation.Properties("optimize") = True
    myRs.Sort = "Rotation desc"
    Set mPrepRecordsetAll = myRs
End Function
Private Function mPrepRecordsetTimeFacts() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Rotation", adInteger
            .Append "Index", adInteger
            .Append "Definition", adChar, 255
            '6762
            .Append "Stations", adChar, 600
            .Append "Files", adChar, 255
        End With
    myRs.Open
    myRs!Index.Properties("optimize") = True
    myRs.Sort = "index"
    Set mPrepRecordsetTimeFacts = myRs
End Function
Private Sub Class_Terminate()
    Set myService = Nothing
    If Not myRsSchedule Is Nothing Then
        If (myRsSchedule.State And adStateOpen) <> 0 Then
            myRsSchedule.Close
        End If
        Set myRsSchedule = Nothing
    End If
    If Not myRsTimeFacts Is Nothing Then
        If (myRsTimeFacts.State And adStateOpen) <> 0 Then
            myRsTimeFacts.Close
        End If
        Set myRsTimeFacts = Nothing
    End If
    Erase tmRegionalTime
    Erase tmRegionalInfo
 '   Erase dmScheduleDates

End Sub
Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function
Private Sub mTempRegionalInfoTime(myValue As CcsiToIdcRegionalInfo, slTemp() As CcsiToIdcRegionalInfo, slSites As String)
    'replace sites with the new values
    Dim ilCurrent As Integer
    
    If Len(slSites) > 0 Then
        ilCurrent = UBound(slTemp)
        Set slTemp(ilCurrent) = New CcsiToIdcRegionalInfo
        With slTemp(ilCurrent)
            .Definition = myValue.Definition
            .endDate = myValue.endDate
            .RegionalISCIs = myValue.RegionalISCIs
            .Sites = slSites
            .startDate = myValue.startDate
            .Rotation = myValue.Rotation
            .startTime = myValue.startTime
            .endTime = myValue.endTime
        End With
        ReDim Preserve slTemp(ilCurrent + 1)
    End If

End Sub
Private Sub mCreateTimeFacts(llTrigger As Long)
    Dim c As Integer
    '6411
    If Not myRsTimeFacts Is Nothing And llTrigger > 0 Then
        For c = 0 To UBound(tmAllTemp) - 1
            With tmAllTemp(c)
                myRsTimeFacts.AddNew Array("Index", "Rotation", "Definition", "Stations", "Files"), Array(llTrigger, .Rotation, .Definition, .Sites, .RegionalISCIs)
            End With
        Next c
    End If

End Sub
