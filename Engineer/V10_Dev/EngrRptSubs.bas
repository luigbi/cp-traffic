Attribute VB_Name = "EngrRptSubs"
Public rstLibrary As ADODB.Recordset
Public rstLibEvts As ADODB.Recordset
Public rstAIERpt As ADODB.Recordset
Public rstItemIDRpt As ADODB.Recordset
Public rstSchedRpt As ADODB.Recordset
Public rstTextrpt As ADODB.Recordset

Public Sub gGenerateRstText()
    rstTextrpt.Fields.Append "Text", adChar, 255            'one string of text message
End Sub
Public Sub gGenerateItemIDRst()
    rstItemIDRpt.Fields.Append "Date", adDBDate, 4      'date of of itemid check
    rstItemIDRpt.Fields.Append "SelectItemID", adChar, 32     'selective item id
    rstItemIDRpt.Fields.Append "DiscrepOnly", adChar, 1  'Y = discrep only, else N
    rstItemIDRpt.Fields.Append "ItemID", adChar, 32         'record item id
    rstItemIDRpt.Fields.Append "Title", adChar, 35          'short title
    rstItemIDRpt.Fields.Append "PrimProt", adChar, 4        'pri or protection
    rstItemIDRpt.Fields.Append "ReturnTitle", adChar, 35    'returned short title
End Sub

Public Sub gGenerateRstAIERpt()
    rstAIERpt.Fields.Append "TypeOfChange", adChar, 20      'Library, Events, Lists etc
    rstAIERpt.Fields.Append "Time", adDBTime, 4              'time of change
    rstAIERpt.Fields.Append "Date", adDBDate, 4        'date of change
    rstAIERpt.Fields.Append "User", adChar, 30              'user
    rstAIERpt.Fields.Append "DescKeyField", adChar, 80 'description of key field changed:  i.e. library names/subnames, list names
    rstAIERpt.Fields.Append "DescSecField", adChar, 80  'description of secondary key field changed: i.e. addl library evnt or sched info
End Sub
Public Sub gGeneraterstLibrary()
    rstLibrary.Fields.Append "Name", adChar, 20
    rstLibrary.Fields.Append "Subname", adChar, 20
    rstLibrary.Fields.Append "Desc", adChar, 66
    rstLibrary.Fields.Append "StartDate", adChar, 10
    rstLibrary.Fields.Append "EndDate", adChar, 10
    rstLibrary.Fields.Append "StartdateSort", adInteger, 4      'date used for sorting within library name
    rstLibrary.Fields.Append "Days", adChar, 7
    rstLibrary.Fields.Append "StartTime", adChar, 10            'time displacement in miliseconds
    rstLibrary.Fields.Append "Length", adChar, 10
    rstLibrary.Fields.Append "Hours", adChar, 50
    rstLibrary.Fields.Append "BusGroup", adChar, 20
    rstLibrary.Fields.Append "Bus", adChar, 20
    rstLibrary.Fields.Append "State", adChar, 1
    rstLibrary.Fields.Append "Version", adInteger, 2
    rstLibrary.Fields.Append "SubVersion", adInteger, 2
    rstLibrary.Fields.Append "Sequence", adInteger, 2
    rstLibrary.Fields.Append "DateEntered", adDBDate, 4        'date entered for history
    rstLibrary.Fields.Append "User", adChar, 20                 'user entered for history
    rstLibrary.Fields.Append "TimeEntered", adDBTime, 4        'time entered for history
    rstLibrary.Fields.Append "FileCode", adInteger, 4           'original summary header for history sorting
    rstLibrary.Fields.Append "OrigAIECode", adInteger, 4        'original AIE code to tie the current/previous versions together
    rstLibrary.Fields.Append "IgnoreConflicts", adChar, 1       'ignore conflicts flag (ignore B: bus, A:Audio, I:both, N:flag the conflicts)
    rstLibrary.Fields.Append "DHEHeaderCode", adInteger, 4      'DHE code
End Sub
'
'       DDFs for Library events report and Events Snapshot
'
'       3-31-06 Added event description as string for Snapshot
'               Added event buses as string for snapshot

Public Sub gGeneraterstLibEvts()
    rstLibEvts.Fields.Append "Name", adChar, 20
    rstLibEvts.Fields.Append "Subname", adChar, 20
    rstLibEvts.Fields.Append "DescCteCode", adInteger, 4
    rstLibEvts.Fields.Append "StartDate", adChar, 10
    rstLibEvts.Fields.Append "StartDateSort", adInteger, 4  'start date for sorting when libraray names the same
    rstLibEvts.Fields.Append "EndDate", adChar, 10
    rstLibEvts.Fields.Append "StartTime", adChar, 10
    rstLibEvts.Fields.Append "Length", adChar, 10
    rstLibEvts.Fields.Append "State", adChar, 1
    rstLibEvts.Fields.Append "EventType", adChar, 1     'event type (p= pgm, a = avail)
    rstLibEvts.Fields.Append "EvBusDeeCode", adInteger, 4   'reference to 1 to many buses (ebe to dee)
    rstLibEvts.Fields.Append "EvbusCtl", adChar, 1      'bus control
    rstLibEvts.Fields.Append "EvStarttime", adChar, 10  'event start time displacement
    rstLibEvts.Fields.Append "EvStartTimeSort", adInteger, 4    'start time for sorting in crystal
    rstLibEvts.Fields.Append "EvStartType", adChar, 3
    rstLibEvts.Fields.Append "EvFix", adChar, 1
    rstLibEvts.Fields.Append "EvEndType", adChar, 3
    rstLibEvts.Fields.Append "EvDur", adChar, 10         'event duration in 10th of seconds
    rstLibEvts.Fields.Append "EvDays", adChar, 14       'days to air:  M-F Ss-Su
    rstLibEvts.Fields.Append "EvHours", adChar, 50      '3-4-09 chged from 28 to 50 ,hours to air: 5-10 3-7
    rstLibEvts.Fields.Append "EvMatType", adChar, 3
    rstLibEvts.Fields.Append "EvAudName1", adChar, 8    'primary audio name
    rstLibEvts.Fields.Append "EvItem1", adChar, 32      'primary item name
    'rstLibEvts.Fields.Append "EvISCI1", adChar, 20      'primary ISCI
    rstLibEvts.Fields.Append "EvCtl1", adChar, 1         'primary control char"
    rstLibEvts.Fields.Append "EvAudName2", adChar, 8    'backup audio name
    rstLibEvts.Fields.Append "EvCtl2", adChar, 1        'backup control char
    rstLibEvts.Fields.Append "EvAudName3", adChar, 8    'protection audio name
    rstLibEvts.Fields.Append "Evitem3", adChar, 32      'protection item name
    'rstLibEvts.Fields.Append "EvISCI3", adChar, 20      'protection item name
    rstLibEvts.Fields.Append "EvCtl3", adChar, 1        'protection control char
    rstLibEvts.Fields.Append "EvRelay1", adChar, 8         'relay 1
    rstLibEvts.Fields.Append "EvRelay2", adChar, 8          'relay 2
    rstLibEvts.Fields.Append "EvFollow", adChar, 19
    rstLibEvts.Fields.Append "EvTime", adChar, 10
    rstLibEvts.Fields.Append "EvSilence1", adChar, 1
    rstLibEvts.Fields.Append "EvSilence2", adChar, 1
    rstLibEvts.Fields.Append "EvSilence3", adChar, 1
    rstLibEvts.Fields.Append "EvSilence4", adChar, 1
    rstLibEvts.Fields.Append "EvNetCue1", adChar, 3     'start net cue
    rstLibEvts.Fields.Append "EvNetCue2", adChar, 3     'stop net cue
    rstLibEvts.Fields.Append "EvTitle1CteCode", adInteger, 4 'Title 1 comment
    rstLibEvts.Fields.Append "EvTitle2CteCode", adInteger, 4 'Title 2 comment
    '1-4-05 added for history
    rstLibEvts.Fields.Append "Version", adInteger, 2
    rstLibEvts.Fields.Append "SubVersion", adInteger, 2
    rstLibEvts.Fields.Append "Sequence", adInteger, 2
    rstLibEvts.Fields.Append "DateEntered", adDBDate, 4        'date entered for history
    rstLibEvts.Fields.Append "User", adChar, 20                 'user entered for history
    rstLibEvts.Fields.Append "TimeEntered", adDBTime, 4        'time entered for history
    rstLibEvts.Fields.Append "FileCode", adInteger, 4           'original summary header for history sorting
    rstLibEvts.Fields.Append "OrigAIECode", adInteger, 4
    '2-21-06 Ignore conflicts flag
    rstLibEvts.Fields.Append "IgnoreConflicts", adChar, 1        'ignore conflicts flag (ignore B: bus, A:Audio, I:both, N:flag the conflicts)
    '3-31-06  For Library events snapshot, everything in strings.  Pointers not used since user may change entry
    rstLibEvts.Fields.Append "EvBusDesc", adChar, 40               'bus descriptions (i.e. B,C,J2,K)
    rstLibEvts.Fields.Append "EvTitle1Desc", adChar, 66
    rstLibEvts.Fields.Append "EvTitle2Desc", adChar, 90
    rstLibEvts.Fields.Append "DHEHeaderCode", adInteger, 4          'Library header code
    'rstLibEvts.Fields.Append "EvABCFormat", adChar, 1      'ABC custom defined fields
    'rstLibEvts.Fields.Append "EvABCPgmCode", adChar, 25      'ABC custom defined fields
    'rstLibEvts.Fields.Append "EvABCXDSMode", adChar, 2      'ABC custom defined fields
    'rstLibEvts.Fields.Append "EvABCRecordItem", adChar, 1      'ABC custom defined fields
    rstLibEvts.Fields.Append "EvABCCustomFields", adChar, 100   'ABC custom defined fields, concatenated
End Sub
'
'       Build the field definitions for Schedule report
'       Generate Snapshot of Schedule screen
'
Public Sub gGenerateRstSchedule()
    rstSchedRpt.Fields.Append "StartDate", adChar, 10
    rstSchedRpt.Fields.Append "StartDateSort", adInteger, 4  'start date for sorting when libraray names the same
    rstSchedRpt.Fields.Append "EventType", adChar, 1     'event type (p= pgm, a = avail)
    rstSchedRpt.Fields.Append "Event ID", adChar, 10     'event ID
    rstSchedRpt.Fields.Append "EvBusName", adChar, 8   'reference to 1 to many buses (ebe to dee)
    rstSchedRpt.Fields.Append "EvbusCtl", adChar, 1      'bus control
    rstSchedRpt.Fields.Append "EvStarttime", adChar, 10  'event start time displacement
    rstSchedRpt.Fields.Append "EvStartTimeSort", adInteger, 4    'start time for sorting in crystal
    rstSchedRpt.Fields.Append "EvStartType", adChar, 3
    rstSchedRpt.Fields.Append "EvFix", adChar, 1
    rstSchedRpt.Fields.Append "EvEndType", adChar, 3
    rstSchedRpt.Fields.Append "EvDur", adChar, 10         'event duration in 10th of seconds
    rstSchedRpt.Fields.Append "EvMatType", adChar, 3
    rstSchedRpt.Fields.Append "EvAudName1", adChar, 8    'primary audio name
    rstSchedRpt.Fields.Append "EvItem1", adChar, 32      'primary item name
    rstSchedRpt.Fields.Append "EvCtl1", adChar, 1         'primary control char"
    rstSchedRpt.Fields.Append "EvAudName2", adChar, 8    'backup audio name
    rstSchedRpt.Fields.Append "EvCtl2", adChar, 1        'backup control char
    rstSchedRpt.Fields.Append "EvAudName3", adChar, 8    'protection audio name
    rstSchedRpt.Fields.Append "Evitem3", adChar, 32      'protection item name
    rstSchedRpt.Fields.Append "EvCtl3", adChar, 1        'protection control char
    rstSchedRpt.Fields.Append "EvRelay1", adChar, 8         'relay 1
    rstSchedRpt.Fields.Append "EvRelay2", adChar, 8          'relay 2
    rstSchedRpt.Fields.Append "EvFollow", adChar, 19
    rstSchedRpt.Fields.Append "EvSilenceTime", adChar, 10
    rstSchedRpt.Fields.Append "EvSilence1", adChar, 1
    rstSchedRpt.Fields.Append "EvSilence2", adChar, 1
    rstSchedRpt.Fields.Append "EvSilence3", adChar, 1
    rstSchedRpt.Fields.Append "EvSilence4", adChar, 1
    rstSchedRpt.Fields.Append "EvNetCue1", adChar, 3     'start net cue
    rstSchedRpt.Fields.Append "EvNetCue2", adChar, 3     'stop net cue
    rstSchedRpt.Fields.Append "EvTitle1", adChar, 66   'Title 1 comment
    rstSchedRpt.Fields.Append "EvTitle2", adChar, 90   'Title 2 comment
    
    '2-20-07 Add new fields for AsAir Compare report
    rstSchedRpt.Fields.Append "AirDateError", adChar, 1     'Air Date Error
    rstSchedRpt.Fields.Append "AutoOffError", adChar, 1     'Auto off error
    rstSchedRpt.Fields.Append "DataError", adChar, 1        'date error
    rstSchedRpt.Fields.Append "ScheduleError", adChar, 1    'Schedule error
    rstSchedRpt.Fields.Append "SequenceID", adInteger, 4    'sequence #, for AsAir compare report, must be in same order as the screen output
    'rstSchedRpt.Fields.Append "EvABCFormat", adChar, 1      'ABC custom defined fields
    'rstSchedRpt.Fields.Append "EvABCPgmCode", adChar, 25      'ABC custom defined fields
    'rstSchedRpt.Fields.Append "EvABCXDSMode", adChar, 2      'ABC custom defined fields
    'rstSchedRpt.Fields.Append "EvABCRecordItem", adChar, 1      'ABC custom defined fields
    rstSchedRpt.Fields.Append "EvABCCustomFields", adChar, 100   'abc custom defined field, concatenated

End Sub

'
'       Build the list of all valid export types form Crystal
'
Public Sub gPopExportTypes(cboFileType As Control)
    
    cboFileType.AddItem "Acrobat PDF"
    cboFileType.AddItem "Comma separated value"
    cboFileType.AddItem "Data Interchange"
    cboFileType.AddItem "Excel 7"
    cboFileType.AddItem "Excel 8"
    cboFileType.AddItem "Text"
    cboFileType.AddItem "Rich Text"
    cboFileType.AddItem "Tab separated text"
    cboFileType.AddItem "Paginated Text"
    cboFileType.AddItem "Word for Windows"
    cboFileType.AddItem "Crystal Reports"
End Sub
'
'
'           gChangeCaption - change caption on the report
'           selectivity screen so user knows which report
'           screen is visible
'
'           <input/output>  control that needs to setup the caption
'
'           igRptIndex (global variable) contains the report index to process
Public Sub gChangeCaption(RptOption As Control)
Dim ilLoop As Integer
    For ilLoop = 0 To UBound(tgReportNames) - 1
        If igRptIndex = tgReportNames(ilLoop).iRptIndex Then
            RptOption = tgReportNames(ilLoop).sRptName & " Report Selectivity"     'set the export file name
            Exit For
        End If
    Next ilLoop
    Exit Sub
End Sub
'
'
'           gObtainReportforCrystal - retrieved the Crystal .rpt name
'           to call Crystal
'           <output> slCrystalReportName as string
'                    slExportName as string
'
'           igRptIndex (global variable) contains the report index to process
Public Sub gObtainReportforCrystal(slCrystalReportName As String, slExportName As String)
Dim ilSpecialNeeds As Integer
    slCrystalReportName = ""
    slExportName = ""
    ilSpecialNeeds = False
    If igRptIndex = AUDIOTYPE_RPT Then
        If EngrUserRpt!optOldNew(0).Value = True And EngrUserRpt!ckcInclOther.Value = vbChecked Then
            ilSpecialNeeds = True
            slCrystalReportName = "AudTypeSrc"
            slExportName = "AudTypeSrcRpt"
        End If
    ElseIf igRptIndex = BUSGROUP_RPT Then
        If EngrUserRpt!optOldNew(0).Value = True And EngrUserRpt!ckcInclOther.Value = vbChecked Then
            ilSpecialNeeds = True
            slCrystalReportName = "BusGroupDef"
            slExportName = "BusGroupDefRpt"
        End If
    
    End If
    If Not ilSpecialNeeds Then
        For ilLoop = 0 To UBound(tgReportNames) - 1
            If igRptIndex = tgReportNames(ilLoop).iRptIndex Then
                slCrystalReportName = Trim$(tgReportNames(ilLoop).sCrystalName)           'crystal report name to call
                slExportName = Trim$(tgReportNames(ilLoop).sCrystalName) & "rpt"        'set the export file name
                Exit For
            End If
        Next ilLoop
    End If
    Exit Sub
End Sub
'
'
'           gEditDateInput - edit the date entered and default to earliest
'           or latest possible date if blank
'           <input>  slInput - date entered
'                    slDateDefault - earliest (1/1/1970) or latest (12/31/2069) possible date if left blank
'           <return>  datedefault or valid date entered.  Return blank if invalid
'
Public Function gEditDateInput(slInput As String, slDateDefault) As String
Dim slSQLDate As String
Dim slDate As String

    slDate = slInput
    If slDate = "" Then
        slSQLDate = slDateDefault
    Else
        If Not gIsDate(slDate) Then
            Screen.MousePointer = vbDefault
            Beep
            MsgBox "Invalid Date"
            gEditDateInput = ""     'invalid, dont return a date
            Exit Function
        End If
        slSQLDate = slDate
    End If
    gEditDateInput = slSQLDate
End Function
'
'
'           gFormatDays - format the days of the week requested based on the start & end dates
'                         and the days of the week selected.
'           i.e.  if more than one week requested, determine days of week from user day selection.
'                 if less than one week requested, turn all days off not requested
'           <input> slStartDate - requested start date
'                   slEndDate - requeted end date
'                   ckcDays() - array of 7 days, user selected check boxes on/off
'           <return> formatted days of week selected (i.e. M-SU, M-TU, SaSU)
Public Function gFormatDays(slStartDate As String, slEndDate As String, ckcDays() As Integer) As String
Dim ilLoop As Integer
Dim ilTempDays(0 To 6) As Integer       'Days of the week check box answers
Dim llStartDate As Long
Dim llEndDate As Long
Dim llDate As Long
Dim ilWeekDay As Integer
Dim slStr As String

        gFormatDays = ""
        For ilLoop = 0 To 6
            ilTempDays(ilLoop) = vbUnchecked
        Next ilLoop
        llStartDate = gDateValue(slStartDate)
        llEndDate = gDateValue(slEndDate)
        If llEndDate - llStartDate > 6 Then
            llEndDate = llStartDate + 6
        End If
        
        For llDate = llStartDate To llEndDate
            ilWeekDay = Weekday(llDate) - 2
            If ilWeekDay < 0 Then
                ilWeekDay = 6
            End If
            ilTempDays(ilWeekDay) = ckcDays(ilWeekDay)
        Next llDate
        
        slStr = ""
        For ilLoop = 0 To 6
            If ilTempDays(ilLoop) = vbChecked Then
                slStr = Trim$(slStr) & "Y"
            Else
                slStr = Trim$(slStr) & "N"
            End If
        Next ilLoop
        slStr = gDayMap(slStr)
        gFormatDays = slStr
        Exit Function
End Function
