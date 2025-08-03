Attribute VB_Name = "EngrPopSubs"
'
' Release: 1.0
'
' Description:
'   This file contains the Populate routines declarations
Option Explicit


'
'
'            gPopReportNmaes - build list of report names to show in user list box
'
'           'build in tgReportNames array:
'                   Report Name
'                   Report Index
'                   Crystal file name (ifhistory reports, only base name:
'                           ie  TimeType.rpt stored for TimeTypehist.rpt
'                   Report bmp name (or .jpg) for report sample
'                   Report description
'
Public Function gPopReportNames()

Dim ilUpper As Integer

    ReDim tgReportNames(0 To 0) As REPORTNAMES
    ilUpper = UBound(tgReportNames)
    tgReportNames(ilUpper).sRptName = "Material Type Names"
    tgReportNames(ilUpper).iRptIndex = MATTYPE_RPT
    tgReportNames(ilUpper).sCrystalName = "MatType"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Material Type names defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Relay Names"
    tgReportNames(ilUpper).iRptIndex = RELAY_RPT
    tgReportNames(ilUpper).sCrystalName = "Relay"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Relay Names defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "User Options"
    tgReportNames(ilUpper).iRptIndex = USER_RPT
    tgReportNames(ilUpper).sCrystalName = "User"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "Lists the options defined for each user"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Silence Names"
    tgReportNames(ilUpper).iRptIndex = SILENCE_RPT
    tgReportNames(ilUpper).sCrystalName = "Silence"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Silence Names defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Time Types"
    tgReportNames(ilUpper).iRptIndex = TIMETYPE_RPT
    tgReportNames(ilUpper).sCrystalName = "TimeType"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Time Types defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Follow Names"
    tgReportNames(ilUpper).iRptIndex = FOLLOW_RPT
    tgReportNames(ilUpper).sCrystalName = "Follow"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Follow Names defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Audio Names"
    tgReportNames(ilUpper).iRptIndex = AUDIONAME_RPT
    tgReportNames(ilUpper).sCrystalName = "AudName"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Audio Names defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Audio Types"
    tgReportNames(ilUpper).iRptIndex = AUDIOTYPE_RPT
    tgReportNames(ilUpper).sCrystalName = "AudType"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Audio Types defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Audio Sources"
    tgReportNames(ilUpper).iRptIndex = AUDIOSOURCE_RPT
    tgReportNames(ilUpper).sCrystalName = "AudSource"

    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Audio Sources defined "
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Site Options"
    tgReportNames(ilUpper).iRptIndex = SITE_RPT
    tgReportNames(ilUpper).sCrystalName = "Site"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of system options defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Bus Groups"
    tgReportNames(ilUpper).iRptIndex = BUSGROUP_RPT
    tgReportNames(ilUpper).sCrystalName = "BusGroup"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Bus Groups defined"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Bus Definitions"
    tgReportNames(ilUpper).iRptIndex = BUS_RPT
    tgReportNames(ilUpper).sCrystalName = "Bus"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Bus Definitions"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Netcue Names"
    tgReportNames(ilUpper).iRptIndex = NETCUE_RPT
    tgReportNames(ilUpper).sCrystalName = "Netcue"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Netcue names"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Control Names"
    tgReportNames(ilUpper).iRptIndex = CONTROL_RPT
    tgReportNames(ilUpper).sCrystalName = "Control"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of control names"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    ilUpper = ilUpper + 1
    
    tgReportNames(ilUpper).sRptName = "Comments"
    tgReportNames(ilUpper).iRptIndex = COMMENT_RPT
    tgReportNames(ilUpper).sCrystalName = "Comment"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of comments"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Event Types"
    tgReportNames(ilUpper).iRptIndex = EVENT_RPT
    tgReportNames(ilUpper).sCrystalName = "Event"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Events"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
     '6-8-11 remove activity report due to removing aie table
'    ilUpper = ilUpper + 1
'    tgReportNames(ilUpper).sRptName = "Change Activity Summary"
'    tgReportNames(ilUpper).iRptIndex = ACTIVITY_RPT
'    tgReportNames(ilUpper).sCrystalName = "ACTIVITY"
'    tgReportNames(ilUpper).sRptPicture = ""
'    tgReportNames(ilUpper).sRptDesc = "List of changes made by date and time"
'    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Automation"
    tgReportNames(ilUpper).iRptIndex = AUTOMATION_RPT
    tgReportNames(ilUpper).sCrystalName = "Automation"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of Automation systems"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Library Summary"
    tgReportNames(ilUpper).iRptIndex = LIBRARY_RPT
    tgReportNames(ilUpper).sCrystalName = "Library"         'full name = librarysum.rpt
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "Library Summary selectable by name, dates, and bus group"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Library Events"
    tgReportNames(ilUpper).iRptIndex = LIBRARYEVENT_RPT
    tgReportNames(ilUpper).sCrystalName = "Library"         'full name = librarydet.rpt
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "Library Events selectable by name and dates"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Audio Sources In-Use"
    tgReportNames(ilUpper).iRptIndex = AUDIOINUSE_RPT
    tgReportNames(ilUpper).sCrystalName = "AudioUse"         'full name = Audiouse.rpt
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of audio sources that are in-use or available, selectable by date, time, bus and audio type"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Template Summary"
    tgReportNames(ilUpper).iRptIndex = TEMPLATE_RPT
    tgReportNames(ilUpper).sCrystalName = "Template"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of template names and subnames"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Template Events"
    tgReportNames(ilUpper).iRptIndex = TEMPLATEEVENT_RPT
    tgReportNames(ilUpper).sCrystalName = "TemplateEvt"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of template names and subnames"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES
    
    ilUpper = ilUpper + 1
    tgReportNames(ilUpper).sRptName = "Template Air Info"
    tgReportNames(ilUpper).iRptIndex = TEMPLATEAIR_RPT
    tgReportNames(ilUpper).sCrystalName = "TemplateAir"
    tgReportNames(ilUpper).sRptPicture = ""
    tgReportNames(ilUpper).sRptDesc = "List of templates and their airing dates"
    ReDim Preserve tgReportNames(0 To ilUpper + 1) As REPORTNAMES

    
    
'    Dim iUpper As Integer
'    Dim sChar  As String * 1
'
'    On Error GoTo ErrHand
'
'    iUpper = 0
'    ReDim tgRnfInfo(0 To 0) As RNFINFO
'    SQLQuery = "SELECT rnfName, rnfRptExe, rnfCode"
'    sgSQLQuery = sgSQLQuery + " FROM RNF_Report_Name"
'    'sgSQLQuery = sgSQLQuery & " WHERE (rnfType = 'R') And (rnfName BEGINS WITH 'L' OR rnfName BEGINS WITH 'C')"
'    sgSQLQuery = sgSQLQuery & " WHERE (rnfType = 'R') And (rnfName Like 'L%' OR rnfName Like 'C%')"
'    sgSQLQuery = sgSQLQuery + " ORDER BY rnfName"
'    Set rst = cnn.OpenResultset(SQLQuery)
'    While Not rst.EOF
'        sChar = Mid$(rst!rnfName, 2, 1)
'        If (sChar >= "0") And (sChar <= "9") Then
'            tgRnfInfo(iUpper).iCode = rst!rnfCode
'            tgRnfInfo(iUpper).sName = rst!rnfName
'            tgRnfInfo(iUpper).sRptExe = rst!rnfRptExe
'            iUpper = iUpper + 1
'            ReDim Preserve tgRnfInfo(0 To iUpper) As RNFINFO
'        End If
'        rst.MoveNext
'    Wend
'
'    gPopReportNames = True
'    Exit Function
'ErrHand:
'    gMsg = ""
'    For Each gErrSQL In rdoErrors
'        If gErrSQL.Number <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsg = "A SQL error has occured in gPopReportNames: "
'            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number, vbCritical
'        End If
'    Next gErrSQL
'    If (Err.Number <> 0) And (gMsg = "") Then
'        gMsg = "A general error has occured in gPopReportNames: "
'        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
'    End If
'    gPopReportNames = False
'    Exit Function
    Exit Function
End Function


