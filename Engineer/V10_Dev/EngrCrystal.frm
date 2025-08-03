VERSION 5.00
Begin VB.Form EngrCrystal 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "EngrCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
' Copyright: Counterpoint Software, Inc.
' Date: September 2004
' Name: gCrystlReports
'
' The main entry point for all Crystal Reports. From here various subs get
' called to process the report.
'**************************************************************************
Public Sub gCrystlReports(sSQLString As String, iExportType As Integer, iRptDest As Integer, sRptName As String, sExpName As String)
    
    'gCrystlReports(sSQLString, iExportType, iRptDest, sRptName, sExpName)
    
    'sSQLString - SQL string created by report modules
    'iExportType - Current range is 1-10; Creates Exports to PDF, CSV, DIF, XLS etc.
    'iRptDest - Current Range 0-2; 0 = Display, 1 = Print, 2 = Export
    'sRptName - Name of the report - aflabels.rpt, afdelqvh.rpt etc.
    'sExpName - File Name used when creating disk file w/o extension
    
    Dim ilLoop As Integer
    Dim PathAndName As String
    Dim fNewForm As New EngrViewReport
    
    Dim crxTables As CRAXDRT.DatabaseTables
    Dim crxTable As CRAXDRT.DatabaseTable
    Dim crxSections As CRAXDRT.Sections
    Dim crxSection As CRAXDRT.Section
    Dim crxSubreportObj As CRAXDRT.SubreportObject
    Dim crxReportObjects As CRAXDRT.ReportObjects
    Dim crxSubreport As CRAXDRT.Report
    Dim ReportObject As Object
    Dim sTableName As String
    Dim ilLoopOnFile As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim slTempStr As String
    Dim ilConvertToSQL As Integer
    Dim ilIndexToTableLoc As Integer
    Dim crxConnectionProperties As CRAXDRT.ConnectionProperties
    Dim subConnecionProperties As CRAXDRT.ConnectionProperties
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass

    'Make the full path name and open the report
    PathAndName = sgReportDirectory + sRptName
    
    On Error GoTo ErrRptNotFoundHand
    Set fNewForm.Report = Appl.OpenReport(PathAndName, 1)
        
    On Error GoTo ErrHand
    fNewForm.Report.ReportTitle = sRptName
    
    'make sure odbc driver is used in case report developed under native btrieve
    ilConvertToSQL = False
    'If Trim$(fNewForm.Report.Database.Tables(1).DllName) <> "crdb_odbc.dll" Then
    '    fNewForm.Report.Database.ConvertDatabaseDriver "crdb_odbc.dll", True      '2-22-05 chg from p2sodbc to crdb_odbc.dll; convert to odbc from btrieve driver
    '    'Open the tables used by the report-need to set the locations parameters which is stored differnt than btrieve
    '    ilConvertToSQL = True
    '    For ilLoop = 1 To fNewForm.Report.Database.Tables.Count
    '        fNewForm.Report.Database.Tables(ilLoop).SetLogOnInfo sgDatabaseName, sgDatabaseName, "", ""
    '        sTableName = fNewForm.Report.Database.Tables(ilLoop).Name   'name of table
    '
    '        'look for the full name in valid array of filenames--if it exists from the DDF file, then
    '        'it isnt an alias table in the report.
    '        ilFound = mFindDDFTableName(sTableName)
    '        If ilFound Then
    '            fNewForm.Report.Database.Tables(ilLoop).Location = sgDatabaseName & "." & sTableName
    '        Else
    '            'valid file name not found.  This must be an alias table defined.
    '            'look for valid 1st 3or 4 character name, then pick up the associated full filename for the location definition
    '            ilFound = mFindAliasTableName(sTableName, ilIndexToTableLoc)
    '            If ilFound = True Then
    '                fNewForm.Report.Database.Tables(ilLoop).Location = sgDatabaseName & "." & Trim$(tgDDFFileNames(ilIndexToTableLoc).sLongName)
    '            Else
    '                MsgBox "Cannot set database location: gCrystlReports for " & sTableName
    '                Unload fNewForm
    '                Unload EngrViewReport
    '            End If
    '        End If
    '
    '
    '    Next ilLoop
    'Else
        'Open the tables used by the report.   Connection should propagate to the other tables
        Set crxConnectionProperties = fNewForm.Report.Database.Tables(1).ConnectionProperties
        With crxConnectionProperties
            .Item("DSN") = sgDatabaseName
        End With
    'End If
    Debug.Print fNewForm.Report.Database.Tables(1).ConnectionProperties.Item("DSN")
  
     'get the sections from the main report
    Set crxSections = fNewForm.Report.Sections
    
    'go through each section in the main report and find if a subreport exists
    For Each crxSection In crxSections
        'get all the objects in this section
        Set crxReportObjects = crxSection.ReportObjects
        'go throught eachobject in the report objects for this section
        For Each ReportObject In crxReportObjects
            'find the object which is the subreport object
            If ReportObject.Kind = crSubreportObject Then
                Set crxSubreportObj = ReportObject
                'open the subreport and treat it as any other report
                Set crxSubreport = crxSubreportObj.OpenSubreport
                
                For ilLoop = 1 To crxSubreport.Database.Tables.Count
  
                    'If ilConvertToSQL Then
                    '    crxSubreport.Database.ConvertDatabaseDriver "crdb_odbc.dll", True      '2-22-05 chg from p2sodbc to crdb_odbc.dll; convert to odbc from btrieve driver
                    '    crxSubreport.Database.Tables.Item(ilLoop).SetLogOnInfo sgDatabaseName, "", "", ""       'set the database name
                    '    crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & crxSubreport.Database.Tables(ilLoop).Name   'set the location, too
    
                    'Else
                        Set crxConnectionProperties = crxSubreport.Database.Tables(1).ConnectionProperties
                        With crxConnectionProperties
                            .Item("DSN") = sgDatabaseName

                        End With
                    'End If
                    
                    Debug.Print crxSubreport.Database.Tables(ilLoop).ConnectionProperties.Item("DSN")

                    sTableName = crxSubreport.Database.Tables(ilLoop).Name   'name of table
                   
                    'look for the full name in valid array of filenames--if it exists from the DDF file, then
                    'it isnt an alias table in the report.
                    ilFound = mFindDDFTableName(sTableName)
                    If ilFound Then
                    '    crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & sTableName
                    Else
                        'valid file name not found.  This must be an alias table defined.
                        'look for valid 1st 3or 4 character name, then pick up the associated full filename for the location definition
                        ilFound = mFindAliasTableName(sTableName, ilIndexToTableLoc)
                        If ilFound = True Then
                            'crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & Trim$(tgDDFFileNames(ilIndexToTableLoc).sLongName)
                        Else
                            MsgBox "Cannot set database location for subreport: gCrystlReports for " & sTableName
                            Unload fNewForm
                            Unload EngrViewReport
                        End If
                    End If
                Next ilLoop
            End If
        Next ReportObject
    Next crxSection
    
    
    'Pass the SQL String to Crystal
    fNewForm.Report.SQLQueryString = sSQLString
    
    'Set up the formulas for a given report
    mSetFormulas fNewForm
    
    'Without this do event you get errors when chosing destination printer
    DoEvents
    
   
    'MsgBox fNewForm.Report.Database.Tables(1).DllName      'check which database driver is used
    'Handle the report destination - Display, Print, Export
    Screen.MousePointer = vbDefault
    If iRptDest = 0 Then                   'Display Option
        fNewForm.Show igRptSource       'if coming from snspshot icon show modal,
                                        'if coming from report list show modeless
    ElseIf iRptDest = 1 Then               'Print Option
        fNewForm.Report.PrintOut False
    Else                                    'Export to File Options
        mCrystlExport fNewForm, iExportType, sExpName
    End If
 
    'If we export to anything other than Display we need to unload the form
    'and the object otherwise we have a memory leak.  The reason is show never gets
    'called so the the unload event never gets called.
    If iRptDest <> 0 Then
        Unload fNewForm
        Unload EngrViewReport
    End If
        
    
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in EngrCrystal - gCrystlReports: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
ErrRptNotFoundHand:
    Screen.MousePointer = vbDefault
    MsgBox "The file " & "'" & sRptName & "'" & " could not be found in the path: " & "'" & sgReportDirectory & "'" & sgCRLF & " Please check your Engineer.ini file for the correct Reports path.", vbCritical
    Exit Sub
    
End Sub

'**************************************************************************
' Copyright: Counterpoint Software, Inc. 2002
' Created by: Doug Smith
' Date: August 2002
' Name: mCrystlExport
'
' Calls the correct Crystal routine to export to disk the user chosen
' export type. PDF, Excel, etc.. Displays message to user about the path
' and file name that it exported to.
'**************************************************************************
Private Sub mCrystlExport(fNewForm As Form, iExportType As Integer, sExportName As String)
        
    ' mCrystlExport(iExportType, sExportName)
    
    ' Input parameters: fNewForm - a new instance of the form to be displayed
    '                   iExportType - allowable export types - PDF, CSV etc
    '                   sExportName - the disk file name to be used w/o its
    '                                 extension
    
    On Error GoTo ErrHand
    
    Select Case iExportType
        'change relative start of iExport from 1 to 0
        Case 0 'Adobe PDF
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".pdf"
            fNewForm.Report.ExportOptions.FormatType = crEFTPortableDocFormat
        Case 1 'Comma Seperated Values
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".csv"
            fNewForm.Report.ExportOptions.FormatType = crEFTCommaSeparatedValues
        Case 2 'Data Interchange
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".dif"
            fNewForm.Report.ExportOptions.FormatType = crEFTDataInterchange
        Case 3 'Excel 7
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".xls"
            fNewForm.Report.ExportOptions.FormatType = crEFTExcel70Tabular
        Case 4 'Excel 8
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".xls"
            fNewForm.Report.ExportOptions.FormatType = crEFTExcel80Tabular
        '3-12-04 Insert Text and move all other options down
        Case 5 'Text
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".txt"
            fNewForm.Report.ExportOptions.FormatType = crEFTText
        Case 6 'Rich Text
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".rtf"
            fNewForm.Report.ExportOptions.FormatType = crEFTRichText
        Case 7 'Tab Seperated Text
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".txt"
            fNewForm.Report.ExportOptions.FormatType = crEFTTabSeparatedText
        Case 8 'Paginated Text
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".txt"
            fNewForm.Report.ExportOptions.FormatType = crEFTPaginatedText
            fNewForm.Report.ExportOptions.NumberOfLinesPerPage = 60     '1-6-04
        Case 9 'Word For Windows
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".doc"
            fNewForm.Report.ExportOptions.FormatType = crEFTWordForWindows
        Case 10 'Crystal Reports 7.0
            fNewForm.Report.ExportOptions.DiskFileName = sgExportDirectory + sExportName + ".rpt"
            fNewForm.Report.ExportOptions.FormatType = crEFTCrystalReport70
    End Select
    fNewForm.Report.ExportOptions.DestinationType = crEDTDiskFile
    fNewForm.Report.Export False
           
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in EngrCrystal - mCrystlExport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub
'**************************************************************************
'       mSetformulas - Determine what formulas should be sent to each report
'           <input> - form for Crystal Viewer
'
'
'**************************************************************************
Private Sub mSetFormulas(fNewForm As Form)
Dim ilLoop As Integer

    ' mSetFormulas(fNewForm )
    ' Input parameters: fNewForm - a new instance of the form to be displayed
     
    'Dim CRXFormulaFields As CRAXDRT.FormulaFieldDefinitions     'collection of objects
    'Dim CRXFormulaField As CRAXDRT.FormulaFieldDefinition
    'Dim ilIndex As Integer
    
    On Error GoTo ErrHand
       
    DoEvents
    'Set CRXFormulaFields = fNewForm.Report.FormulaFields
    
    'find the matching report type and set its report caption for display mode only
    For ilLoop = 0 To UBound(tgReportNames) - 1
        If igRptIndex = tgReportNames(ilLoop).iRptIndex Then
            fNewForm.Caption = tgReportNames(ilLoop).sRptName & " Report"
            Exit For
        End If
    Next ilLoop
  
    'fNewForm.Caption = "User Options Report"

    'For Each CRXFormulaField In CRXFormulaFields
    '    If CRXFormulaField.Name = "{@PreparedBy}" Then
    '        CRXFormulaField.Text = "'" & Trim(tgUIE.sShowName) & "'"
    '    ElseIf CRXFormulaField.Name = "{@IncludeHistory}" Then
    '        CRXFormulaField.Text = Trim(sgCrystlFormula1)
    '    End If
    'Next
        
    'do more formulas if required in the report
    Select Case igRptIndex
        Case RELAY_RPT, SILENCE_RPT, TIMETYPE_RPT, NETCUE_RPT, MATTYPE_RPT, FOLLOW_RPT, AUDIOTYPE_RPT, AUDIONAME_RPT, BUS_RPT, BUSGROUP_RPT, AUDIOSOURCE_RPT
            'start/end date spans for history
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
        Case CONTROL_RPT, COMMENT_RPT, CONTROL_RPT, USER_RPT, EVENT_RPT
            'start/end date spans for history
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
        Case SITE_RPT
             'EngrViewReport.CRViewer1.DisplayGroupTree = True
             'EngrViewReport.CRViewer1.EnableGroupTree = True
             
        Case ACTIVITY_RPT
            'pass start & end dates & times to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
            mFindFormula fNewForm, "{@FromTime}", "'" & Trim(sgCrystlFormula4) & "'"
            mFindFormula fNewForm, "{@ToTime}", "'" & Trim(sgCrystlFormula5) & "'"
        Case AUTOMATION_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
            'pass whether to show automation grid.  Only show them in report if
            'signed in as Counterpoint or Guide with RadioXXXX  password
            If ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 5)) Or _
               (((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 4))) Then
                sgCrystlFormula4 = "'Y'"
            Else
                sgCrystlFormula4 = "'N'"
            End If
            mFindFormula fNewForm, "{@ShowGrid}", sgCrystlFormula4
            If sgClientFields = "A" Then            'ABC
                sgCrystlFormula5 = "'Y'"
            Else
                sgCrystlFormula5 = "'N'"
            End If
            mFindFormula fNewForm, "{@ABCCustomFields}", sgCrystlFormula5
        Case LIBRARY_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
            mFindFormula fNewForm, "{@SortBy}", sgCrystlFormula4       'sort by library or bus
        Case LIBRARYEVENT_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
            mFindFormula fNewForm, "{@DaysRequested}", "'" & Trim$(sgCrystlFormula4) & "'"            'valid days requested for header
        Case AUDIOINUSE_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
        Case ITEMIDCHK_RPT
            mFindFormula fNewForm, "{@WhichOption}", sgCrystlFormula2       'by date or item id
            mFindFormula fNewForm, "{@DiscrepOnly}", sgCrystlFormula3       'Y = discrep only, else N
        Case SCHED_RPT
            mFindFormula fNewForm, "{@DaysRequested}", sgCrystlFormula2       'dates for heading
            mFindFormula fNewForm, "{@ShowFilter}", sgCrystlFormula3               ' user filters selected
        Case TEXT_RPT               'print system messages
            mFindFormula fNewForm, "{@MessageFile}", sgCrystlFormula2       'message file name for heading
        Case TEMPLATE_RPT
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
        Case TEMPLATEAIR_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@FromDate}", sgCrystlFormula2
            mFindFormula fNewForm, "{@ToDate}", sgCrystlFormula3
            mFindFormula fNewForm, "{@DaysRequested}", "'" & Trim$(sgCrystlFormula4) & "'"            'valid days requested for header
        Case ASAIRCOMPARE_RPT
            'pass start & end dates to show in report header
            mFindFormula fNewForm, "{@DaysRequested}", sgCrystlFormula2       'dates for heading
            mFindFormula fNewForm, "{@ShowFilter}", sgCrystlFormula3               ' user filters selected
    End Select
    'Pass to all reports who requested the report
    mFindFormula fNewForm, "{@PreparedBy}", "'" & Trim(tgUIE.sShowName) & "'"

    Exit Sub
ErrHand:
        Screen.MousePointer = vbDefault
        gMsg = ""
        If (Err.Number <> 0) And (gMsg = "") Then
            gMsg = "A general error has occured in EngrCrystal - mSetFormulas: "
            MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload EngrCrystal
End Sub
'
'       mFindFormula - find the matching formula name in the crystal report
'       and set the values for viewing
'
'       <input> slFormulaName - formula name to match
'               slFormulaText - formula value to set
Public Sub mFindFormula(fNewForm As Form, slFormulaName As String, slFormulaText As String)
Dim ilLoop As Integer
Dim CRXFormulaFields As CRAXDRT.FormulaFieldDefinitions     'collection of objects
Dim CRXFormulaField As CRAXDRT.FormulaFieldDefinition

    Set CRXFormulaFields = fNewForm.Report.FormulaFields

    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = Trim(slFormulaName) Then
            CRXFormulaField.text = Trim(slFormulaText)
            Exit For
        End If
    Next
    Exit Sub
End Sub

Public Sub gActiveCrystalReports(iExportType As Integer, iRptDest As Integer, slRptName As String, sExpName As String, rstActive As ADODB.Recordset)
Dim fNewForm As New EngrViewReport
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crxSubreportObj As CRAXDRT.SubreportObject
Dim crxReportObjects As CRAXDRT.ReportObjects
Dim crxSubreport As CRAXDRT.Report
Dim ReportObject As Object
Dim ilLoop As Integer
Dim slTableName As String
Dim ilFound As Integer
Dim ilIndexToTableLoc As Integer
Dim subConnectionProperties As CRAXDRT.ConnectionProperties


    Set fNewForm.Report = Appl.OpenReport(sgReportDirectory + slRptName)
    fNewForm.Report.Database.Tables(1).SetDataSource rstActive, 3
    
    'Without this do event you get errors when chosing destination printer
     DoEvents
    
     'set locations to the subreports
    'get the sections from the main report
    Set crxSections = fNewForm.Report.Sections
    
    'go through each section in the main report and find if a subreport exists
    For Each crxSection In crxSections
        'get all the objects in this section
        Set crxReportObjects = crxSection.ReportObjects
        'go throught eachobject in the report objects for this section
        For Each ReportObject In crxReportObjects
            'find the object which is the subreport object
            If ReportObject.Kind = crSubreportObject Then
                Set crxSubreportObj = ReportObject
                'open the subreport and treat it as any other report
                Set crxSubreport = crxSubreportObj.OpenSubreport
                
                For ilLoop = 1 To crxSubreport.Database.Tables.Count
                    
                    'crxSubreport.Database.Tables.Item(ilLoop).SetLogOnInfo sgDatabaseName, "", "", ""       'set the database name
                    'crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & crxSubreport.Database.Tables(ilLoop).Name  'set the location, too
                    Set subConnectionProperties = crxSubreport.Database.Tables(1).ConnectionProperties
                    With subConnectionProperties
                        .Item("DSN") = sgDatabaseName
                    End With
                    slTableName = crxSubreport.Database.Tables(ilLoop).Name   'name of table
                   
                    'look for the full name in valid array of filenames--if it exists from the DDF file, then
                    'it isnt an alias table in the report.
                    ilFound = mFindDDFTableName(slTableName)
                    If ilFound Then
                        'crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & slTableName
                    Else
                        'valid file name not found.  This must be an alias table defined.
                        'look for valid 1st 3or 4 character name, then pick up the associated full filename for the location definition
                        ilFound = mFindAliasTableName(slTableName, ilIndexToTableLoc)
                        If ilFound = True Then
                        '    crxSubreport.Database.Tables(ilLoop).Location = sgDatabaseName & "." & Trim$(tgDDFFileNames(ilIndexToTableLoc).sLongName)
                        Else
                            MsgBox "Cannot set database location for subreport: gCrystlReports for " & slTableName
                            Unload fNewForm
                            Unload EngrViewReport
                        End If
                    End If
                Next ilLoop
            End If
        Next ReportObject
    Next crxSection
    
    'Set up the formulas for a given report
    mSetFormulas fNewForm
  
    'MsgBox fNewForm.Report.Database.Tables(1).DllName      'check which database driver is used
    'Handle the report destination - Display, Print, Export
    Screen.MousePointer = vbDefault
    If iRptDest = 0 Then                   'Display Option
        fNewForm.Show igRptSource       'if coming from snspshot icon show modal,
                                        'if coming from report list show modeless
    ElseIf iRptDest = 1 Then               'Print Option
        DoEvents
        fNewForm.Report.PrintOut False
    
    Else                                    'Export to File Options
        mCrystlExport fNewForm, iExportType, sExpName
    End If
 
    'If we export to anything other than Display we need to unload the form
    'and the object otherwise we have a memory leak.  The reason is show never gets
    'called so the the unload event never gets called.
    If iRptDest <> 0 Then
        Unload fNewForm
        Unload EngrViewReport
    End If
    
    'fNewForm.Show igRptSource   'vbModeless
   
End Sub
'
'
'           mFindDDFTableName - find the matching table name from DDFs
'           so that a database location can be set
'
'           <input> full table name from crystal
'           <output> none
'           return - true if a table name has been found
'
Public Function mFindDDFTableName(slTableName As String) As Integer
Dim ilLoopOnFile As Integer
Dim ilFound As Integer

    'look for the full name in valid array of filenames--if it exists from the DDF file, then
    'it isnt an alias table in the report.
    ilFound = False
    For ilLoopOnFile = LBound(tgDDFFileNames) To UBound(tgDDFFileNames) - 1
        If Trim$(slTableName) = Trim$(tgDDFFileNames(ilLoopOnFile).sLongName) Then
            ilFound = True
            Exit For
        End If
    Next ilLoopOnFile
        mFindDDFTableName = ilFound
End Function
'
'
'       find the base table name (i.e. for alias file naming:  uie.mkd
'           has been aliased to uie_to.  By finding the base table name,
'           the location can be set for the table
'
'           <Input> slTableName - Full alias name
'           <output> ilIndexToTableLoc - index to the base table name
'           Return - true  if the base table name found
'
Public Function mFindAliasTableName(slTableName As String, ilIndexToTableLoc As Integer) As Integer
Dim ilLoopOnFile As Integer
Dim slTempStr As String
Dim ilRet As Integer
Dim ilFound As Integer
    
    ilFound = False
    'valid file name not found.  This must be an alias table defined.
    'look for valid 1st 3or 4 character name, then pick up the associated full filename for the location definition
    For ilLoopOnFile = LBound(tgDDFFileNames) To UBound(tgDDFFileNames) - 1
        slTempStr = RTrim$(tgDDFFileNames(ilLoopOnFile).sShortName)
        ilRet = InStr(1, slTableName, Trim$(slTempStr))
        If ilRet > 0 Then
            'fNewForm.Report.Database.Tables(ilLoop).Location = sgDatabaseName & "." & Trim$(tgDDFFileNames(ilLoopOnFile).sLongName)
            ilFound = True
            ilIndexToTableLoc = ilLoopOnFile
            Exit For
        End If
    Next ilLoopOnFile
    mFindAliasTableName = ilFound
End Function
