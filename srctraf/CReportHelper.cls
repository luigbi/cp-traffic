VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************CReportHelper******************************************************
'dependent upon:
    ' pdfsplitmerge, craxdrt library 11.5, microsoft scripting runtime, microsoft activex 2.8,  cLogoStandard,ogPdf
'public methods: OpenReport  OpenAndSetDataSource SetFormula  SetSelection ModifyLogo PrintOut Export ReadRecords RemoveReport
'public properties: Logos Reports UserWantsToQuit TreatAsLastReport System iLastPrintJob CurrentReportName AlternateLogo ReportBin
Option Explicit
Private omReportApplication As CRAXDRT.Application
Public Logos As Dictionary
Public Reports As Dictionary
Public LogoPath As String
Private imLastPrintJob As Integer
Public UserWantsToQuit As Boolean
Private smReportName As String
Private omReportSource As ReportSourceRouter
Private imNumberForSameName As Integer      'reconciliation may call same report more than once.  I append number onto report name to add to dictionary
Private omLogoCondition As IConditionalLogo
'property below won't be used much.  This was for invoice report, which called rptSelIn multiple times where # of jobs was always 1.  If I just set ilastprintjob to 2, then 'open' couldn't have 'safety' of 'if last = 1, then clear out any old reports in collection
Public TreatAsLastReport As Boolean     '
'make affiliate and traffic have same variable names
'use from myconnection for dsnName, so don't have to look in ini file for traffic
Private smGlobalReportPath As String
Private smGlobalDatabasePath As String
Private smGlobalReportSavePath As String           '8-29-16 make this variable global; as the output path will be changed in export when generating separate pdfs by agency

Private myConnection As CCsiSystemConnection
Private smSystem As SystemType
'for clear channel issue: all reports need to be odbc.  User must have a dsn name!
'currently, this is never set to true
Private bmForceToOdbc As Boolean
Private omPdf As PDFSplitMerge.CPDFSplitMergeObj

Public PdfPageFirst As Integer
Public PdfPageLast As Integer
Public AddToSelection As String


Private Sub Class_Initialize()
    imLastPrintJob = 1
    Set Reports = New Dictionary
    Set Logos = New Dictionary
    Set omLogoCondition = New CLogoStandard
    Set myConnection = New CCsiSystemConnection
'    Set ogPdf = New PDFSplitMerge.CPDFSplitMergeObj
'    ogPdf.SetCode "989888DFE6F0FF0056681A"

End Sub

Private Sub Class_Terminate()
    Set omReportApplication = Nothing
    Set Logos = Nothing
    Set Reports = Nothing
    Set omReportSource = Nothing
    Set omLogoCondition = Nothing
    Set myConnection = Nothing
'    Set ogPdf = Nothing
End Sub
Public Property Get System() As SystemType
    System = myConnection.System
End Property
Public Property Get iLastPrintJob() As Integer
    iLastPrintJob = imLastPrintJob
End Property
Public Property Let iLastPrintJob(ilNumberOfJobs As Integer)
    If ilNumberOfJobs > 0 Then
        imLastPrintJob = ilNumberOfJobs
    End If
End Property
Public Property Let DiscardSavedData(blDiscard As Boolean)
    If blDiscard And Len(smReportName) > 0 Then
        Reports.Item(smReportName).DiscardSavedData
    End If
End Property
Public Property Get CurrentReportName() As String
    CurrentReportName = smReportName
End Property
Public Property Get PDFPathName() As String
    PDFPathName = smGlobalReportSavePath
End Property
Public Property Let PDFPathName(slNewPDFPath As String)
    smGlobalReportSavePath = slNewPDFPath
End Property
Public Property Get CurrentReportObject() As CRAXDRT.Report
    If Not smReportName = "" Then
        Set CurrentReportObject = Reports.Item(smReportName)
    Else
        MsgBox "Please OpenReport before attempting to get current report", vbExclamation + vbOKOnly
    End If
End Property
Public Property Set Alternatelogo(ilConditionClass As IConditionalLogo)
    Set omLogoCondition = ilConditionClass
    Set ilConditionClass = Nothing
End Property
Public Property Get ReportBin() As ReportSourceRouter
'reportSourceRouter passes all reports in collection at one time.  This passes those reports to crViewer
    Dim vlKey As Variant
    Dim slKey As String
    Dim Report As CRAXDRT.Report
    Dim blOneSet As Boolean
    
    Set omReportSource = New ReportSourceRouter
    For Each vlKey In Reports
        blOneSet = True
        Set Report = Reports.Item(vlKey)
        omReportSource.AddReport Report
    Next vlKey
    If blOneSet Then
        Set ReportBin = omReportSource
    Else
        Set ReportBin = Nothing
    End If
End Property
Public Function OpenReport(slReportName As String, Optional blCustomizeLogo As Boolean) As Integer
'openReport calls ModifyLogo and sets it to the default CLogoStandard. bldontsetlogo skips this step.
    'Input blCustomizedLogo:  false to use general logo location and rptlogo.bmp. If true,send only the logo location and the .rpt will concatenate the logo name found with the sales source
    
    Dim olReport As CRAXDRT.Report
    Dim Section As CRAXDRT.Section
    Dim ilReportObject As Integer
    Dim subReportObject As CRAXDRT.subReportObject
    Dim subReport As CRAXDRT.Report
    Dim blRet As Boolean
    
    mSetPathsAndNames
    On Error GoTo ERRORBOX
    Set omReportApplication = New Application
    Set olReport = omReportApplication.OpenReport(smGlobalReportPath & slReportName)
    Debug.Print "Open Report: " & smGlobalReportPath & slReportName
    blRet = mSetLocation(olReport)
    If Not blRet Then
        MsgBox "Couldn't set location of reports " & err.Description, vbCritical, "Report Error" '"Please set a DSN before setting via ODBC", vbCritical, "No DSN"
        OpenReport = 0
        GoTo FINISH
    End If
    'now set subreports
    For Each Section In olReport.Sections
        For ilReportObject = 1 To Section.ReportObjects.Count
            If Section.ReportObjects(ilReportObject).Kind = crSubreportObject Then
                Set subReportObject = Section.ReportObjects(ilReportObject)
                Set subReport = subReportObject.OpenSubreport
                blRet = mSetLocation(subReport)
                If Not blRet Then
                    MsgBox "Please set a DSN before setting via ODBC", vbCritical, "No DSN"
                    OpenReport = 0
                    GoTo FINISH
                End If
            End If      'have subreport
        Next ilReportObject
    Next Section
    olReport.EnableAsyncQuery = True
    blRet = mAddReportToDictionary(olReport, slReportName)
    If Not blRet Then
        OpenReport = 0
        GoTo FINISH
    End If
    mSetJpg blCustomizeLogo            'determine if .jpg exists (vs .bmp)
    ModifyLogo      'sets logo
    OpenReport = -1
FINISH:
    Set olReport = Nothing
    Set Section = Nothing
    Set subReportObject = Nothing
    Set subReport = Nothing
    Exit Function
ERRORBOX:
    MsgBox "Unable to open Print Job for " & slReportName & " " & err.Number & vbCrLf & err.Description, vbCritical & vbOKOnly, "Error"
    OpenReport = 0
    GoTo FINISH
End Function
Public Function OpenAndSetDataSource(slReportName As String, rstActive As ADODB.Recordset) As Integer
    Dim olReport As CRAXDRT.Report
    Dim blRet As Boolean
    mSetPathsAndNames
    On Error GoTo ERRORBOX
    Set omReportApplication = New Application
    Set olReport = omReportApplication.OpenReport(smGlobalReportPath & slReportName)
    Debug.Print smGlobalReportPath & slReportName
    olReport.EnableAsyncQuery = True
    olReport.Database.Tables(1).SetDataSource rstActive, 3
    Debug.Print "set data source"
    blRet = mAddReportToDictionary(olReport, slReportName)
    If Not blRet Then
        OpenAndSetDataSource = 0
        GoTo FINISH
    End If
    ModifyLogo  'set logo
    OpenAndSetDataSource = -1
FINISH:
Set olReport = Nothing
Exit Function
ERRORBOX:
    MsgBox "Unable to open Print Job for " & slReportName & " " & err.Number & vbCrLf & err.Description, vbCritical & vbOKOnly, "Error"
    OpenAndSetDataSource = 0
    GoTo FINISH
End Function
Public Function SetFormula(slName As String, slFormulaValue, Optional slReportName As String) As Integer
    Dim ilLengthWanted As Integer
    Dim olFormulaField As CRAXDRT.FormulaFieldDefinition
    If smReportName <> "" Then
        'strip {@ and } out.  Can't use for "getItemByName"
        If slName Like "{@*" Then
            ilLengthWanted = Len(slName) - 3
            slName = Mid$(slName, 3, ilLengthWanted)
        End If
        On Error GoTo ERRORBOX
        If slReportName = "" Then
            slReportName = smReportName
        End If
        Set olFormulaField = Reports.Item(slReportName).FormulaFields.GetItemByName(slName)
        On Error GoTo 0
        olFormulaField.Text = slFormulaValue
        SetFormula = True
    Else
        MsgBox "Please OpenReport before attempting to set formula", vbExclamation + vbOKOnly
        SetFormula = False
    End If
FINISH:
    Set olFormulaField = Nothing
    Exit Function
ERRORBOX:
   ' MsgBox "Unable to find Formula Name matching " & slName & " ", vbCritical & vbOKOnly, "Error"
    SetFormula = False
    GoTo FINISH
End Function
Public Function SetSelection(slSelection As String, Optional slReportName As String) As Integer
    If smReportName <> "" Then
        If slReportName = "" Then
            slReportName = smReportName
        End If
        On Error GoTo ERRORBOX
        Reports.Item(slReportName).RecordSelectionFormula = slSelection   'ogReport.
        SetSelection = True
    Else
        MsgBox "Please OpenReport before attempting to set selection", vbExclamation + vbOKOnly
        SetSelection = False
    End If
    Exit Function
ERRORBOX:
    MsgBox "Selection error-" & err.Number & vbCrLf & err.Description & vbCrLf & "Selection=" & slSelection, vbCritical & vbOKOnly, "Error"
    SetSelection = False
End Function
Public Sub ModifyLogo(Optional slReportNameIfNotCurrent As String)
'1)sets CLogoStandard as default logo condition (sets logo path to database)
'2) use optional slReportNameIfNotCurrent to set a previously loaded report in collection rather than the one currently opened.
    Dim slReportName As String
    If Not smReportName = "" Then   'no report name set, programmer never opened report
        If slReportNameIfNotCurrent = "" Then   'use current reportName
            slReportName = smReportName
        Else
            slReportName = slReportNameIfNotCurrent
        End If
        If omLogoCondition Is Nothing Then
            Set omLogoCondition = New CLogoStandard
        End If
        'find previous setting and replace
        If Logos.Exists(slReportName) Then
            Set Logos.Item(slReportName) = omLogoCondition
        Else
            Logos.Add slReportName, omLogoCondition
        End If
        Logos.Item(slReportName).Condition slReportName
        Set omLogoCondition = Nothing
    Else
        MsgBox "Please OpenReport before attempting to modify logo", vbExclamation + vbOKOnly
    End If
End Sub
Public Function PrintOut(ilCopies As Integer, Optional blCurrentOnly As Boolean = False) As Integer
    Dim ilCounter As Integer
    Dim vlArray As Variant
'    #If Not Traffic Then    'affiliate only
'    Screen.MousePointer = vbHourglass
'    #End If
    If ilCopies <= 0 Then
        ilCopies = 1
    End If
    If smReportName <> "" Then
        On Error GoTo ERRORBOX
        If blCurrentOnly Then   'not really necessary, as array will only have one item.  Thought it might be a little faster
            Reports.Item(smReportName).PrintOut False, ilCopies
        Else
        'dump dictionary of reports into array
            vlArray = Reports.Items
        'print array
            For ilCounter = 0 To ogReport.Reports.Count - 1
                vlArray(ilCounter).PrintOut False, ilCopies
            Next ilCounter
        End If
        PrintOut = True
    Else
        MsgBox "Please OpenReport before printing", vbExclamation + vbOKOnly
        PrintOut = False
    End If
FINISH:
    Screen.MousePointer = vbDefault
    Exit Function
ERRORBOX:
    MsgBox "Printing to Printer error- " & err.Number & vbCrLf & err.Description, vbCritical & vbOKOnly, "Error"
    PrintOut = False
    GoTo FINISH
End Function
'TTP 10549 - Learfield Cloud printing 911 - added tempFolder, if used, file is exported and opened from temp folder
'Public Function Export(slFileName As String, ilFTSelectedIndex As Integer, Optional blCurrentOnly As Boolean = False) As Integer
'Fix TTP 10826 / TTP 10813 - RE: v81 TTP 10826 - updated test results Issue #4
'Public Function Export(slFileName As String, ilFTSelectedIndex As Integer, Optional blCurrentOnly As Boolean = False, Optional slTempFolder As String = "") As Integer
Public Function Export(slFileName As String, ilFTSelectedIndex As Integer, Optional blCurrentOnly As Boolean = False, Optional slTempFolder As String = "", Optional blHasNTR As Boolean = True) As Integer
'10/12/11 Dan M.  exporting had a possible problem where, after setting 'browse' and file name, user changes the file type. Fixed by deleting extension
'if it exists.
Dim slReptDest As String
Dim ilFileType As Integer
Dim ilRet As Integer
Dim slExt As String
Dim ilExtExists As Integer
Dim ilFormatType As Integer
Dim vlArray As Variant
Dim ilCounter As Integer
Dim slNameWithCounter As String
Dim ogPdf As PDFSplitMerge.CPDFSplitMergeObj
Dim slPdfExport As String
Dim slTotalExportStatement As String
Dim olTempFolder As FileSystemObject
Dim slRestoreDestination As String  'multi pdf hijacks reptdest with temp folder; this stores user's choice.
Dim vlStrings As Variant        'array to hold filename if split function used
Dim myTempReport As CRAXDRT.Report
Dim blCustomExcel As Boolean
Dim blRet As Boolean
Dim slSaveRecordSelectionFormula As String
      ' comes in as:
      '    0= pdf
      '    1= Excel all headers
      '    2=excel column header
      '    3=excel no header
      '    4= Word
      '    5= Text
      '    6= csv
      '    7= rtf
    Screen.MousePointer = vbHourglass
    ilFileType = ilFTSelectedIndex
    If ilFTSelectedIndex < 0 Then   'Affiliate system index values are negative value offset with -1 so that 0 is -1, 1 is -2,..
        '******* FIX AFFILIATE FORMAT TYPES
      ilFileType = -ilFTSelectedIndex - 1
      'Dan I changed this logic. Before, case statement was in an else that started here. then it ended, and ran export. So if above was true, never went into select case.
    End If
    ''Dan M 10/12/11 get extension and remove. moved here before slReptDest
    'ilExtExists = InStr(1, slFileName, ".")
    'If ilExtExists > 0 Then
    '    slFileName = Mid(slFileName, 1, ilExtExists - 1)
    'End If
    
    'JW 3/19/21 get extension and remove. moved here before slReptDest
    'Dont check the WHOLE path because the path might include a "."
    'example:D:\CSI\V81\Export\Counterpoint.net\IT Department - Documents\TEST\
    'TTP 10186 - 5/21/21 - some filenames coming in here are < 4 chars....
    If Len(slFileName) > 4 And InStr(1, slFileName, ".") > 0 Then
        'Length of filename is Greater than 4, and has a ".", it may be an Extension, if so we'll remove it
        'If it's Not an Extension, it could be part of the path (C:\csi.test\Filename.pdf) - dont chop the Path, just the extension
        ilExtExists = InStr(Len(slFileName) - 4, slFileName, ".")
    Else
        'slFileName = Mid(slFileName, 1, ilExtExists - 1)
        'find the Extension
        ilExtExists = InStr(1, slFileName, ".")
    End If
    If ilExtExists > 0 Then
        slFileName = Mid(slFileName, 1, ilExtExists - 1)
    End If
    'did user type path?
    If (InStr(slFileName, ":") = 0) And (Left$(slFileName, 2) <> "\\") Then
        If slTempFolder <> "" Then
            slReptDest = slTempFolder & slFileName
        Else
            slReptDest = smGlobalReportSavePath & slFileName
        End If
    Else
        slReptDest = slFileName
        If imLastPrintJob > 1 Or TreatAsLastReport Then 'will need the file name stripped of the path
            vlStrings = Split(slFileName, "\", , vbBinaryCompare)
            slFileName = CStr(vlStrings(UBound(vlStrings))) 'ok to overwrite slfilename: it's only done for multi reports
        End If
    End If
    'for multiple pdfs that will be merged into one pdf, create a temporary folder to house pdfs
    If (imLastPrintJob > 1 Or TreatAsLastReport) And ilFileType = 0 Then
        Set olTempFolder = New FileSystemObject
        slRestoreDestination = slReptDest
        If slTempFolder <> "" Then
            slReptDest = slTempFolder & slFileName
        Else
            slReptDest = smGlobalReportSavePath & "pdfTempMerge\" & slFileName    'use this instead of normal report name for multi pdf
        End If
On Error Resume Next

        If slTempFolder <> "" Then
            olTempFolder.CreateFolder slTempFolder & "pdfTempMerge\"
        Else
            olTempFolder.CreateFolder smGlobalReportSavePath & "pdfTempMerge\"
        End If
    End If
    slExt = ""
'    ilExtExists = True
'    If InStr(slFileName, ".") = 0 Then  'no extension specified
'        ilExtExists = False
'    End If
    Select Case ilFileType
        Case 0      'pdf
            slExt = ".pdf"
            ilFormatType = crEFTPortableDocFormat
        Case 1, 2, 3
            slExt = ".xls"
            ilFormatType = crEFTExcelDataOnly
        Case 4      'rtf-word
            slExt = ".doc"
            ilFormatType = crEFTWordForWindows
        Case 5      'txt
            slExt = ".txt"
            ilFormatType = crEFTText
        Case 6       'csv
            slExt = ".csv"
            ilFormatType = crEFTCommaSeparatedValues
        Case 7
            slExt = ".rtf"
            ilFormatType = crEFTExactRichText
    End Select
'    If ilExtExists Then
'        slExt = ""
'    End If

''reference:  pdfsplitmerge 1.0 type library
'Set Pdf = New PDFSplitMerge.CPDFSplitMergeObj
'Pdf.SetCode "989888DFE6F0FF0056681A"
''various options:
''pdf.Split "..\1.pdf", "1;1;1", "..\sp1-%d.pdf"
''pdf.Split "..\1.pdf", "1;2;3;4;1-2,3-4;1", "..\sp%d.pdf"
''
''pdf.Merge "..\sp0.pdf|..\sp1.pdf", "..\m1.pdf"
''pdf.Merge "..\1.pdf?1-2|..\sp1.pdf", "..\m2.pdf"
''pdf.Merge "..\1.pdf?1-2|..\sp1.pdf?2", "..\m3.pdf"
'my examples:
''pdf.Merge "C:\pdftest\1.pdf|C:\pdftest\2.pdf", "C:\pdftest\PdfForDarlene.pdf"
'using variables
'pdf.Merge "C:\pdftest\" & pdf1 & "|C:\pdftest\" & pdf2, "C:\pdftest\" & pdfmerged
'Set pdf = Nothing

On Error GoTo ERRORBOX
        'export array
    If blCurrentOnly Then   'single report, skip array
        If ilFormatType = crEFTExcelDataOnly Then
            blCustomExcel = mSetExcelOptions(Reports.Item(smReportName), ilFileType)
        ElseIf ilFormatType = crEFTPortableDocFormat Then
          ' Set myTempReport = Reports.Item(smReportName)
            mSetPdfOptions Reports.Item(smReportName)
        End If
        With Reports.Item(smReportName)
            .ExportOptions.DestinationType = crEDTDiskFile
            .ExportOptions.FormatType = ilFormatType
            .ExportOptions.DiskFileName = slReptDest & slExt
'On Error GoTo ERROREXPORT
            Reports.Item(smReportName).Export False
            mFixExcel blCustomExcel
'On Error GoTo errorbox
       End With
    Else
        vlArray = Reports.Items
        For ilCounter = 0 To ogReport.Reports.Count - 1
            Debug.Print vlArray(ilCounter).ReportTitle
            'Fix TTP 10826 / TTP 10813 - RE: v81 TTP 10826 - updated test results Issue #4
            If (vlArray(ilCounter).ReportTitle = "Inv_NTR" And blHasNTR = True) Or vlArray(ilCounter).ReportTitle <> "Inv_NTR" Then
                If ilFormatType = crEFTExcelDataOnly Then
                    Set myTempReport = vlArray(ilCounter)
                    blRet = mSetExcelOptions(myTempReport, ilFileType)
                    If blRet Then
                        blCustomExcel = True
                    End If
                ElseIf ilFormatType = crEFTPortableDocFormat Then
                    Set myTempReport = vlArray(ilCounter)
                    mSetPdfOptions myTempReport
                Else
                    Set myTempReport = vlArray(ilCounter)       'exports other than pdf & excel have no options
                End If
                vlArray(ilCounter).ExportOptions.FormatType = ilFormatType
                vlArray(ilCounter).ExportOptions.DestinationType = crEDTDiskFile
                 ' one report, don't worry about adding number.
                'Debug.Print slReptDest & slExt
                vlArray(ilCounter).ExportOptions.DiskFileName = slReptDest & slExt
                 'multiple print jobs? change name to have "-#" appended.  Remember that slReptDest has been changed to temp folder for pdfs.
                If imLastPrintJob > 1 Or TreatAsLastReport Then
                    slNameWithCounter = slReptDest & "-" & ilCounter + 1
    '                If ilExtExists Then         'test  slext to see if need to split name before adding. Then add ilcounter to name so each report has different name
    '                    slNameWithCounter = Replace(slReptDest, ".", "-" & ilCounter + 1 & ".", 1, 1, vbBinaryCompare)
    '                End If
                    If ilFormatType = crEFTPortableDocFormat Then   'preparing merge string
                        slPdfExport = slPdfExport & slNameWithCounter & slExt & "|"
                    End If
                    vlArray(ilCounter).ExportOptions.DiskFileName = slNameWithCounter & slExt
                    Debug.Print vlArray(ilCounter).ExportOptions.DiskFileName
                End If
                slSaveRecordSelectionFormula = myTempReport.RecordSelectionFormula
                If Len(AddToSelection) > 0 Then
                    myTempReport.DiscardSavedData       'Discard needs to be done with temporary report
                    myTempReport.RecordSelectionFormula = myTempReport.RecordSelectionFormula & AddToSelection
                End If
                vlArray(ilCounter).Export False
                myTempReport.RecordSelectionFormula = Trim$(slSaveRecordSelectionFormula)       'restore of selection must be done with temporary report
                mFixExcel blCustomExcel
            End If
         Next ilCounter
         Screen.MousePointer = vbHourglass
         If ilFormatType = crEFTPortableDocFormat And (imLastPrintJob > 1 Or TreatAsLastReport) Then     'made our pdfs.  Now merge them if mulitple
             'get rid of last "|"
             slTotalExportStatement = Mid$(slPdfExport, 1, Len(slPdfExport) - 1)
             
             'JW 09/09/2022 - I think pdf.merge crashes if Crystal is still finishing the last pdf file; added 1 sec delay
             Sleep (1000)
             Set ogPdf = New PDFSplitMerge.CPDFSplitMergeObj
             ogPdf.SetCode "989888DFE6F0FF0056681A"
             ogPdf.Merge slTotalExportStatement, slRestoreDestination & slExt
        End If
    End If  'blcurrentonly
    
    'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe to view document
    If bgUseAdobe = True And slTempFolder <> "" And sgReportFilename <> "" Then
        'Wait for PDF / pdf.Merge before opening PDF file
        Sleep (1000)
        'Open PDF file and Wait, then delete it...
        'slFileName
        If gFileExist(sgReportTempFolder & sgReportFilename & slExt) = FILEEXISTS Then
            ShellAdobeAndDelete sgReportTempFolder & sgReportFilename & slExt
        Else
            ShellAdobeAndDelete sgReportTempFolder & slFileName & slExt
        End If
    End If
    'send message in affiliate that reports have been exported
    Export = True
FINISH:
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'now get rid of all those pdfs that we merged.
    If Not olTempFolder Is Nothing Then
        If olTempFolder.FolderExists(smGlobalReportSavePath & "pdfTempMerge") Then
             olTempFolder.DeleteFolder smGlobalReportSavePath & "pdftempmerge", True
        End If
    End If
    Set ogPdf = Nothing
    Set olTempFolder = Nothing
    Set myTempReport = Nothing
    Exit Function
'ERROREXPORT:
'    If Reports.Item(smReportName).ExportOptions.ExcelChopPageHeader = False Then
'        Reports.Item(smReportName).ExportOptions.ExcelChopPageHeader = True
'        Resume
'    End If
ERRORBOX:
    MsgBox "Failed to export report" & err.Number & vbCrLf & err.Description, vbCritical & vbOKOnly, "Error"
    Export = False
    GoTo FINISH
End Function
Public Sub ReadRecords()
Reports.Item(smReportName).ReadRecords
End Sub
Public Sub RemoveReport()
On Error Resume Next
    Reports.Remove Me.CurrentReportName
End Sub
Private Sub mFixExcel(blCustomExcel As Boolean)
    If blCustomExcel Then
        SetFormula "ShowHeader", "'Y'"
        blCustomExcel = False
    End If
End Sub
Private Sub mSetPdfOptions(myReport As CRAXDRT.Report)
If PdfPageFirst > 0 Or PdfPageLast > 0 Then
    With myReport.ExportOptions
        .PDFExportAllPages = False
        If PdfPageFirst > 0 Then
            .PDFFirstPageNumber = PdfPageFirst
        End If
        If PdfPageLast > 0 Then
            .PDFLastPageNumber = PdfPageLast
        End If
    End With
Else
    myReport.ExportOptions.PDFExportAllPages = True
    PdfPageFirst = 0
    PdfPageLast = 0
End If
End Sub
Private Function mSetExcelOptions(myReport As CRAXDRT.Report, ilOption As Integer) As Boolean
    Dim blRet As Boolean
    
    With myReport.ExportOptions
        .ExcelMaintainColumnAlignment = True
        .ExcelMaintainRelativeObjectPosition = True
        .ExcelUseFormatInDataOnly = True
        .ExcelExportPageAreaPair = crExportPageAreaPairOncePerReport
        Select Case ilOption
            'all headers
            Case 1
                .ExcelChopPageHeader = False
                blRet = Me.SetFormula("ShowHeader", "'Y'")
           'some
            Case 2
                .ExcelChopPageHeader = True
                blRet = Me.SetFormula("ShowHeader", "'X'")
            'no
            Case 3
                .ExcelExportPageAreaPair = crExportNoPageAreaPair
                blRet = Me.SetFormula("ShowHeader", "'N'")
        End Select
    End With
    mSetExcelOptions = blRet
End Function
Private Function mSetLocation(olReport As CRAXDRT.Report) As Boolean
    Dim blDoOdbc As Boolean
    Dim slGlobalDsnName As String
    On Error GoTo ERRBOX
    With olReport.Database.Tables(1)
        If bmForceToOdbc Or InStr(1, .DllName, "odbc.dll", vbTextCompare) > 0 Then
            slGlobalDsnName = myConnection.GlobalDSNName
            If Len(slGlobalDsnName) > 0 Then
                  blDoOdbc = True
            End If
        End If
        If blDoOdbc Then
            'Debug.Print "crdb_odbc.dll"
            .DllName = "crdb_odbc.dll"
            .ConnectionProperties.DeleteAll
            .ConnectionProperties.Add "DSN", slGlobalDsnName
            .ConnectionProperties.Add "database", slGlobalDsnName
        Else
            'Debug.Print "p2bbtrv.dll"
            .DllName = "crdb_p2bbtrv.dll"
            .ConnectionProperties.DeleteAll
            .ConnectionProperties.Add "Data File", smGlobalDatabasePath & "file.ddf"
            .ConnectionProperties.Add "Data File Search Path", smGlobalDatabasePath
        End If
    End With
    mSetLocation = True
    Exit Function
ERRBOX:
    mSetLocation = False
    
End Function

Private Function mAddReportToDictionary(olReport As CRAXDRT.Report, slReportName As String) As Boolean
    If imLastPrintJob = 1 And Reports.Count > 0 Then 'only one job, and something already in dictionary? get rid of it.
        Reports.removeAll
    End If
    On Error GoTo alreadyOpened
    If Reports.Exists(slReportName) Then    'same report called twice, so give new name to recent one.
        slReportName = slReportName & "-" & (imNumberForSameName + 1)
    End If
    Reports.Add slReportName, olReport
    smReportName = slReportName
    mAddReportToDictionary = True
    Exit Function
alreadyOpened:
    MsgBox "That report is already opened!", vbCritical & vbOKOnly, "Error"
    mAddReportToDictionary = False
End Function
Private Sub mSetPathsAndNames()
    
    With myConnection
        smGlobalReportPath = .GlobalReportPath
        smGlobalReportSavePath = .GlobalReportSavePath
        smGlobalDatabasePath = .GlobalDatabasePath
        LogoPath = .GlobalLogoPath
        smSystem = .System
    End With
End Sub
Private Sub mSetJpg(blCustomizeLogo As Boolean)
    '<input>  blCustomizeLogo 12-1-14 true if customized logo defined in the .rpt to use sales source logo name
    Dim slReportImage As String
    
    'slReportImage = myConnection.RootDrive & "csi\rptLogo.jpg"
    
    
    '12-1-14 the .rpt will be concatenating the logo location with the logo name defined with the sales source (i.e. invoices, forms reports)
    If blCustomizeLogo Then
        slReportImage = myConnection.GlobalLogoPath
        '8886
        '6-18-18 chg back to use Dir as gFileExists does not have a filename to be tested.  The CSI Star was defaulted to printing
       '8/19/18: Changed to use the gFolderExist to replace Dir. The gFileExist only checks if files exist and should not have been used
        'If Dir(slReportImage) > "" Then         'customized exist in folder?
        If gFolderExist(slReportImage) Then
        '***** for c# - test directory.exists(path) for conversion
        'If gFileExist(slReportImage) = FILEEXISTS Then
            slReportImage = "'" & slReportImage & "'"
            SetFormula "logolocation", slReportImage
        Else            'get default
'            slReportImage = myConnection.GlobalLogoPath & "rptlogo.jpg"
'            If Dir(slReportImage) > "" Then
'                slReportImage = "'" & slReportImage & "'"
'                SetFormula "logolocation", slReportImage
'            Else
'                slReportImage = myConnection.GlobalLogoPath & "rptlogo.bmp"
'                slReportImage = "'" & slReportImage & "'"
'                SetFormula "logolocation", slReportImage
'            End If
        End If
    Else
        '8-14-14
        slReportImage = myConnection.GlobalLogoPath & "rptlogo.jpg"
        '8886
       ' If Dir(slReportImage) > "" Then
        If gFileExist(slReportImage) = FILEEXISTS Then
            slReportImage = "'" & slReportImage & "'"
            SetFormula "logolocation", slReportImage
        Else
            slReportImage = myConnection.GlobalLogoPath & "rptlogo.bmp"
            slReportImage = "'" & slReportImage & "'"
            SetFormula "logolocation", slReportImage
        End If
    End If

End Sub

Public Function SetSelectionPlus(slSelection As String, Optional slReportName As String) As Integer
    If smReportName <> "" Then
        If slReportName = "" Then
            slReportName = smReportName
        End If
        On Error GoTo ERRORBOX
        Reports.Item(slReportName).RecordSelectionFormula = Reports.Item(slReportName).RecordSelectionFormula & slSelection   'ogReport.
        SetSelectionPlus = True
    Else
        MsgBox "Please OpenReport before attempting to set selection", vbExclamation + vbOKOnly
        SetSelectionPlus = False
    End If
    Exit Function
ERRORBOX:
    MsgBox "Selection error-" & err.Number & vbCrLf & err.Description, vbCritical & vbOKOnly, "Error"
    SetSelectionPlus = False
End Function
