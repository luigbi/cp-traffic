Attribute VB_Name = "CRPEVAR"
' Dan M 8-26-08 changed file to  use rdc rather than api calls.
Option Explicit
Public ogReport As CReportHelper
'Public sgDatabaseName As String 'dsn name
'Public sgLogoName As String
Public igDestination As Integer 'not used for rdc, but used in rest of traffic
' 11/17/11 to make work
Public bgComingFromInvoice As Boolean
'Public sgLogoName As String
Public bgReportModuleRunning As Boolean
Public igPreviousTimes() As Integer
Public igPreviousDates() As Integer
'Dan rollback copybook 5/12/09 removed 9/03/09
'Public bgRollback As Boolean
Public igPrtJob As Integer  'Print job number
Public igErrorCode As Integer   'Error code
Public sgErrorMsg As String 'Error message
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
'8/26/11 currently unused:
Public Enum CsiReportCall
    StartReports
    Normal
    FinishReports
End Enum

Function gOpenPrtJob(slReportName As String, Optional olConditionChoice As IConditionalLogo, Optional blCustomizeLogo As Boolean = False) As Integer
'12-1-14  ANY REPORT THAT REQUIRES CUSTOMIZED LOGOS MUST PASS blCustomizeLogo FLAG AS TRUE WHEN OPENING THE CRYSTAL .RPT
'Purpose:  Set database table for report and any subreports, and sets oleobject in page header.
'ogreport holds a collection of report and logo objects. These are available for the programmer to modify after ogreport.openreport, and before the call to peClosePrintJob.  See class for examples
'Input: Report Name.  Optional: olConditionChoice if don't want to set logo to default path and name.
'To use optional, might want to call like this:
'            If ilListIndex = 0 Then
'               Dim olConditionChoice As LogoExample
'               Set olConditionChoice = New LogoExample
'               If Not gOpenPrtJob("Adf.Rpt", olconditionChoice) Then
'                    gGenReport = False
'                    Exit Function
'                End If
    Dim ilRet As Integer
    'Debug.Print "Opening Report: " & slReportName

    If ogReport Is Nothing Then     'only 'something' if multiple report job.
        Set ogReport = New CReportHelper
    End If
    If Not olConditionChoice Is Nothing Then        'passed a condition choice
        Set ogReport.Alternatelogo = olConditionChoice
    End If
    'ogReport.Connect = Native
    ilRet = ogReport.OpenReport(slReportName, blCustomizeLogo)
    If Not ilRet Then
        gOpenPrtJob = False
        Set ogReport = Nothing
        Exit Function
    End If
    gOpenPrtJob = True
    
    #If Reports = 1 Then
    '5/24/19: Save Control settings
    gSaveReportCtrlsSetting
    #End If
    
    gUserActivityLog "S", sgReportListName & ": Prepass"
    Set olConditionChoice = Nothing
End Function

Function gSetFormula(slName As String, slFormulaValue, Optional blClosePrintjobOnFail As Boolean = True) As Integer
    Dim ilRet As Integer
    Debug.Print " -> gSetFormula; " & slName & "=" & slFormulaValue
    
    ilRet = ogReport.SetFormula(slName, slFormulaValue)
    If ilRet Then
        gSetFormula = True
    Else
        Debug.Print " ---> gSetFormula Failed:" & slName
        gSetFormula = False
        If blClosePrintjobOnFail Then
            Debug.Print " ---> closing PrintJob!!!"
            PEClosePrintJob
        End If
    End If
End Function

Function gSetSelection(slSelection As String) As Integer
    Debug.Print " -> gSetSelection; " & slSelection
    Dim ilRet As Integer
    ilRet = ogReport.SetSelection(slSelection)
    If ilRet Then
        gSetSelection = True
    Else
        Debug.Print " ---> gSetSelection Failed:" & slSelection
        gSetSelection = False
        PEClosePrintJob
    End If
End Function

'TTP 10549 - Learfield Cloud printing 911 - added tempFolder, if used, file is exported and opened from temp folder
'Public Function gExportCRW(slFileName As String, ilFTSelectedIndex As Integer, Optional blKeepReportOpen As Boolean = False) As Integer
'Fix TTP 10826 / TTP 10813 - RE: v81 TTP 10826 - updated test results Issue #4
'Public Function gExportCRW(slFileName As String, ilFTSelectedIndex As Integer, Optional blKeepReportOpen As Boolean = False, Optional slTempFolder As String = "") As Integer
Public Function gExportCRW(slFileName As String, ilFTSelectedIndex As Integer, Optional blKeepReportOpen As Boolean = False, Optional slTempFolder As String = "", Optional blHasNTR As Boolean = True) As Integer
    Dim ilRet As Integer
    'dan 8/26/11 run by Dick
    gUserActivityLog "E", sgReportListName & ": Prepass"
    gUserActivityLog "S", sgReportListName & ": Exporting"
    'TTP 10549 - Learfield Cloud printing 911 - added tempFolder, if used, file is exported and opened from temp folder
    'ilRet = ogReport.Export(slFileName, ilFTSelectedIndex, False)
    'Fix TTP 10826 / TTP 10813 - RE: v81 TTP 10826 - updated test results Issue #4
    'ilRet = ogReport.Export(slFileName, ilFTSelectedIndex, False, slTempFolder)
    ilRet = ogReport.Export(slFileName, ilFTSelectedIndex, False, slTempFolder, blHasNTR)
    If ilRet Then
        gExportCRW = True
        If blKeepReportOpen = False Then    'added this call from display.  If close report, get an error.
            PEClosePrintJob
        End If
    Else
        gExportCRW = False
    End If
    gUserActivityLog "E", sgReportListName & ": Exporting"
End Function

Public Function gPopExportTypes(cbcFileType As Control, Optional blShowFirst As Boolean)
    With cbcFileType
        .AddItem "Adobe Acrobat(PDF)"
        .AddItem "Excel(XLS)-All headers"
        .AddItem "Excel(XLS)-Column headers"
        .AddItem "Excel(XLS)-No headers"
        .AddItem "Word(DOC)"
        .AddItem "Text(TXT)"
        .AddItem "Comma Separated Values(CSV)"
        .AddItem "Rich Text File(RTF)"
        .ListIndex = 0
    End With
End Function

Function PEOpenEngine() As Integer
    'No longer needed for RDC
    PEOpenEngine = True
End Function

Sub PECloseEngine()
    'No longer needed for RDC
    Set ogReport = Nothing              '2-19-17 all reports already call this routine, so clear the print object here
End Sub

Sub PEClosePrintJob()
    'only called by crpevar functions and report unload
    Set ogReport = Nothing
End Sub

Function gOutputToPrinter(ilCopies As Integer, Optional blKeepReportOpen As Boolean = False) As Integer
    Dim ilRet As Integer

    'dan 8/26/11 run by Dick
    gUserActivityLog "E", sgReportListName & ": Prepass"
    gUserActivityLog "S", sgReportListName & ": Printing"
    ilRet = ogReport.PrintOut(ilCopies)
    If ilRet Then
        gOutputToPrinter = True
        If blKeepReportOpen = False Then    'added this call from display.  If close report, get an error.
            PEClosePrintJob
        End If
    Else
        gOutputToPrinter = False
    End If
    
    'PEClosePrintJob
    gUserActivityLog "E", sgReportListName & ": Printing"
End Function

''Dan M 8/18/2010
Public Sub gAdjustCDCFilter(ilChoice As Integer, myDialog As CommonDialog)
    With myDialog
        Select Case ilChoice
            Case 1, 2, 3 'xls
                .Filter = "Excel(*.xls)|*.xls"
            Case 4  '2 'doc
                .Filter = "Word(*.doc)|*.doc"
            Case 5  '3 'txt
                .Filter = "Text(*.txt)|*.txt"
            Case 6  '4 'csv
                .Filter = "Comma Separated Values(*.csv)|*.csv"
            Case 7  '5 'rtf
                .Filter = "Rich Text(*.rtf)|*.rtf"
            Case Else '0=pdf
                .Filter = "PDF(*.pdf)|*.pdf"
        End Select
    End With
End Sub

'Dan testing only.  In reportList(?), pass each report form to this to be able to set common dialog options
Public Sub gAssignCDCFilter(ilForm As Form)
    'common dialog controls don't have generic settings for file types.  I am trying to set that here.
    Dim olControl As Control
    Dim CDC As CommonDialog
    On Error Resume Next
    Set CDC = ilForm.Controls("cdcsetup")
    If err.Number = 0 Then
        CDC.Filter = "Dan set(*.dan)|Text(*.txt)"
    Else
        For Each olControl In ilForm.Controls
            If TypeOf olControl Is CommonDialog Then
                MsgBox olControl.Name
                Exit Sub
            End If
        Next olControl
    End If
End Sub

