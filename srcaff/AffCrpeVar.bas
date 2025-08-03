Attribute VB_Name = "modCrpeVar"
Option Explicit
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CRPEVAR.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Global variables
Public igPrtJob As Integer  'Print job number
Public igErrorCode As Integer   'Error code
Public sgErrorMsg As String 'Error message
Public igDestination As Integer
Public Const WS_VISIBLE = &H10000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000


Private tmPEExportOptions As PEExportOptions

Function gGetFormulaString(ilPrtJob As Integer, slFormulaName As String) As String
    Dim llTextHandle As Long
    Dim ilTextLength As Integer
    Dim ilRet As Integer
    Dim slFormula As String
    ilRet = PEGetFormula(ilPrtJob, slFormulaName, llTextHandle, ilTextLength)
    slFormula = String$(ilTextLength + 1, " ")
    ilRet = PEGetHandleString(llTextHandle, slFormula, ilTextLength)
    gGetFormulaString = slFormula
End Function
Sub gGetPrtErrorString(ilPrtJob As Integer, slMsg As String)
    Dim llTextHandle As Long
    Dim ilTextLength As Integer
    Dim ilRet As Integer
    igErrorCode = PEGetErrorCode(ilPrtJob)
    If igErrorCode > 0 Then
        ilRet = PEGetErrorText(ilPrtJob, llTextHandle, ilTextLength)
        sgErrorMsg = String$(ilTextLength + 1, " ")
        ilRet = PEGetHandleString(llTextHandle, sgErrorMsg, ilTextLength)
        MsgBox slMsg & sgErrorMsg & Str$(igErrorCode)
    End If
End Sub
Function gGetSelectionString(ilPrtJob As Integer) As String
    Dim llTextHandle As Long
    Dim ilTextLength As Integer
    Dim ilRet As Integer
    Dim slFormula As String
    ilRet = PEGetSelectionFormula(ilPrtJob, llTextHandle, ilTextLength)
    slFormula = String$(ilTextLength + 1, " ")
    ilRet = PEGetHandleString(llTextHandle, slFormula, ilTextLength)
    gGetSelectionString = slFormula
End Function
Function gOpenPrtJob(slRptName As String) As Integer
    Dim ilNumTables As Integer
    Dim ilLoop As Integer
    Dim tlLocTable As PETableLocation
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim slChar As String
    Dim slStr As String
    Dim ilTempFd As Integer
    Dim ilTemp As Integer
    'Test if name exist
    ilRet = 0
    On Error GoTo gOpenPrtJobErr
    'slStr = FileDateTime(sgRptPath & slRptName)
    'slStr = FileDateTime("d:\csi\prod_v50\afrepor\" & "aflabels.rpt")
    On Error GoTo 0
    If ilRet = 1 Then
        'MsgBox "Unable to find Print File " & slRptName, MB_OK, "Error"
        gOpenPrtJob = False
        Exit Function
    End If
    'igPrtJob = PEOpenPrintJob(sgRptPath & slRptName)
    igPrtJob = PEOpenPrintJob("d:\csi\prod_v50\affrepor\aflabels.rpt")
    If igPrtJob = 0 Then
        'MsgBox "Unable to Open Print Job for " & slRptName, MB_OK, "Error"
        gOpenPrtJob = False
        Exit Function
    End If
    'Set path for Database
    ilNumTables = PEGetNTables(igPrtJob)
    For ilLoop = 0 To ilNumTables - 1 Step 1
        tlLocTable.StructSize = Len(tlLocTable)
        ilRet = PEGetNthTableLocation(igPrtJob, ilLoop, tlLocTable)
        ilPos = InStr(tlLocTable.Location, Chr$(0))
        If ilPos > 0 Then
            ilStart = ilPos
            Do
                ilStart = ilStart - 1
                If ilStart <= 0 Then
                    Exit Do
                End If
                slChar = Mid$(tlLocTable.Location, ilStart, 1)
            Loop While slChar <> "\"
            If slChar = "\" Then
                slChar = Mid$(tlLocTable.Location, ilStart + 1, ilPos - ilStart)
                ilTempFd = False
                'For ilTemp = LBound(sgTDBNames) To UBound(sgTDBNames) Step 1
                '    If StrComp(slChar, sgTDBNames(ilTemp), 1) = 0 Then
                '        ilTempFd = True
                '        Exit For
                '    End If
                'Next ilTemp
                If ilTempFd Then
                    'tlLocTable.Location = sgTDBPath & slChar
                Else
                    'tlLocTable.Location = sgDBPath & slChar
                End If
                ilRet = PESetNthTableLocation(igPrtJob, ilLoop, tlLocTable)
            End If
        End If
    Next ilLoop
    gOpenPrtJob = True
    Exit Function
gOpenPrtJobErr:
    ilRet = 1
    Resume Next
End Function
Function gOutputToFile(slFileName As String, ilFTIndex As Integer)
    Dim ilFileType As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slName As String
    Dim FormatDLLName As String
    Dim formatType As Long
    Dim FormatOptions As Long
    Dim DestinationDLLName As String
    Dim UseSameDateFormat As Integer
    Dim UseSameNumberFormat As Integer
    Dim StringDelimiter As String
    Dim FieldDelimiter As String
            
    On Error Resume Next
    
    
    If InStr(slFileName, ":") = 0 Then
        'slName = sgRptSavePath & slFileName
    Else
        slName = slFileName
    End If
    Kill slName   'Remove file as crystal must be allow the same name to be used
    On Error GoTo 0
    If ilFTIndex < 0 Then   'Affiliate system index values are negative value offset with -1 so that 0 is -1, 1 is -2,..
        ilFileType = -ilFTIndex - 1
    Else
        Select Case ilFTIndex
            Case 0  'Report equals Text style
                ilFileType = 2
            Case 1  'Fixed Column Width equals report style
                ilFileType = 0
            Case 2  'Comma-Separated with quotes equals CSV
                ilFileType = 4
            Case 3  'Tab-Separated with quotes equals Tab separated
                ilFileType = 1
            Case 4  'Tab-Separated w/o quotes
                ilFileType = 6
            Case 5  'DIF
                ilFileType = 3
            Case Else   'rtf        '11-11-99
                ilFileType = 5
        End Select
    End If
    If ilFileType = 5 Then          'rtf format
        FormatDLLName = "uxfrtf.dll"
        formatType = 0 'crUXFRichTextFormatType
        UseSameNumberFormat = True
        UseSameDateFormat = True
        StringDelimiter = ""
        FieldDelimiter = ""
        'ilRet = crPEExportToDisk(igPrtJob, slName & vbNullChar, FormatDLLName & vbNullChar, FormatType, UseSameNumberFormat, UseSameDateFormat, StringDelimiter & vbNullChar, FieldDelimiter & vbNullChar)
        'ilRet = crPEExportToDisk(igPrtJob, slName, FormatDLLName, FormatType, UseSameNumberFormat, UseSameDateFormat, StringDelimiter, FieldDelimiter)
    Else
        ilRet = PEOutputToFile(igPrtJob, slName, ilFileType, 0)
    End If
    ilRet = PEStartPrintJob(igPrtJob, True)
    If ilRet = 0 Then
        gGetPrtErrorString igPrtJob, "Printing to file error- "
        PEClosePrintJob igPrtJob
        gOutputToFile = False
    Else
        PEClosePrintJob igPrtJob
        gOutputToFile = True
    End If
    Exit Function
End Function
Function gOutputToPrinter(ilCopies As Integer) As Integer
    Dim ilRet As Integer
    Dim slStr As String
    'Dim X As Printer
    If ilCopies <= 0 Then
        ilCopies = 1
    End If
    'For Each X In Printers
    '    If StrComp(X.DeviceName, "Acrobat PDFWriter", vbTextCompare) = 0 Then
    '        ilRet = PESelectPrinter(igPrtJob, X.driverName, X.DeviceName, X.Port, 0)
    '        Exit For
    '    End If
    'Next
    ilRet = PEOutputToPrinter(igPrtJob, ilCopies)
    ilRet = PEStartPrintJob(igPrtJob, True)
    If ilRet = 0 Then
        gGetPrtErrorString igPrtJob, "Printing to printer error- "
        PEClosePrintJob igPrtJob
        gOutputToPrinter = False
    Else
        PEClosePrintJob igPrtJob
        gOutputToPrinter = True
    End If
    Exit Function
End Function
Function gSetFormula(slFormulaName, slFormulaValue)
    Dim slName As String
    Dim slValue As String
    Dim ilRet As Integer
    Dim slStr As String
    
    slName = slFormulaName
    slValue = slFormulaValue
    If (slName <> "") And (slValue <> "") Then
        slName = slName & Chr$(0)
        slValue = slValue & Chr$(0)
        ilRet = PESetFormula(igPrtJob, slName, slValue)
        If ilRet = 0 Then
            'gGetPrtErrorString igPrtJob, "Formula error- "
            gGetPrtErrorString igPrtJob, "Formula (" & slFormulaName & ": " & slFormulaValue & ") error-"
            PEClosePrintJob igPrtJob
            gSetFormula = False
        Else
            gSetFormula = True
        End If
    Else
        gSetFormula = True
    End If
End Function
Function gSetSelection(slSelection As String)
    Dim ilRet As Integer
    Dim slStr As String
    Dim slSel As String
    slSel = slSelection
    If slSel <> "" Then
        slSel = "(" & slSel & ")"
        slSel = slSel & Chr$(0)
        ilRet = PESetSelectionFormula(igPrtJob, slSel)
        If ilRet = 0 Then
            gGetPrtErrorString igPrtJob, "Selection error- "
            PEClosePrintJob igPrtJob
            gSetSelection = False
        Else
            gSetSelection = True
        End If
    Else
        gSetSelection = True
    End If
End Function
Public Function gExportCRW(slFileName As String, imFTSelectedIndex As Integer) As Integer

Dim slDLLName As String
Dim slReptDest As String
Dim llCRReportType As Long
Dim ilFileType As Integer
Dim llCR_parm4 As Integer
Dim llCR_parm5 As Long
Dim llCR_parm6 As Long
Dim slCR_parm7 As String
Dim slCR_parm8 As String
Dim ilRet As Integer
Dim slExt As String
Dim ilExtExists As Integer
    
    ' 0 = "Adobe PDF"
    ' 1 = "Comma separated value"
    ' 2 = "Data Interchange"
    ' 3 = "Excel7"
    ' 4 = "Excel8"
    ' 5 = "RTF"
    ' 6 = "Tab separated text"
    ' 7 = "Text"
    ' 8 = "Word for Windows"    'initialize to default values
    
    llCR_parm4 = 0
    llCR_parm5 = 0
    llCR_parm6 = 0
    slCR_parm7 = ""
    slCR_parm8 = ""
 
    ilFileType = imFTSelectedIndex
    If InStr(slFileName, ":") = 0 Then
        'slReptDest = sgRptSavePath & slFileName
        slReptDest = "C:\" & "TestPdf.pdf"
    Else
        slReptDest = slFileName
    End If
    
    slExt = ""
    ilExtExists = True
    If InStr(slFileName, ".") = 0 Then  'no extension specified
        ilExtExists = False
    End If
    
    If imFTSelectedIndex < 0 Then   'Affiliate system index values are negative value offset with -1 so that 0 is -1, 1 is -2,..
        '******* FIX AFFILIATE FORMAT TYPES
        
        ilFileType = -imFTSelectedIndex - 1
    Else
        Select Case ilFileType
            Case 0                  'Adobe
                slDLLName = "crxf_pdf.dll"
                llCRReportType = 0
                slExt = ".pdf"
                
            Case 1                  'Comma Sep Val
                slDLLName = "u2fsepv.dll"
                llCRReportType = crUXFCommaSeparatedType
                slExt = ".csv"
                 
            Case 2                  'DIF
                slDLLName = "u2fdif.dll"
                llCRReportType = crUXFDIFType
                slExt = ".dif"
                 
            Case 3                  'XLS 7
                slDLLName = "u2fxls.dll"
                llCRReportType = crUXFXls5Type
                slExt = ".xls"
                llCR_parm4 = 7
            Case 4                  'XLS 8 tabulated
                slDLLName = "u2fxls.dll"
                llCRReportType = crUXFXls5Type
                slExt = ".xls"
                llCR_parm4 = 8
            Case 5
                slDLLName = "u2frtf.dll"
                llCRReportType = crUXFRichTextFormatType
                slExt = ".rtf"
            Case 6                  'Tab separated
                slDLLName = "u2ftext.dll"
                llCRReportType = crUXFTabbedTextType
                slExt = ".txt"
              
            Case 7                  'text
                slDLLName = "u2ftext.dll"
                llCRReportType = crUXFPaginatedTextType
                slExt = ".txt"
                llCR_parm5 = 60
            Case 8                  'Word
                slDLLName = "u2fwordw.dll"
                llCRReportType = crUXFWordWinType
                slExt = ".doc"
                
        End Select
        If ilExtExists Then
            slExt = ""
        End If
    End If
    ilRet = crPEExportToDisk(igPrtJob, slReptDest & slExt & vbNullChar, slDLLName & vbNullChar, llCRReportType, llCR_parm5, llCR_parm6, slCR_parm7 & vbNullChar, slCR_parm8 & vbNullChar)
        
    
    If ilRet = 0 Then
        gGetPrtErrorString igPrtJob, "Printing to file error- "
        PEClosePrintJob igPrtJob
        gExportCRW = False
    Else
        ilRet = PEStartPrintJob(igPrtJob, 1)
        PEClosePrintJob igPrtJob
        gExportCRW = True
    End If

    Exit Function
End Function
 
'
'       Build the list of all valid export types form Crystal
'       10-18-01

Public Function gPopExportTypes(cbcFileType As Control)
    
    cbcFileType.AddItem "Acrobat PDF"
    cbcFileType.AddItem "Comma separated value"
    cbcFileType.AddItem "Data Interchange"
    cbcFileType.AddItem "Excel 7"
    cbcFileType.AddItem "Excel 8"
    cbcFileType.AddItem "Rich Text"
    cbcFileType.AddItem "Tab separated text"
    cbcFileType.AddItem "Paginated Text"
    cbcFileType.AddItem "Word for Windows"
    cbcFileType.ListIndex = 0
     
End Function

