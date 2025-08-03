Attribute VB_Name = "crwrap"
'
'               Visual Basic Declarations of crwrap32.DLL
'               =====================================
'
'       File:         crwrap.BAS
'
'       Author:       Seagate Software Information Management Group, Inc.
'
'
'       Language:     Visual Basic for Windows
'
'       Copyright (c) 1992 - 1999 Seagate Software Information Management Group, Inc.
'
'       Revisions:
'       SM- Nov 99, removed obsolete exporting formats
'           Excel 2.1
'            Global Const crUXFXls2Type% = 0
'           Excel 3.0
'            Global Const crUXFXls3Type% = 1
'           Excel 4.0
'            Global Const crUXFXls4Type% = 2
'           HTML 3.0
'            Global Const crUXFHTML3Type% = 0   ' Draft HTML 3.0 tags
'
'
'================================================================================
' Constants and Function Definitions for Exporting
'================================================================================
'******************
'** Format Types **
'******************

'Separated Values           (DLL: "uxfsepv.dll")
Global Const crUXFCommaSeparatedType% = 200

'Data Interchange Format    (DLL: "uxfdif.dll")
Global Const crUXFDIFType% = 400

'Record Style Format        (DLL: "uxfrec.dll")


'Crystal Report             (DLL: "uxfcr.dll")

'RTF                        (DLL: "uxfrtf.dll")
Global Const crUXFRichTextFormatType% = 0

'Text                       (DLL: "uxftext.dll")
Global Const crUXFTextType% = 0
Global Const crUXFTabbedTextType% = 1
Global Const crUXFPaginatedTextType% = 600

'Lotus                      (DLL: "uxfwks.dll")

'Word for Windows           (DLL: "uxfwordw.dll")
Global Const crUXFWordWinType% = 0

'Excel                      (DLL: "uxfxls.dll")
Global Const crUXFXls5Type% = 3

'Report Definition          (DLL: "uxfrdef.dll")

#If Win16 Then
'Word for DOS & WordPerfect (DLL: "uxfdoc.dll")

'Quattro Pro                (DLL: "uxfqp.dll")

#End If

'****************
'** HTML Types **
'****************

'HTML


'DateTime structure


'Devmode structure
Type crDEVMODE
    dmDriverVersion As Integer  ' printer driver version number (usually not required)
#If Win16 Then                  ' add padding so it aligns the same way under both 16-bit and 32-bit environments
    pad1 As Integer
#End If
    dmFields As Long           'flags indicating fields to modify (required)
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
End Type

'/* field selection bits */

'/* orientation selections */

'/* paper selections */
'/*  Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' *  50 and 56.  Don't use this range when defining new paper sizes.
' */
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41  '/* German Legal Fanfold 8 1/2 x 13 in */



'/* bin selections */
Global Const DMBIN_CASSETTE = 14


'/* print qualities */

'/* color enable/disable for color printers */

'/* duplex enable */

'/* TrueType options */



'================================================================================
' Type Declarations for 4-byte aligned functions
'================================================================================
#If Win32 Then
    Type PEJobInfo4
        StructSize As Integer  ' initialize to PE_SIZEOF_JOB_INFO

        NumRecordsRead As Long
        NumRecordsSelected As Long
        NumRecordsPrinted As Long

        DisplayPageN As Integer
        LatestPageN As Integer
        StartPageN As Integer

        PrintEnded As Long
    End Type

    Type PETableType4
        StructSize As Integer   ' initialize to # bytes in PETableType

        DLLName As String * PE_DLL_NAME_LEN
        DescriptiveName  As String * PE_FULL_NAME_LEN

        DBType As Integer
    End Type

    Type PELogOnInfo4
        ' initialize to # bytes in PELogOnInfo
        StructSize As Integer

        ' For any of the following values an empty string ("") means to use
        ' the value already set in the report.  To override a value in the
        ' report use a non-empty string (e.g. "Server A").
        '
        ' For Netware SQL, pass the dictionary path name in ServerName and
        ' data path name in DatabaseName.

        ServerName As String * PE_SERVERNAME_LEN
        DatabaseName  As String * PE_DATABASENAME_LEN
        UserID As String * PE_USERID_LEN

        ' Password is undefined when getting information from report.

        Password  As String * PE_PASSWORD_LEN
    End Type


    Type PESessionInfo4
        'initialize to # bytes in PESessionInfo
        StructSize As Integer

        ' Password is undefined when getting information from report.
        UserID As String * PE_SESS_USERID_LEN
        Password As String * PE_SESS_PASSWORD_LEN

        ' SessionHandle is undefined when getting information from report.
        ' When setting information, if it is = 0 the UserID and Password
        ' settings are used, otherwise the SessionHandle is used.
        SessionHandle As Long
    End Type


    Type PEGraphOptions4
        StructSize     As Integer  ' initialize to # bytes in PEGraphOptions
        GraphMaxValue  As Double
        GraphMinValue  As Double
        ShowDataValue  As Long  ' Show data values on risers.
        ShowGridLine   As Long
        VerticalBars   As Long
        ShowLegend     As Long
        FontFaceName   As String * PE_GRAPH_TEXT_LEN
    End Type

#End If

'================================================================================
' Function Declarations
'================================================================================
#If Win16 Then
    '** new CRPE wrapper Functions


    '** Export Functions **
    Declare Function crPEExportToDisk Lib "crwrap16.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer

#ElseIf Win32 Then
    '** new CRPE wrapper Functions


    '** Export Functions **
    Declare Function crPEExportToDisk Lib "crwrap32.dll" (ByVal printJob%, ByVal fileName$, ByVal FormatDLLName$, ByVal formatType As Long, ByVal useNumFormat As Long, ByVal useDateFormat As Long, ByVal StringDelimiter$, ByVal FieldDelimiter$) As Integer

    '** Functions which accept VB's native 4-byte aligned "Types"

#End If

