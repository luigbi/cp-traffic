Attribute VB_Name = "DDFREORGSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: DDFReorg.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the DDF Reorganize data code
Option Explicit
Option Compare Text
'This works but is not any faster then using VB code (gPDNToStr)
'Declare Sub ePDNToStr Lib "CSI.DLL" (ByVal i As String, ByVal j As Integer, ByVal k As Integer, ByVal o As String)
'Declare Function GetPrivateProfileString% Lib "kernel" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%, ByVal lpFileName$)
Declare Function GetPrivateProfileString& Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize As Long, ByVal lpFileName$)
'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function SendMessage& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Declare Function SendMessageByNum& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam&)
Declare Function SendMessageByString& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam$)
'Declare Sub HMemCpy Lib "kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Declare Sub btrCreClear Lib "cbtrvbx.dll" (ByVal Ohnd%)
'Declare Sub btrCreClearCollate Lib "cbtrvbx.dll" (ByVal Ohnd%)
'Declare Sub btrCreClearFile Lib "cbtrvbx.dll" (ByVal Ohnd%)
'Declare Sub btrCreClearKey Lib "cbtrvbx.dll" (ByVal Ohnd%)
'Declare Function btrCreCollate Lib "cbtrvbx.dll" (ByVal Ohnd%, ByVal SequenceName$, CollateSeq As Any) As Integer
'Declare Function btrCreCollateFile Lib "cbtrvbx.dll" (ByVal Ohnd%, ByVal CollateFileName$) As Integer
'Declare Function btrCreFile Lib "cbtrvbx.dll" (ByVal Ohnd%, ByVal RecLength%, ByVal PageSize%, ByVal NumIndexes%, ByVal uFileFlags%, ByVal uAllocation%) As Integer
'Declare Function btrCreKey Lib "cbtrvbx.dll" (ByVal Ohnd%, ByVal KeyPosition%, ByVal KeyLength%, ByVal KeyFlags%, ByVal ExtendKeyType%, ByVal NullValue%) As Integer
'Declare Function btrCreCreate Lib "cbtrvbx.dll" (ByVal Ohnd%, ByVal FileName$, ByVal fOverWrite%) As Integer
Declare Sub HMemCpy Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'TTP 10334 - last file to get processed doesn't remove the _.NEW file
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
' File Flag Defines
Public Const BTRV_FF_VARIABLE = 1
Public Const BTRV_FF_BLANK_TRUNC = 2
Public Const BTRV_FF_PREALLOCATION = 4
Public Const BTRV_FF_DATA_COMPRESS = 8
Public Const BTRV_FF_KEY_ONLY = 16
Public Const BTRV_FF_FREE_SPACE10 = 64
Public Const BTRV_FF_FREE_SPACE20 = 128
Public Const BTRV_FF_FREE_SPACE30 = 192
' Key Flag Defines
Public Const BTRV_KF_DUPLICATE = 1
Public Const BTRV_KF_MODIFIABLE = 2
Public Const BTRV_KF_BINARY = 4
Public Const BTRV_KF_NULL = 8
Public Const BTRV_KF_SEGMENTED = 16
Public Const BTRV_KF_ALT_COL_SEQ = 32
Public Const BTRV_KF_DESCENDING = 64
Public Const BTRV_KF_SUPPLEMENTAL = 128
Public Const BTRV_KF_EXTENDED = 256
Public Const BTRV_KF_MANUAL = 512
Public Const INDEXKEY0 = 0
Public Const INDEXKEY1 = 1
Public Const INDEXKEY2 = 2
 
Public Const DEFAULT = 0           ' 0 - Default
Public Const HOURGLASS = 11        ' 11 - Hourglass
' MsgBox parameters
Public Const vbOkOnly = 0                 ' OK button only
Public Const vbOkCancel = 1           ' OK and Cancel buttons
Public Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Public Const vbYesNoCancel = 3        ' Yes, No, and Cancel buttons
Public Const vbYesNo = 4              ' Yes and No buttons
Public Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons
Public Const vbCritical = 16          ' Critical message
Public Const vbQuestion = 32      ' Warning query
Public Const vbExclamation = 48   ' Warning message
Public Const vbInformation = 64   ' Information message
Public Const MB_DEFBUTTON1 = 0         ' First button is default
Public Const vbDefaultButton2 = 256       ' Second button is default
Public Const MB_DEFBUTTON3 = 512       ' Third button is default
Public Const vbApplicationModal = 0
Public Const MB_SYSTEMMODAL = 4096
' MsgBox return values
Public Const vbOk = 1                  ' OK button pressed
Public Const vbCancel = 2              ' Cancel button pressed
Public Const ID_ABORT = 3               ' Abort button pressed
Public Const ID_RETRY = 4               ' Retry button pressed
Public Const ID_IGNORE = 5              ' Ignore button pressed
Public Const vbYes = 6                 ' Yes button pressed
Public Const vbNo = 7                  ' No button pressed
'Common Dialog Control
'Action Property
Public Const DLG_FILE_OPEN = 1
Public Const DLG_FILE_SAVE = 2
'Global Const DLG_COLOR = 3
'Global Const DLG_FONT = 4
Public Const DLG_PRINT = 5
Public Const DLG_HELP = 6
'File Open/Save Dialog Flags
Public Const OFN_READONLY = &H1&
Public Const cdlOFNOverwritePrompt = &H2&
Public Const cdlOFNHideReadOnly = &H4&
Public Const cdlOFNNoChangeDir = &H8&
Public Const OFN_SHOWHELP = &H10&
Public Const OFN_NOVALIDATE = &H100&
Public Const OFN_ALLOWMULTISELECT = &H200&
Public Const OFN_EXTENSIONDIFFERENT = &H400&
Public Const cdlOFNPathMustExist = &H800&
Public Const OFN_FILEMUSTEXIST = &H1000&
Public Const cdlOFNCreatePrompt = &H2000&
Public Const OFN_SHAREAWARE = &H4000&
Public Const CdlOFNNoReadOnlyReturn = &H8000&
'Printer Dialog Flags
Public Const PD_ALLPAGES = &H0&
Public Const PD_SELECTION = &H1&
Public Const PD_PAGENUMS = &H2&
Public Const PD_NOSELECTION = &H4&
Public Const PD_NOPAGENUMS = &H8&
Public Const PD_COLLATE = &H10&
Public Const PD_PRINTTOFILE = &H20&
Public Const cdlPDPrintSetup = &H40&
Public Const PD_NOWARNING = &H80&
Public Const PD_RETURNDC = &H100&
Public Const PD_RETURNIC = &H200&
Public Const PD_RETURNDEFAULT = &H400&
Public Const PD_SHOWHELP = &H800&
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const CF_LINK = &HBF00   'DDE Conversation information
Public Const vbCFText = 1        'Text format
Public Const CF_BITMAP = 2      'Bitmap (.BMP) files
Public Const CF_METAFILE = 3    'Metafile (.WMP files)
'  Edit Control Messages
Public Const WM_USER = &H400
Public Const EM_UNDO = WM_USER + 23
Public Const LB_SELITEMRANGE = (WM_USER + 28)
Public Const EM_SETREADONLY = WM_USER + 31
'Help Constants
Public Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Public Const HELP_QUIT = &H2              'Terminate help
Public Const cdlHelpIndex = &H3             'Display index
Public Const HELP_CONTENTS = &H3
Public Const cdlHelpHelpOnHelp = &H4        'Display help on using help
Public Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Public Const HELP_SETCONTENTS = &H5
Public Const HELP_CONTEXTPOPUP = &H8
Public Const HELP_FORCEFILE = &H9
Public Const HELP_KEY = &H101             'Display topic for keyword in offabData
Public Const HELP_COMMAND = &H102
Public Const cdlHelpPartialKey = &H105      'call the search engine in winhelp
'Declare Function AbortDoc Lib "GDI" (ByVal hDC As Integer) As Integer
'Rm**Declare Sub ISortT2 Lib "QPRO200.DLL" (Array As Any, Index%, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%)
'Rm**Declare Sub ArraySortTyp Lib "QPRO200.DLL" (Array() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%, ByVal CaseSenitive%)
'ArraySortTyp fnAV(tlSort(),0), ilUpper, 0, LenB(tlSort(0)), 0, Len(tlSort(0).sKey), 0
'Rm**'Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**'Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**'Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Type INT2
    sInt2 As String * 2
End Type
Type INT4
    sInt4 As String * 4
End Type
Type BASEREC
    sChar(1 To 32000) As Byte   'String * 1 'Record
End Type
Type KEYREC
    sChar(1 To 1000) As Byte    'String * 1
End Type
Type KEYFILE
    iFileID As Integer
    hFile As Integer
End Type
Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
End Type
Type DDFFIELD
    iFieldID As Integer         'Field ID
    iFileID As Integer          'File ID from FILE.DDF
    sName As String * 20        'Field Name
    sDataType As String * 1     'Data Type Code (Use ASC to convert)
                                'Code  Description
                                '  0   String
                                '  1   Integer
                                '  2   IEEE Float
                                '  3   Btrieve Date
                                '  4   Btrieve Time
                                '  5   COBOL Decimal COMP-3
                                '  6   COBOL Money
                                '  7   Logical
                                '  8   COBOL Numeric
                                '  9   BASIC BFloat
                                ' 10   Pascal LString
                                ' 11   C ZString
                                ' 12   Variable Length Note
                                ' 13   LVar
                                ' 14   Unsigned Binary
                                ' 15   AutoIncrement
                                ' 16   Bit
                                ' 17   COBOL Numeric STS
    iOffset As Integer          'Field Offset
    iSize As Integer            'Field Size
    sDec As String * 1          'Decimal/Delimiter/Bit Position
    iFlags As Integer           'Case Flag for String data type
End Type
Type DDFFIELD1
    iFileID As Integer          'File ID from FILE.DDF
End Type
Type DDFINDEX
    iFileID As Integer          'File ID from FILE.DDF
    iFieldID As Integer         'Field ID from FIELD.DDF
    iNumber As Integer          'Index ID Number within file
    iPart As Integer            'Segment ID number within index
    iFlag As Integer            'Btrieve Index Flag
End Type
Type DDFINDEX0
    iFileID As Integer          'File ID from FILE.DDF
End Type
Type DDFINDEX2
    iFileID As Integer          'File ID from FILE.DDF
    iNumber As Integer          'Index ID Number within file
    iPart As Integer            'Segment ID number within index
End Type
Type INITVALUES
    sName As String             'Field name
    sInitValue As String        'Initial value or Field Name from FieldSet.New
    iOper As Integer            'Operation if Field name (0=Non; 1=Add; 2=Sub; 3= Mul; 4=Div)
    sValue As String            'Value from FieldSet.New
End Type
Type ADJUSTVALUES
    sName As String             'Field name
    iOper As Integer            'Operation (0=Non; 1=Add; 2=Sub; 3= Mul; 4=Div)
    sValue As String            'Value from FieldSet.New
End Type
Public tgNewRec As BASEREC

'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
Public Const OF_EXIST         As Long = &H4000
Public Const OFS_MAXPATHNAME  As Long = 128
Public Const HFILE_ERROR      As Long = -1
Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Declare Function OpenFile Lib "Kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
 
'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveZero                     *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove trialing zeros           *
'*                                                     *
'*******************************************************
Function gRemoveZero(slInStr As String) As String
    Dim slStr As String
    Dim slChar As String
    Dim ilPos As Integer
    ilPos = 1
    slStr = ""
    Do While ilPos <= Len(slInStr)
        slChar = Mid$(slInStr, ilPos, 1)
        If Asc(slChar) = 0 Then
            gRemoveZero = slStr
            Exit Function
        End If
        slStr = slStr & slChar
        ilPos = ilPos + 1
    Loop
    gRemoveZero = slStr
    Exit Function
End Function
