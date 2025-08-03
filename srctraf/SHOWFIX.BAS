Attribute VB_Name = "SHOWFIXSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ShowFix.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Show/Fix Product screen code
Option Explicit
Option Compare Text
'This works but is not any faster then using VB code (gPDNToStr)
'Declare Sub ePDNToStr Lib "CSI.DLL" (ByVal i As String, ByVal j As Integer, ByVal k As Integer, ByVal o As String)
'Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize%, ByVal lpFileName$)
'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
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
'Color &HBBGGRR
' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
Public Const BLACK = &H0&
Public Const RED = &HFF&
Public Const GREEN = &HFF00&
Public Const YELLOW = &HFFFF&
Public Const BLUE = &HFF0000
Public Const MAGENTA = &HFF00FF
Public Const CYAN = &HFFFF00
Public Const WHITE = &HFFFFFF
' Extra background colors
Public Const DARKBLUE = &H800000
Public Const DARKGREEN = &H8000&
Public Const DARKCYAN = &H808000
Public Const DARKRED = &H80&
Public Const DARKPURPLE = &H800080
Public Const DARKYELLOW = &H8080&
Public Const DARKGRAY = &H808080
Public Const GRAY = &HC0C0C0
Public Const PURPLE = &HFF00FF
Public Const LIGHTYELLOW = &HBFFFFF '&H80FFFF '&HBFFFFF     'Non-edit field color
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
Declare Function AbortDoc Lib "GDI" (ByVal hDC As Integer) As Integer
'Rm**Declare Sub ISortT2 Lib "QPRO200.DLL" (Array As Any, Index%, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%)
'Rm**Declare Sub ArraySortTyp Lib "QPRO200.DLL" (Array() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%, ByVal CaseSenitive%)
'ArraySortTyp fnAV(tlSort(),0), ilUpper, 0, LenB(tlSort(0)), 0, Len(tlSort(0).sKey), 0
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24
Public Const MODAL = 1
'Drag
Public Const DRAGCANCEL = 0
Public Const DRAGSTART = 1
Public Const DRAGEND = 2
Public Const DRAGENTER = 0
Public Const DRAGLEAVE = 1
Public Const DRAGOVER = 2
'Shift values
Public Const SHIFT1 = &H2
Public Const SHIFT2 = &H4
Public Const SHIFT3 = &H8
Public Const SHIFT4 = &H10
Public Const SHIFT5 = &H20
Public Const SHIFT6 = &H40
Public Const SHIFT7 = &H80
Public Const SHIFT8 = &H100
Public Const SHIFT9 = &H200
Public Const SHIFT10 = &H400
Public Const SHIFT11 = &H800
Public Const SHIFT12 = &H1000
Public Const SHIFT13 = &H2000
Public Const SHIFT14 = &H4000
Public Const SHIFT15 = &H8000
Public Const SHIFT16 = &H10000
Public Const SHIFT17 = &H20000
Public Const SHIFT18 = &H40000
Public Const SHIFT19 = &H80000
Public Const SHIFT20 = &H100000
Public Const SHIFT21 = &H200000
Public Const SHIFT22 = &H400000
Public Const SHIFT23 = &H800000
Public Const SHIFT24 = &H1000000
Public Const SHIFT25 = &H2000000
Public Const SHIFT26 = &H4000000
Public Const SHIFT27 = &H8000000
Public Const SHIFT28 = &H10000000
Public Const SHIFT29 = &H20000000
Public Const SHIFT30 = &H40000000
'Spot Summary flags
Public Const SSLOCK = &H40 'Avail Locked
Public Const SSLOCKSPOT = &H80    'Spots locked
Public Const SSSUSTAINING = &H100
Public Const SSSPONSORSHIP = &H200
Public Const SSOPENBB = &H10
Public Const SSCLOSEBB = &H20
Public Const SSFLOATERBB = &H40
Public Const SSANYBB = &H80
Public Const SSDONUT = &H100
Public Const SSBOOKEND = &H200
Public Const SSAVAILBUY = &H400
Public Const SSPREEMPTIBLE = &H800
Public Const SSEXAVAILBUY = &H1000
Public Const SSLIBBUY = &H2000
Public Const SSEXCLUSIONS = &H4000

Global Const SSSOLOAVAIL = &H20000000
Global Const SS1STPOSITION = &H40000000

Public Const RANKMASK = &H7FF
Public Const PRICELEVELMASK = &H7800

' Button and Shift (KeyDown, KeyUp, MouseDown, MouseMove, MouseUp)
Public Const vbShiftMask = 1
Public Const vbCtrlMask = 2
Public Const vbAltMask = 4
Public Const vbLeftButton = 1
Public Const vbRightButton = 2
Public Const KEYLBUTTON = &H1
Public Const KEYRBUTTON = &H2
Public Const KEYCANCEL = &H3
'Global Const KEYMBUTTON = &H4         ' NOT contiguous with L & RBUTTON
Public Const KEYBACK = &H8
Public Const KEYTAB = &H9
'Global Const KEYCLEAR = &HC
'Global Const KEYRETURN = &HD
Public Const KEYSHIFT = &H10
Public Const KEYCONTROL = &H11
Public Const KEYMENU = &H12
'Global Const KEYPAUSE = &H13
'Global Const KEYCAPITAL = &H14
Public Const KEYESCAPE = &H1B
Public Const KEYSPACE = &H20
'Global Const KEYPRIOR = &H21
'Global Const KEYNEXT = &H22
'Global Const KEYEND = &H23
'Global Const KEYHOME = &H24
Public Const KEYLEFT = &H25
Public Const KEYUP = &H26
Public Const KEYRIGHT = &H27
Public Const KEYDOWN = &H28
'Global Const KEYSELECT = &H29
'Global Const KEYPRINT = &H2A
'Global Const KEYEXECUTE = &H2B
'Global Const KEYSNAPSHOT = &H2C
Public Const KEYINSERT = &H2D
Public Const KEYDELETE = &H2E
Public Const KEYHELP = &H2F
' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'
'Global Const KEYNUMPAD0 = &H60
'Global Const KEYNUMPAD1 = &H61
'Global Const KEYNUMPAD2 = &H62
'Global Const KEYNUMPAD3 = &H63
'Global Const KEYNUMPAD4 = &H64
'Global Const KEYNUMPAD5 = &H65
'Global Const KEYNUMPAD6 = &H66
'Global Const KEYNUMPAD7 = &H67
'Global Const KEYNUMPAD8 = &H68
'Global Const KEYNUMPAD9 = &H69
'Global Const KEYMULTIPLY = &H6A
'Global Const KEYADD = &H6B
'Global Const KEYSEPARATOR = &H6C
'Global Const KEYSUBTRACT = &H6D
'Global Const KEYDECIMAL = &H6E
'Global Const KEYDIVIDE = &H6F
Public Const KEYF1 = &H70
'Global Const KEYF2 = &H71
'Global Const KEYF3 = &H72
Global Const KEYF4 = &H73
Global Const KEYF5 = &H74
'Global Const KEYF6 = &H75
'Global Const KEYF7 = &H76
'Global Const KEYF8 = &H77
'Global Const KEYF9 = &H78
'Global Const KEYF10 = &H79
'Global Const KEYF11 = &H7A
'Global Const KEYF12 = &H7B
'Global Const KEYF13 = &H7C
'Global Const KEYF14 = &H7D
'Global Const KEYF15 = &H7E
'Global Const KEYF16 = &H7F
'Global Const KEYNUMLOCK = &H90
'Number key filters
Public Const KEYBACKSPACE = 8   'Back space key pressed
Public Const KEYPOS = 43        '+ key pressed
Public Const KEYCOMMA = 44      ', key pressed
Public Const KEYNEG = 45        '- key pressed
Public Const KEYDECPOINT = 46   'Decimal point pressed
Public Const KEY0 = 48          '0 key pressed
Public Const KEY9 = 57          '9 key pressed
Public Const KEYUA = 65         'Upper case A
Public Const KEYUZ = 90         'Upper case Z
Public Const KEYLA = 97         'Lower case a
Public Const KEYLZ = 122        'Lower case z
Public Const KEYASTERISK = 42   'Asterisk key pressed
Public Const KEYSLASH = 47      '/ Slash key

Type INT2
    sInt2 As String * 2
End Type
Type INT4
    sInt4 As String * 4
End Type
Type BYTE4
    bByte(0 To 3) As Byte
End Type
Type BASEREC
    sChar(1 To 32000) As Byte 'Record
End Type
Type KEYREC
    sChar(1 To 1000) As Byte
End Type
Type RECINFO
    lRecPos As Long                 'Record position
    iRecLen As Integer                 'Record length
    iRecIndex1 As Integer
    iRecIndex2 As Integer
    iRowLinkIndex1 As Integer
    iRowLinkIndex2 As Integer
    lRecNo As Long
End Type
Type ROWLINKLIST
    lRowNo As Long
    iGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iFirstIndex1 As Integer
    iFirstIndex2 As Integer
    iRecIndex1 As Integer   'Record index
    iRecIndex2 As Integer   'Record index
End Type
Type GRID1DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 10
End Type
Type GRID2DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 16
End Type
Type GRID3DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 128
End Type
Type GRID4DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 1024
End Type
Type GRID5DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 4096
End Type
Type GRID6DATA
    iNextIndex1 As Integer
    iNextIndex2 As Integer
    iNextGridType As Integer    '0-None; 1=GRID1DATA; 2=GRID2DATA
    iDDFFieldIndex As Integer
    sValue As String * 32000
End Type
Type SELCOMP
    iFieldIndex1 As Integer         'tmDFFField Index
    sValue1 As String          'Not Fixed Length String
    iOper As Integer            '-1=None, 0="=";1="<>"; 2"<"; 3=">"; 4="<="; 5=">="
    iFieldIndex2 As Integer         'tmDFFField Index
    sValue2 As String          'Not Fixed Length String
    iResult As Integer         'True or False
End Type
Type LINK
    iKeyNo As Integer       'Key index
    iToFileID As Integer      'File ID
    iToFieldID As Integer
    iKeyOffset As Integer
    iFromFileID As Integer
    iFromFieldID As Integer
    iFromOffset As Integer
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
    iNumber As Integer          'Index Number
    iPart As Integer            'Segment Part Number
    iFlag As Integer            'Btrieve Index Flag
End Type
Type DDFINDEX0
    iFileID As Integer          'File ID from FILE.DDF
End Type
Public tgKeyRec As BASEREC
Public tgKey As KEYREC
Public fgListHTArial825 As Single
Public fgListHtSerif825 As Single
Public igAlignCharWidth As Integer
Public igScrollBarWidth As Integer
Public fgBoxInsetX As Single
Public fgBoxGridH As Single
Public igPasswordOk As Integer
Public tgUrf() As URF 'User Option file
Public igBkgdProg As Integer    '0=No; 1=Scheduler; 2=Set Credit; 10=Initializing Background Mode (Log errors to TrafficErrors.txt)


'Public sgDBPath As String
'*******************************************************
'*                                                     *
'*      Procedure Name:gParseItemFields                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain a substring from a string*
'*                                                     *
'*******************************************************
Sub gParseItemFields(ByVal slInputStr As String, slDelimiter As String, slFields() As String)
'
'   gParseItemToArray slInputStr, slDelimiter, slFields()
'   Where:
'       slInputStr (I)-string from which to obtain substring
'       slFields() (O)- fields parsed from comma delimited string
'

    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
        slFields(ilIndex) = ""
    Next ilIndex
    ilFieldNo = 1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slInputStr))
        slChar = Mid$(slInputStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = slDelimiter Then
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    Exit Sub
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                ilFieldType = 0 'Quote field
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" Then
                    If ilIndex = Len(Trim$(slInputStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slInputStr, ilIndex + 1, 1)
                        If slNextChar = slDelimiter Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> slDelimiter Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gAdjShowLen                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Adjust length of string to fit *
'*                 within specified area minus margins *
'*                                                     *
'*******************************************************
Sub gAdjShowLen(Ctrl As Control, slShow As String, flWidth As Single)
'
'   gAdjShowLen Ctrl, sShow, BoxWidth
'   Where:
'       Ctrl (I)- Control where string will be painted
'       sShow (I/O)- String to be shown
'       ilBoxWidth (I)- width of control area
'

    Dim flPicWidth
    flPicWidth = flWidth - 2 * fgBoxInsetX
    Do While (Ctrl.TextWidth(slShow) > flPicWidth)
        slShow = Left$(slShow, Len(slShow) - 1)
    Loop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintArea                      *
'*                                                     *
'*             Created:7/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear area                      *
'*                                                     *
'*******************************************************
Sub gPaintArea(pbcCtrl As Control, flCurrX As Single, flCurrY As Single, flWidth As Single, flHeight As Single, llPaintColor As Long)
'
'   gPaintArea pbcCtrl, lX, lY, lW, lH, lC
'   Where:
'       pbcCtrl (I)- Picture box control containing area to be painted
'       lX (I)- Left position of area to be painted
'       lY (I)- Top position of area to be painted
'       lH (I)- Height of area to be painted
'       lW (I)- Width of area to be painted
'       lC (I)- Color of area to be painted
'

    pbcCtrl.CurrentX = flCurrX
    pbcCtrl.CurrentY = flCurrY
    pbcCtrl.Line Step(0, 0)-Step(flWidth, flHeight), llPaintColor, BF
End Sub

Public Function gSSFGetDirect(hlSsf As Integer, tlSsf As SSF, llSsfRecLen As Long, llRecPos As Long, ilKeyNo As Integer, ilLock As Integer) As Integer
    Dim ilRet As Integer
    
    'imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    'ReDim bgByteArray(LenB(tmSsf))
    'HMemCpy bgByteArray(0), tmSsf, LenB(tmSsf)
    'ilRet = btrGetDirect(hmSsf, bgByteArray(0), imSsfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    'HMemCpy tmSsf, bgByteArray(0), LenB(tmSsf)
    gSSFGetDirect = btrGetDirect(hlSsf, tlSsf, llSsfRecLen, llRecPos, ilKeyNo, ilLock)
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCompNumberStr                  *
'*                                                     *
'*             Created:10/07/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compare two number strings.     *
'*                     The strings are converted to    *
'*                     numeric values, then compared.  *
'*                                                     *
'*******************************************************
Function gCompNumberStr(slCompStr1 As String, slCompStr2 As String) As Integer
'
'   iRet = gCompNumberStr(sStr1, sStr2)
'   Where:
'       sStr1 (I)- String to be compared
'       sStr2 (I)- String to compare
'       iRet (O)- 0 if sStr1=sStr2; -1 if sStr1 < sStr2; 1 If sStr1 > sStr2
'
    Dim clNum1 As Currency
    Dim clNum2 As Currency
    Dim slInput1 As String
    Dim slInput2 As String
    slInput1 = Trim$(slCompStr1)
    slInput2 = Trim$(slCompStr2)
    clNum1 = Val(slInput1)
    'If clNum1 < 0 Then
    '    clNum1 = -clNum1
    'End If
    clNum2 = Val(slInput2)
    'If clNum2 < 0 Then
    '    clNum2 = -clNum2
    'End If
    If clNum1 = clNum2 Then
        gCompNumberStr = 0
    ElseIf clNum1 > clNum2 Then
        gCompNumberStr = 1
    Else
        gCompNumberStr = -1
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gLengthToCurrency               *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert length to currency      *
'*                     (for precision -                *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gLengthToCurrency(slInpLength As String) As Currency
'
'   clRetLength = gLengthToCurrency(slLength)
'   Where:
'       slLength (I)- Length as string to be converted to currency
'       clRetLength (O)- length as currency
'
    Dim slLength As String
    Dim clLength As Currency
    Dim ilPos As Integer
    Dim ilLoc As Integer
    Dim slAnyLength As String
    slAnyLength = Trim$(slInpLength)
    On Error GoTo gLengthToCurrencyErr
    ilPos = InStr(slAnyLength, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gLengthToCurrency = 0
            Exit Function
        End If
        slLength = UCase$(Mid$(slAnyLength, 2))
    Else
        slLength = UCase$(slAnyLength)
    End If
    clLength = 0
    ilLoc = InStr(1, slLength, "H", 1)
    If ilLoc <> 0 Then
        clLength = clLength + 3600 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "M", 1)
    If ilLoc <> 0 Then
        clLength = clLength + 60 * Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    End If
    ilLoc = InStr(1, slLength, "S", 1)
    If ilLoc <> 0 Then
        clLength = clLength + Val(Left$(slLength, ilLoc - 1))
        slLength = Mid$(slLength, ilLoc + 1)
    Else
        clLength = clLength + Val(slLength)
    End If
    If ilPos = 0 Then
        gLengthToCurrency = clLength
    Else
        gLengthToCurrency = -clLength
    End If
    Exit Function
gLengthToCurrencyErr:
    On Error GoTo 0
    gLengthToCurrency = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gMonthYearFormat                *
'*                                                     *
'*             Created:11/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain month name, year from    *
'*                     Date                            *
'*                                                     *
'*******************************************************
Function gMonthYearFormat(slInpDate As String) As String
'
'   sRetDate = gMonthYearFormat(sDate)
'   Where:
'       sDate (I)- Date string for which month name and year are obtained
'       sRetDate (O)- Month name (3 characters), Year
'
    Dim ilMonth As Integer
    Dim slStr As String
    Dim slDate As String
    Dim slAnyDate As String
    slAnyDate = Trim$(slInpDate)
    slDate = gAdjYear(slAnyDate)
    On Error GoTo gMonthYearFormatErr
    ilMonth = Month(slDate)
    Select Case ilMonth
        Case 1
            slStr = "Jan, "
        Case 2
            slStr = "Feb, "
        Case 3
            slStr = "Mar, "
        Case 4
            slStr = "Apr, "
        Case 5
            slStr = "May, "
        Case 6
            slStr = "June, "
        Case 7
            slStr = "July, "
        Case 8
            slStr = "Aug, "
        Case 9
            slStr = "Sept, "
        Case 10
            slStr = "Oct, "
        Case 11
            slStr = "Nov, "
        Case 12
            slStr = "Dec, "
    End Select
    gMonthYearFormat = slStr & Trim$(Str$(Year(slDate)))
    Exit Function
gMonthYearFormatErr:
    gMonthYearFormat = ""
    On Error GoTo 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMonthYear                *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Month and year given date*
'*                                                     *
'*******************************************************
Sub gObtainMonthYear(ilType As Integer, slInpDate As String, ilMonth As Integer, ilYear As Integer)
'
'   gObtainMonthYear ilType, sDate, ilMonth, ilYear
'   Where:
'       ilType (I)- 0=Standard month; 1= Regular month; 4=Corp
'       sDate (I)- Date to obtain Month and year
'       ilMonth (O) - Month
'       ilYear (O) - Year
'
    Dim llDate As Long
    Dim ilMatchDay As Integer
    Dim slDate As String
    Dim slAnyDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    slAnyDate = Trim$(slInpDate)
    If ilType = 0 Then  'Std
        'slDate = gObtainEndStd(slAnyDate)
        'slDate = gAdjYear(slDate)
    ElseIf ilType = 4 Then  'Std
        'slStartDate = gObtainStartCorp(slAnyDate, True)
        'slEndDate = gObtainEndCorp(slAnyDate, True)
        'slDate = Format$(gDateValue(slStartDate) + (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 2, "m/d/yy")
        'slDate = gAdjYear(slDate)
    Else
        slDate = gAdjYear(slAnyDate)
    End If
    ilMonth = Month(slDate)
    ilYear = Year(slDate)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gPackLength                      *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Pack length in btrieve format   *
'*                                                     *
'*******************************************************
Sub gPackLength(slInpTime As String, ilHsSec As Integer, ilMinHr As Integer)
'
'   gPackLength slTime, ilHsSec, ilMinHr
'   Where:
'       slTime (I) - Length of Time to be packed in btrieve format
'       ilHsSec (O)- High order byte = hundredths of seconds; low order byte =                  '       seconds
'       ilMinHr (O)- High order byte = minute; low order byte = hours
'
    Dim ilPos As Integer
    Dim slLen As String
    Dim slHour As String
    Dim ilHour As Integer
    Dim slMin As String
    Dim ilMin As Integer
    Dim slSec As String
    Dim ilSec As Integer
    Dim ilFormat As Integer
    Dim slTime As String
    slTime = Trim$(slInpTime)
    If Len(slTime) = 0 Then
        ilHsSec = 1 'High order byte = hundredths of seconds; low order byte = seconds
        ilMinHr = 0 'High order byte = minute; low order byte = hours
        Exit Sub
    End If
    slHour = ""
    slMin = ""
    slSec = ""
    slLen = Trim$(slTime)
    ilPos = InStr(1, slLen, ":")
    If ilPos > 0 Then
        ilFormat = 1
    Else
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            ilFormat = 3
        Else
            ilPos = InStr(1, slLen, "m", 1)
            If ilPos > 0 Then
                ilFormat = 3
            Else
                ilPos = InStr(1, slLen, "s", 1)
                If ilPos > 0 Then
                    ilFormat = 3
                Else
                    ilFormat = 2
                End If
            End If
        End If
    End If
    If ilFormat = 2 Then 'hh mm'ss"
        ilPos = InStr(1, slLen, " ")
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "'")
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            End If
         Else
            ilPos = InStr(1, slLen, """")
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                End If
            End If
         End If
    ElseIf ilFormat = 3 Then 'hhHmmMssS
        ilPos = InStr(1, slLen, "h", 1)
        If ilPos > 0 Then
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
        End If
        ilPos = InStr(1, slLen, "m", 1)
        If ilPos > 0 Then
            slMin = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            End If
         Else
            ilPos = InStr(1, slLen, "s", 1)
            If ilPos > 0 Then
                slSec = Left$(slLen, ilPos - 1)
            Else
                If slHour = "" Then
                    slHour = slLen
                End If
            End If
         End If
    Else    'format hh:mm:ss
        ilPos = InStr(slLen, ":")
        If ilPos > 0 Then   'Might be hour/min/sec or hour/min only
            slHour = Left$(slLen, ilPos - 1)
            slLen = Mid$(slLen, ilPos + 1)
            ilPos = InStr(1, slLen, ":")
            If ilPos > 0 Then
                slMin = Left$(slLen, ilPos - 1)
                slSec = Mid$(slLen, ilPos + 1)
            Else
                slMin = slLen
            End If
        Else
            slHour = slLen
        End If
    End If
    If slHour <> "" Then
        ilHour = Val(slHour)
    Else
        ilHour = 0
    End If
    If slMin <> "" Then
        ilMin = Val(slMin)
    Else
        ilMin = 0
    End If
    If slSec <> "" Then
        ilSec = Val(slSec)
    Else
        ilSec = 0
    End If
    If ilSec > 59 Then
        ilMin = ilMin + ilSec \ 60
        ilSec = ilSec Mod 60
    End If
    If ilMin > 59 Then
        ilHour = ilHour + ilMin \ 60
        ilMin = ilMin Mod 60
    End If
    ilHsSec = ilSec * 256 'High order byte = hundredths of seconds; low order byte = seconds
    ilMinHr = ilMin + ilHour * 256 'High order byte = minute; low order byte = hours
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gPaintCalendar                  *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Paint standard or regular       *
'*                     calendar                        *
'*                                                     *
'*******************************************************
Sub gPaintCalendar(ilMonth As Integer, ilYear As Integer, ilType As Integer, pbcCtrl As PictureBox, tlCtrls() As FIELDAREA, llStartDate As Long, llEndDate As Long)
'
'   gPaintCalendar ilMonth, ilYear, ilType, pbcCtrl, tmCtrl, llStart, llEnd
'   Where:
'       ilMonth (I) - Calendar month to be painted (1 thru 12)
'       ilYear (I) - Calendar year to be painted (101 thru 9998) (100 & 9999 not                        '       allowed because of standard)
'       ilType (I)- 0=Paint standard month; 1= Paint regular month; 2=Julian +;
'                   3=Julian -; 4=Paint corporate
'       pbcCtrl (I)- Picture area control to be painted
'       tlCtrl() (I)- Array of control information about paint area
'       llStart (O)- Calendar start date
'       llEnd (O)- Calendar end date
'
    Static ilCalMonth As Integer
    Static ilCalYear As Integer
    Dim llDate As Long
    Dim llTempDate As Long
    Dim llStdDate As Long
    Dim slDate As String
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim slJulian As String
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim ilStdMonth As Integer
    Dim llLastDate As Long
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim flBoxInsetX As Single
    Dim flBoxInsetY As Single
    Dim ilAdjYear As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slLastDate As String
    If (ilYear < 101) Or (ilYear > 9998) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    If (ilMonth < 1) Or (ilMonth > 12) Then
        pbcCtrl.Cls
        Exit Sub
    End If
    If ilType = 0 Then  'Standard month
        slDate = Trim$(Str$(ilMonth)) & "/1/" & Trim$(Str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = Year(slDate)
            llDate = gDateValue(slDate)
            llStdDate = llDate
        Else
            ilAdjYear = Year(slDate)
            llDate = DateValue(slDate)
            llStdDate = llDate
        End If
        Do While gWeekDayLong(llDate) <> 0   '0=monday
            llDate = llDate - 1
        Loop
        Do
            If gWeekDayLong(llStdDate) = 6 Then  'Save last sunday
                llLastDate = llStdDate
            End If
            llStdDate = llStdDate + 1
        Loop Until Month(llStdDate) <> ilMonth
    ElseIf ilType = 4 Then  'Corporate
    Else
        slDate = Trim$(Str$(ilMonth)) & "/1/" & Trim$(Str$(ilYear))
        If ilYear < 100 Then
            slDate = gAdjYear(slDate)
            ilAdjYear = Year(slDate)
            llDate = gDateValue(slDate)
        Else
            ilAdjYear = Year(slDate)
            llDate = DateValue(slDate)
        End If
    End If
'    If gValidDate(slDate) = No Then
'        pbcCtrl.Cls
'        Exit Sub
'    End If
    If (ilMonth <> ilCalMonth) Or (ilAdjYear <> ilCalYear) Then
        pbcCtrl.Cls
        ilCalMonth = ilMonth
        ilCalYear = ilAdjYear
    End If
    ilRowNo = 0
    llStartDate = 0
    llEndDate = 0
    slFontName = pbcCtrl.FontName
    flFontSize = pbcCtrl.FontSize
    If (ilType <= 1) Or (ilType = 4) Then
        pbcCtrl.FontBold = True
        flBoxInsetX = fgBoxInsetX
        flBoxInsetY = -15
    Else
        pbcCtrl.FontBold = False
        pbcCtrl.FontSize = 7
        pbcCtrl.FontName = "Arial"
        pbcCtrl.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        flBoxInsetX = 5
        flBoxInsetY = 5
    End If
''    If (ilType <> 1) And (Weekday(llDate) <> 2) Then   'for julian by std
    llColor = pbcCtrl.ForeColor
'    If (ilType = 0) And (gWeekDayLong(llDate) <> 0) Then  '0=monday
'        pbcCtrl.ForeColor = MAGENTA 'GREEN
'        llStdDate = llDate
'        Do While gWeekDayLong(llStdDate) <> 0   '0=monday
'            llStdDate = llStdDate - 1
'        Loop
'        llStartDate = llStdDate
'        ilStdMonth = Month(llStdDate)
'        Do
'            ilWkDay = gWeekDayLong(llStdDate)
'            If (ilType = 0) Or (ilType = 4) Then
'                slDay = Trim$(Str$(Day(llStdDate)))
'                If Len(slDay) <= 1 Then
'                    slDay = " " & slDay
'                End If
'            ElseIf ilType = 2 Then  'Julian +, left code in if want julian by std
'                slDay = Format$(llStdDate, "y")
'                If Len(slDay) <= 1 Then
'                    slDay = "   " & slDay
'                ElseIf Len(slDay) = 2 Then
'                    If Val(slDay) < 20 Then
'                        slDay = "  " & slDay
'                    Else
'                        slDay = " " & slDay
'                    End If
'                End If
'            ElseIf ilType = 3 Then  'Julian -, left doe in if want julian by std
'                slJulian = "12/31/" & Format$(llStdDate, "yyyy")
'                slDay = Trim$(Str$(Val(Format$(slJulian, "y")) - Val(Format$(llStdDate, "y"))))
'                If Len(slDay) <= 1 Then
'                    slDay = "   " & slDay
'                ElseIf Len(slDay) = 2 Then
'                    If Val(slDay) < 20 Then
'                        slDay = "  " & slDay
'                    Else
'                        slDay = " " & slDay
'                    End If
'                End If
'            End If
''            gSetShow pbcCtrl, slDay, tlCtrls(ilWkDay + 1)
'            pbcCtrl.CurrentX = tlCtrls(ilWkDay + 1).fBoxX + flBoxInsetX
'            pbcCtrl.CurrentY = tlCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tlCtrls(ilWkDay + 1).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
'            pbcCtrl.Print slDay 'tlCtrls(ilWkDay + 1).sShow
'            If ilWkDay = 6 Then
'                ilRowNo = ilRowNo + 1
'            End If
'            llStdDate = llStdDate + 1
'        Loop Until Month(llStdDate) <> ilStdMonth
'        pbcCtrl.ForeColor = llColor
'    End If
'    If (ilType = 0) Then
'        llStdDate = llDate
'        Do
'            If gWeekDayLong(llStdDate) = 6 Then  'Save last sunday
'                llLastDate = llStdDate
'            End If
'            llStdDate = llStdDate + 1
'        Loop Until Month(llStdDate) <> ilMonth
'    End If
    Do
        ilWkDay = gWeekDayLong(llDate)
        If llStartDate = 0 Then
            llStartDate = llDate
        End If
        If (ilType <= 1) Or (ilType = 4) Then
            slDay = Trim$(Str$(Day(llDate)))
            If Len(slDay) <= 1 Then
                slDay = " " & slDay
            End If
            If ilMonth <> Month(llDate) Then
                pbcCtrl.ForeColor = MAGENTA 'GREEN
            Else
                pbcCtrl.ForeColor = llColor
            End If
        ElseIf ilType = 2 Then  'Julian +
            slDay = Format$(llDate, "y")
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        ElseIf ilType = 3 Then  'Julian -
            slJulian = "12/31/" & Format$(llDate, "yyyy")
            slDay = Trim$(Str$(Val(Format$(slJulian, "y")) - Val(Format$(llDate, "y"))))
            If Len(slDay) <= 1 Then
                slDay = "   " & slDay
            ElseIf Len(slDay) = 2 Then
                If Val(slDay) < 20 Then
                    slDay = "  " & slDay
                Else
                    slDay = " " & slDay
                End If
            End If
        End If
'        gSetShow pbcCtrl, slDay, tlCtrls(ilWkDay + 1)
        pbcCtrl.CurrentX = tlCtrls(ilWkDay + 1).fBoxX + flBoxInsetX
        pbcCtrl.CurrentY = tlCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tlCtrls(ilWkDay + 1).fBoxH + 15) + flBoxInsetY '(fgBoxGridH + 15) -  30'+ fgBoxInsetY
        pbcCtrl.Print slDay 'tlCtrls(ilWkDay + 1).sShow
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
'    Loop Until (Month(llDate) <> ilMonth) Or ((ilType <> 1) And (llDate > llLastDate)) 'julian by std
    Loop Until (((ilType <> 0) And (ilType <> 4)) And (Month(llDate) <> ilMonth)) Or (((ilType = 0) Or (ilType = 4)) And (llDate > llLastDate))
    pbcCtrl.ForeColor = llColor
    llEndDate = llDate - 1
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontName = slFontName
    pbcCtrl.FontSize = flFontSize
    pbcCtrl.FontBold = True
End Sub
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
'*******************************************************
'*                                                     *
'*      Procedure Name:gSetCtrl                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set mouse area within the       *
'*                     control area.                   *
'*                                                     *
'*******************************************************
Sub gSetCtrl(tlCtrlArray As FIELDAREA, flBoxX As Single, flBoxY As Single, flBoxW As Single, flBoxH As Single)
'
'   gSetCtrl CtrlArray(1), fBoxX, fBoxY, fBoxW, fBoxH
'   Where
'       CtrlArray (I/O)- field control array
'       fBoxX (I)- x offset of mouse area within picture control
'       fBoxY (I)- y offset of mouse area within picture control
'       fBoxW (I)- Width of mouse area within picture
'       fBoxH (I)- Height of mouse area within picture
'
    tlCtrlArray.fBoxX = flBoxX
    tlCtrlArray.fBoxY = flBoxY
    tlCtrlArray.fBoxW = flBoxW
    tlCtrlArray.fBoxH = flBoxH
    tlCtrlArray.iReq = True
    tlCtrlArray.iChg = False
    tlCtrlArray.iAlign = 0
    tlCtrlArray.sShow = ""
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gTimeToCurrency                 *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert time to currency (for   *
'*                     precision-                      *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gTimeToCurrency(slInpTime As String, ilChk12M As Integer) As Currency
'
'   clRetTime = gTimeToCurrency(slTime, ilChk12M)
'   Where:
'       slTime (I)- Time as string to be converted to currency
'       ilChk12M(I)- True=If 12M (0) convert to 86400 (24*3600)- handle end time
'                    False=Leave 12m as (0)
'       clRetTime (O)- time as currency
'
    Dim slTime As String
    Dim clTime As Currency
    Dim ilPos As Integer
    Dim slAnyTime As String
    slAnyTime = Trim$(slInpTime)
    On Error GoTo gTimeToCurrencyErr
    ilPos = InStr(slTime, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gTimeToCurrency = 0
            Exit Function
        End If
        slTime = Mid$(slAnyTime, 2)
    Else
        slTime = slAnyTime
    End If
    slTime = gConvertTime(slTime)
    clTime = Hour(slTime) * 3600
    clTime = clTime + Minute(slTime) * 60
    clTime = clTime + Second(slTime)
    If (clTime = 0) And ilChk12M Then
        clTime = 86400
    End If
    If ilPos = 0 Then
        gTimeToCurrency = clTime
    Else
        gTimeToCurrency = -clTime
    End If
    Exit Function
gTimeToCurrencyErr:
    On Error GoTo 0
    gTimeToCurrency = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gUnpackLength                   *
'*                                                     *
'*             Created:5/26/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Unpack length of time in btrieve*
'*                     format                          *
'*                                                     *
'*******************************************************
Sub gUnpackLength(ilHsSec As Integer, ilMinHr As Integer, slFormat As String, ilTest120 As Integer, slLength As String)
'
'   gUnpackLength ilHsSec, ilMinHr, slFormat, ilTest120, slTime
'   Where:
'       ilHsSec (I)- High order byte = hundredths of seconds; low order byte = seconds
'       ilMinHr (I)- High order byte = minute; low order byte = hours
'       slFormat (I)-"1" = hh:mm:ss
'                           "2" = hh mm'ss"
'                           "3" = hhHmmMssS
'       ilTest120 (I)- True=Test if length is 120 or less and show as seconds
'       slTime (O) - Length of Time in format
'
    Dim ilSec As Integer    'Seconds
    Dim ilMin As Integer    'Minutes
    Dim ilHour As Integer   'Hours
    Dim slSec As String     'Seconds
    Dim slMin As String     'Minutes
    Dim slHour As String    'Hours
    Dim llTotalSec As Long
    If (ilHsSec = 1) And (ilMinHr = 0) Then
        slLength = ""
        Exit Sub
    End If
    If (ilHsSec = 0) And (ilMinHr = 0) Then
        If slFormat = "1" Then
            slLength = "0"
            Exit Sub
        ElseIf slFormat = "3" Then
            slLength = "0s"
            Exit Sub
        Else
            slLength = "0" & """"
            Exit Sub
        End If
    End If
    ilSec = ilHsSec \ 256    'Obtain seconds
    slSec = Trim$(Str$(ilSec))
    If Len(slSec) = 1 Then
        slSec = "0" & slSec
    End If
    ilMin = ilMinHr And &HFF 'Obtain Minutes
    slMin = Trim$(Str$(ilMin))
    ilHour = ilMinHr \ 256   'Obtain month
    slHour = Trim$(Str$(ilHour))
    If ilTest120 Then
    llTotalSec = 60 * ilHour
        llTotalSec = 60 * llTotalSec + 60 * ilMin + ilSec
        If llTotalSec <= 120 Then
            ilHour = 0
            ilMin = 0
            ilSec = llTotalSec
            slSec = Trim$(Str$(ilSec))
            If Len(slSec) = 1 Then
                slSec = "0" & slSec
            End If
        End If
    End If
    Select Case slFormat
        Case "1"    'hh:mm:ss
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                slLength = slHour & ":" & slMin & ":" & slSec
            ElseIf ilMin <> 0 Then
                slLength = slMin & ":" & slSec
            Else
                slLength = slSec
            End If
        Case "2"    'hh mm'ss"
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                If (ilMin = 0) And (ilSec = 0) Then
                    slLength = slHour
                ElseIf ilSec = 0 Then
                    slLength = slHour & " " & slMin & "'"
                Else
                    slLength = slHour & " " & slMin & "'" & slSec & """"
                End If
            ElseIf ilMin <> 0 Then
                If ilSec = 0 Then
                    slLength = slMin & "'"
                Else
                    slLength = slMin & "'" & slSec & """"
                End If
            Else
                slLength = slSec & """"
            End If
        Case "3"    'hhHmmMssS
            If ilHour <> 0 Then
                If Len(slMin) = 1 Then
                    slMin = "0" & slMin
                End If
                If (ilMin = 0) And (ilSec = 0) Then
                    slLength = slHour & "h"
                ElseIf ilSec = 0 Then
                    slLength = slHour & "h" & slMin & "m"
                Else
                    slLength = slHour & "h" & slMin & "m" & slSec & "s"
                End If
            ElseIf ilMin <> 0 Then
                If ilSec = 0 Then
                    slLength = slMin & "m"
                Else
                    slLength = slMin & "m" & slSec & "s"
                End If
            Else
                slLength = slSec & "s"
            End If
        Case Else
            slLength = ""
    End Select
    slLength = Trim$(slLength)
End Sub

Public Function gFunctionKeyBranch(KeyCode As Integer)
    Exit Function
End Function
'***************************************************************************************
'*
'* Function Name: gLogMsg
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: A general file routine that shows: Date and Time followed by a message
'*
'***************************************************************************************

Public Sub gLogMsg(sMsg As String, sFileName As String, iKill As Integer)

    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'iKill = True then delete the file first, iKill = False then append to the file
    
    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slToFile As String
    
    slToFile = sgDBPath & "Messages\" & sFileName
    On Error GoTo Error

    If iKill = True Then
        ilRet = 0
        slDateTime = FileDateTime(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
    End If
    
    hlLogFile = FreeFile
    Open slToFile For Append As hlLogFile
    slFullMsg = Format$(Now, "mm/dd/yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " " & sMsg
    Print #hlLogFile, slFullMsg
    Close hlLogFile
    Exit Sub
    
Error:
    ilRet = 1
    Resume Next
    
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:gConvertTime                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Format time so it can be used   *
'*                     VB time procedures              *
'*                                                     *
'*******************************************************
Function gConvertTime(slInpTime As String) As String
'
'   sRetTime = gConvertTime(sTime)
'   Where:
'       sTime (I)- Time string to be checked and formatted
'       sRetTime (O)- Formatted time
'
    Dim slFixedTime As String
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim ilPos3 As Integer
    Dim slTime As String

    slTime = Trim$(slInpTime)
    If slTime = "" Then
        slTime = "12:00AM"
    End If
    slFixedTime = UCase$(slTime)
    ilPos1 = InStr(slFixedTime, "N")
    If ilPos1 <> 0 Then
        slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "PM" & Mid$(slFixedTime, ilPos1 + 1)
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    If (ilPos1 = 0) And (ilPos2 = 0) Then
        ilPos1 = InStr(slFixedTime, "M")
        If ilPos1 <> 0 Then
            slFixedTime = Left$(slFixedTime, ilPos1 - 1) & "AM" & Mid$(slFixedTime, ilPos1 + 1)
        End If
    End If
    ilPos1 = InStr(slFixedTime, "A")
    ilPos2 = InStr(slFixedTime, "P")
    ilPos3 = InStr(slFixedTime, "M")
    If (ilPos3 = 0) And ((ilPos1 <> 0) Or (ilPos2 <> 0)) Then
        slFixedTime = slFixedTime & "M"
    End If
    If InStr(slTime, ":") = 0 Then
        If ilPos1 <> 0 Then
            If Len(slFixedTime) <= 4 Then
                slFixedTime = Left$(slFixedTime, ilPos1 - 1) & ":00AM"
            Else
                If Len(slFixedTime) <= 6 Then
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                Else
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & right$(slFixedTime, 6)
                    slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                End If
            End If
        Else
            If ilPos2 <> 0 Then
                If Len(slFixedTime) <= 4 Then
                    slFixedTime = Left$(slFixedTime, ilPos2 - 1) & ":00PM"
                Else
                    If Len(slFixedTime) <= 6 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 6) & ":" & right$(slFixedTime, 6)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                    End If
                End If
            Else
                If Len(slFixedTime) <= 2 Then
                    slFixedTime = slFixedTime & ":00"
                Else
                    If Len(slFixedTime) <= 4 Then
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & right$(slFixedTime, 2)
                    Else
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 4) & ":" & right$(slFixedTime, 4)
                        slFixedTime = Left$(slFixedTime, Len(slFixedTime) - 2) & ":" & right$(slFixedTime, 2)
                    End If
                End If
            End If
        End If
    End If
    gConvertTime = slFixedTime
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gTimeToCurrency                 *
'*                                                     *
'*             Created:8/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Convert time to currency (for   *
'*                     precision-                      *
'*                     Hours*3600+Min*60+Seconds)      *
'*                                                     *
'*******************************************************
Function gTimeToLong(slInpTime As String, ilChk12M As Integer) As Long
'
'   llRetTime = gTimeToLong(slTime, ilChk12M)
'   Where:
'       slTime (I)- Time as string to be converted to currency
'       ilChk12M(I)- True=If 12M (0) convert to 86400 (24*3600)- handle end time
'                    False=Leave 12m as (0)
'       llRetTime (O)- time as Long
'
    Dim slTime As String
    Dim llTime As Long
    Dim ilPos As Integer
    Dim slAnyTime As String

    slAnyTime = Trim$(slInpTime)
    On Error GoTo gTimeToLongErr
    ilPos = InStr(slTime, "-")
    If ilPos <> 0 Then
        If ilPos <> 1 Then
            gTimeToLong = 0
            Exit Function
        End If
        slTime = Mid$(slAnyTime, 2)
    Else
        slTime = slAnyTime
    End If
    slTime = gConvertTime(slTime)
    llTime = CLng(Hour(slTime)) * 3600
    llTime = llTime + Minute(slTime) * 60
    llTime = llTime + Second(slTime)
    If (llTime = 0) And ilChk12M Then
        llTime = 86400
    End If
    If ilPos = 0 Then
        gTimeToLong = CLng(llTime)
    Else
        gTimeToLong = -llTime
    End If
    Exit Function
gTimeToLongErr:
    On Error GoTo 0
    gTimeToLong = 0
    Exit Function
End Function
