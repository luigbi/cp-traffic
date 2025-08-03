Attribute VB_Name = "modPervasive"
Option Explicit

Type AVAILSS
    iRecType As Integer         '(Bits 15-0 Left to right)Bit 0-3 = record type (avail) = 2-9
                                'Value matches Event type #
    iNoSpotsThis As Integer     'Number of spots associated with this avail;
    iTime(0 To 1) As Integer    'Event Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iLtfCode As Integer            'Used for library buys
    iAvInfo As Integer          'Avail information (bits 15-0 Left to right):
                                '   Bits 0-4 for units;
                                '   Bit 6 for Avail Locked;
                                '   Bit 7 for Spot Locked;
                                '   Bit 8 for sustaining allowed;
                                '   Bit 9 for sponsorship allowed (Not used);
                                '   Bit 10 for Local Only flag (SSLOCALONLY)
                                '   Bit 11 for Feed spot only flag (SSFEEDONLY)
                                '   Bit 12 indicates if across midnight (0=No; 1=Yes) (SSXMID)
    iLen As Integer             'Length in seconds
    iAnfCode As Integer         'Avail name code for booking
    iOrigUnit As Integer        'Original Units, this is only set if avail is overbooked
    iOrigLen As Integer         'Original length, this is only set if avail is overbooked
End Type

Type SSF
    iType                 As Integer         ' 0=Non-Game Image; 1-nn is Game
                                             ' image
    iFillRequired         As Integer         ' Avail needs to be filled (1=Y/0=N).
                                             ' Test for 1=Y
    iCount                As Integer         ' Number of Programs, avails and
                                             ' spots
    iVefCode              As Integer         ' Vehicle code
    iDate(0 To 1)         As Integer         ' Date of summary
    iStartTime(0 To 1)    As Integer         ' Start Time of this record(Byte
                                             ' 0:Hund sec; Byte 1: sec.; Byte 2:
                                             ' min.; Byte 3:hour)
    lCode                 As Long            ' Auto Increment
    'tPAS(1 To 1200) As AVAILSS   '5-10-06 chged from 1000 to 1200 entries, Note field within DDF- this field contains the three subrecords defined above
    tPAS(0 To 1199) As AVAILSS   '5-10-06 chged from 1000 to 1200 entries, Note field within DDF- this field contains the three subrecords defined above
End Type

Type SSFKEY0
    iType As Integer             '0=On Air; 1=Altered (partial day defined)
    iVefCode As Integer             'Vehicle code
    iDate(0 To 1) As Integer        'Date of summary
    iStartTime(0 To 1) As Integer   'Start Time of this record(Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
End Type

Type SSFKEY1
    iVefCode As Integer             'Vehicle code
    iType As Integer                '0=Regular Programming; 1-NN = Sports Programming (Game number)
End Type

Type SSFKEY2
    iVefCode As Integer             'Vehicle code
    iDate(0 To 1) As Integer        'Date of summary
End Type

Type SSFKEY3
    lCode                 As Long
End Type

Type CSPOTSS
    iRecType As Integer         '(Bits 15-0 Left to right):
                                'Bit 0-3 = record type (contract spot) = 10
                                'Bit  Info
                                '  4  Open BB Buy
                                '  5  Close BB Buy
                                '  6  Split Network Primary spot    '  6  Floater BB Buy
                                '  7  Split Network Secondary spot  '  7  Any BB Buy
                                '  8  Donut Buy
                                '  9  Bookend Buy
                                ' 10  Avail Buy
                                ' 11  Pre-emptible
                                ' 12  Exclude Avail Buy
                                ' 13  Library buy
                                ' 14  Exclusions defined
    iRank As Integer            'Bits 0-10 = Number of Quarter Hours that this spots can be booked into
                                '            (End Time - Start Time) * Days\900
                                '            (10a - 6a) * 5 \ 900
                                '            (36000 - 21600) * 5 \ 900
                                '            14400 * 5 \ 900
                                '            80
                                '            Note:  Start and Time time are the true avail range of times
                                '                   Max price value is 1000
                                '                   Special values:
                                '                   1010=Direct Response(R); 1020=Remnant(T); 1030=per Inquiry(Q); 1040=Trade;
                                '                   1050=Promo(M); 1060=PSA(S); 1070=Reservation; 1045=Extra
                                '                   2000 is temporary used in SpotMG to indicate that this is a fill spot
                                'Bits 11-14 = Price Level (0=Fill, 1= N/C; 2-15 obtained from flight price)
                                'Special quarter hours:
    lSdfCode As Long            'Sdf Auto code
    lBkInfo As Long             '(Bits 15-0 Left to right):
                                '  Bits 0-16 for Booked Date (obtained from gDateValue);
                                '  Bits 17-22 for start minutes within hour if avail is in same hour as start time (otherwise 0).  Start Time of line is 6:10am and booked avail is 6:25, therefore 10 will be stored
                                '  Bits 23-28 for end minutes within hour if avail is in the same hour as the end time (otherwise 60)
                                '  Bit  29= Solo Avail
                                '  Bit  30= 1st position
                                '  Bits 0-16 and 17-22 and 23-28 set but not used
    iMnfComp(0 To 1) As Integer 'Competitive code
    iPosLen As Integer          '(Bits 15-0 Left to right):
                                '  Bits 0-11 for Spot length
                                '  Bits 12-15 for position
    iAdfCode As Integer         'Advertiser code number
End Type

Declare Function CBtrvMngrInit Lib "csi_io32.dll" Alias "csiCBtrvMngrInit" (ByVal InitType%, ByVal MDBPath$, ByVal SDBPath$, ByVal TDBPath$, ByVal RetrievalDB%, ByVal GDBPath$) As Integer
Declare Function CBtrvTable Lib "csi_io32.dll" Alias "csiCBtrvTable" (ByVal RetrievalOnly%) As Integer
Declare Function btrOpen Lib "csi_io32.dll" Alias "csiOpen" (ByVal Ohnd%, ByVal OwnerName$, ByVal fileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
Declare Function btrClose Lib "csi_io32.dll" Alias "csiClose" (ByVal Ohnd%) As Integer
Declare Sub btrDestroy Lib "csi_io32.dll" Alias "csiDestroy" (Ohnd As Integer)    '(ByVal Ohnd%)
Declare Function btrGetEqual Lib "csi_io32.dll" Alias "csiGetEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrGetLessOrEqual Lib "csi_io32.dll" Alias "csiGetLessOrEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer
Declare Function btrGetFirst Lib "csi_io32.dll" Alias "csiGetFirst" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrGetNext Lib "csi_io32.dll" Alias "csiGetNext" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrClear Lib "csi_io32.dll" Alias "csiClear" (ByVal Ohnd%) As Integer
Declare Sub btrExtSetBounds Lib "csi_io32.dll" Alias "csiExtSetBounds" (ByVal Ohnd%, ByVal MaxRetrieved%, ByVal MaxSkipped%, ByVal HeaderControl$, ByVal PackName$, ByVal PackStr$)
Declare Function btrExtAddField Lib "csi_io32.dll" Alias "csiExtAddField" (ByVal Ohnd%, ByVal FieldOffset%, ByVal FieldLength%) As Integer
Declare Function btrExtGetNext Lib "csi_io32.dll" Alias "csiExtGetNext" (ByVal Ohnd%, Record As Any, RecordSize As Integer, RecordPosition As Any) As Integer
Declare Sub btrExtClear Lib "csi_io32.dll" Alias "csiExtClear" (ByVal Ohnd%)
'Dan 7/28/09
Declare Function btrGetGreaterOrEqual Lib "csi_io32.dll" Alias "csiGetGreaterOrEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer


Declare Function csiHandleValue Lib "csi_io32.dll" (ByVal hlFile%, ByVal ilRetrieve%) As Integer
'6/29/06: change gGetAstInfo to use API call
Declare Function btrDelete Lib "csi_io32.dll" Alias "csiDelete" (ByVal Ohnd%) As Integer
Declare Function btrInsert Lib "csi_io32.dll" Alias "csiInsert" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%) As Integer
Declare Function btrUpdate Lib "csi_io32.dll" Alias "csiUpdate" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer) As Integer
'6/29/06: End of change

' 02-17-17 CSharp conversion support to replace LSet is this csiCopyMemory function.
Declare Sub csiCopyMemory Lib "csi_io32_cs.dll" (DataBuffer_1 As Any, DataBuffer_1 As Any, nLen_1 As Integer, nLen_2 As Integer)


Declare Sub btrStopAppl Lib "csi_io32.dll" Alias "csiStopAppl" ()

'Use in btrGetEqual, btrGetFirst, btrGetNext, btrGetLast, btrGetPrevious
Global Const SETFORREADONLY = 0
Global Const SETFORWRITE = 1

Global Const TWOHANDLES = 0 'Open file on master and slave datbase path
Global Const ONEHANDLE = 1  'Open file on retrieval database path

Global Const BTRV_LOCK_NONE = 0
Global Const BTRV_OPEN_NORMAL = 1
Global Const BTRV_OPEN_NONSHARE = 0

Global Const INDEXKEY0 = 0
Global Const INDEXKEY1 = 1
Global Const INDEXKEY2 = 2
Global Const INDEXKEY3 = 3

Global Const BTRV_ERR_NONE = 0
Global Const BTRV_ERR_INVALID_OP = 1
Global Const BTRV_ERR_IO_ERR = 2
Global Const BTRV_ERR_NOT_OPEN = 3
Global Const BTRV_ERR_KEY_NOT_FOUND = 4
Global Const BTRV_ERR_DUPLICATE_KEY = 5
Global Const BTRV_ERR_INVALID_KEY = 6
Global Const BTRV_ERR_DIFF_KEY = 7
Global Const BTRV_ERR_INVALID_POS = 8
Global Const BTRV_ERR_END_OF_FILE = 9
Global Const BTRV_ERR_MOD_KEY_VALUE = 10
Global Const BTRV_ERR_INVALID_FNAME = 11
Global Const BTRV_ERR_FILE_NOT_FOUND = 12
Global Const BTRV_ERR_EXT_FILE = 13
Global Const BTRV_ERR_PREIMAGE_OPEN = 14
Global Const BTRV_ERR_PREIMAGE_IO = 15
Global Const BTRV_ERR_EXPANSION = 16
Global Const BTRV_ERR_CLOSE = 17
Global Const BTRV_ERR_DISK_FULL = 18
Global Const BTRV_ERR_UNRECOVERABLE = 19
Global Const BTRV_ERR_MGR_INACTIVE = 20
Global Const BTRV_ERR_KEYBUF_LENGTH = 21
Global Const BTRV_ERR_DATABUF_LENGTH = 22
Global Const BTRV_ERR_POSBLK_LENGTH = 23
Global Const BTRV_ERR_PAGESIZE = 24
Global Const BTRV_ERR_CREATE_IO = 25
Global Const BTRV_ERR_NUMKEYS = 26
Global Const BTRV_ERR_INVALID_KEYPOS = 27
Global Const BTRV_ERR_REC_LENGTH = 28
Global Const BTRV_ERR_KEY_LENGTH = 29
Global Const BTRV_ERR_NOT_BTRV_FILE = 30
Global Const BTRV_ERR_ALREADY_EXTD = 31
Global Const BTRV_ERR_EXTD_IO = 32
Global Const BTRV_ERR_INVALID_EXT_NAME = 34
Global Const BTRV_ERR_DIRECTORY = 35
Global Const BTRV_ERR_TRANSACTION = 36
Global Const BTRV_ERR_TRANS_ACTIVE = 37
Global Const BTRV_ERR_TRANS_FILE_IO = 38
Global Const BTRV_ERR_END_ABORT_TRANS = 39
Global Const BTRV_ERR_TRANS_MAX_FILES = 40
Global Const BTRV_ERR_OP_NOT_ALLOWED = 41
Global Const BTRV_ERR_ACCEL_ACCESS = 42
Global Const BTRV_ERR_INVALID_REC_ADDR = 43
Global Const BTRV_ERR_NULL_KEY_PATH = 44
Global Const BTRV_ERR_INCON_KEY_FLAGS = 45
Global Const BTRV_ERR_ACCESS_DENIED = 46
Global Const BTRV_ERR_MAX_OPEN_FILES = 47
Global Const BTRV_ERR_INVALID_ALT_SEQ = 48
Global Const BTRV_ERR_KEY_TYPE = 49
Global Const BTRV_ERR_OWNER_SET = 50
Global Const BTRV_ERR_INVALID_OWNER = 51
Global Const BTRV_ERR_WRITE_CACHE = 52
Global Const BTRV_ERR_INVALID_INTF = 53
Global Const BTRV_ERR_VARIABLE_PAGE = 54
Global Const BTRV_ERR_INCOMPLT_INDEX = 56
Global Const BTRV_ERR_EXPAND_MEM = 57
Global Const BTRV_ERR_COMPBUF_SIZE = 58
Global Const BTRV_ERR_FILE_EXISTS = 59
Global Const BTRV_ERR_REJECT_COUNT = 60
Global Const BTRV_ERR_WORK_SPACE_SIZE = 61
Global Const BTRV_ERR_INCORRECT_DESCP = 62
Global Const BTRV_ERR_INVALID_EXTDBUF = 63
Global Const BTRV_ERR_FILTER_LIMIT = 64
Global Const BTRV_ERR_INCOR_FLD_OFFSET = 65
Global Const BTRV_ERR_AUTO_TRANS_ABORT = 74
Global Const BTRV_ERR_DEADLOCK_DETECT = 78
Global Const BTRV_ERR_PROGRAMMING = 79
Global Const BTRV_ERR_CONFLICT = 80
Global Const BTRV_ERR_LOCK = 81
Global Const BTRV_ERR_LOST_POS = 82
Global Const BTRV_ERR_READ_TRANS = 83
Global Const BTRV_ERR_REC_LOCKED = 84
Global Const BTRV_ERR_FILE_LOCKED = 85
Global Const BTRV_ERR_FILE_TBL_FULL = 86
Global Const BTRV_ERR_HNDL_TBL_FULL = 87
Global Const BTRV_ERR_INCOM_MODE = 88
Global Const BTRV_ERR_REDIR_DEV_FULL = 90
Global Const BTRV_ERR_SERVER = 91
Global Const BTRV_ERR_TRANS_TBL_FULL = 92
Global Const BTRV_ERR_INCOM_LOCK = 93
Global Const BTRV_ERR_PERMISSION = 94
Global Const BTRV_ERR_SESSION = 95
Global Const BTRV_ERR_COMM_ENV = 96
Global Const BTRV_ERR_DATA_MSGSIZE = 97
Global Const BTRV_ERR_INTERNAL_TRANS = 98

' Data conversion errors.
Global Const BTRV_ERR_CNV_NONE = 5000
Global Const BTRV_ERR_CNV_LENGTH = 5001
Global Const BTRV_ERR_CNV_TRUNC = 5002
Global Const BTRV_ERR_CNV_DATA = 5003
Global Const BTRV_ERR_CNV_OVERFLOW = 5004

' Control errors
Global Const BTRV_ERR_NOMEM = 10000
Global Const BTRV_ERR_TRUNCATE = 10001
Global Const BTRV_ERR_INVALID_CTL = 20100
Global Const VBTRV_ERR_LNK_NOCONTROL = 20500
Global Const VBTRV_ERR_LNK_NOPROP = 20501
Global Const VBTRV_ERR_LNK_UNSUPP_PROP = 20502
Global Const VBTRV_ERR_LNK_NOFIELDS = 20503
Global Const VBTRV_ERR_LNK_BADFIELDS = 20504

Private hmCsfFile As Integer
Private hmCefFile As Integer


'
Type DDFFILE
    iFileID As Integer          'File ID
    sName As String * 20        'Table Name
    sLocation As String * 64    'Table Location
    sFlags As String * 1        'File Flag
    sReserved As String * 10    'Reserved
End Type

Public Const DDFFILEPK As String = "IB30B64BB10"

Public tgDDFFileNames() As DDFFILENAMES     'list of filenames from DDF for conversion of btrieve to odbc drivers for reporting

Type DDFFILENAMES
    sShortName As String * 4                'file name such as vef, shtt, etc.
    sLongName As String * 20                'full file name i.e. vef_vehicles.  when converting from btrieve to odbc drivers,
                                            'the full filename is required in locations field
End Type

Type BASEREC
    'sChar(1 To 200) As Byte 'Record
    sChar(0 To 199) As Byte 'Record
End Type

Dim tmFileDDF As DDFFILE
Dim hmFile As Integer
Dim imFileRecLen As Integer

'12/4/12: File Name to capture activity for
Public sgLogActivityFileName As String
Public sgLogActivityInto As String
'12/4/12: end of change

Public bgIgnoreDuplicateError As Boolean


Public Function mGetCefComment(lCefCode As Long, slComment As String) As Boolean

    Dim ilRet, i, ilLen, ilActualLen As Integer
    Dim ilRecLen As Integer
    Dim tlCEF As CEF
    Dim tlCefSrchKey As LONGKEY0
    'Dim slComment As String
    Dim slTemp As String
    Dim blOneChar As Byte
    
    On Error GoTo ErrHand
    
    mGetCefComment = False
    slComment = ""
    tlCefSrchKey.lCode = lCefCode
    tlCEF.sComment = ""
    ilRecLen = Len(tlCEF) '5011
    ilRet = btrGetEqual(hmCefFile, tlCEF, ilRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet <> BTRV_ERR_NONE Then
        'tlCEF.lCode = 0
        'tlCEF.sComment = ""
        'tlCEF.iStrLen = 0
        'ilRet = btrClose(hmCefFile)
        'btrDestroy hmCefFile
        If lCefCode > 0 Then
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            gMsgBox "btrGetEqual Failed on CEF.BTR " & ilRet
        End If
        Exit Function
    End If

    slComment = gStripChr0(tlCEF.sComment)
    'If tlCEF.iStrLen > 0 Then
    If slComment <> "" Then
        'slComment = Trim$(Left$(tlCEF.sComment, tlCEF.iStrLen))
        ' Strip off any trailing non ascii characters.
        ilLen = Len(slComment)
        ' Find the first valid ascii character from the end and assume the rest of the string is good.
        For i = ilLen To 1 Step -1
            blOneChar = Asc(Mid(slComment, i, 1))
            If blOneChar >= 32 Then
                ' The first valid ASCII character has been found.
                slTemp = Left(slComment, i)
                Exit For
            End If
        Next i
        ilActualLen = i
        ' Scan through and remove any non ASCII characters. This was causing a problem for the web site.
        slComment = ""
        For i = 1 To ilActualLen
            blOneChar = Asc(Mid(slTemp, i, 1))
            If blOneChar >= 32 Then
                slComment = slComment + Mid(slTemp, i, 1)
            Else
                slComment = slComment + " "
            End If
        Next i
    End If

    If slComment <> "" Then
        mGetCefComment = True
    End If
    Exit Function

ErrHand:
    Resume Next
End Function

Public Function mGetCSFComment(lCSFCode As Long, Optional blReplaceCRLF As Boolean = False) As Boolean

    Dim ilRet, i, ilLen, ilActualLen As Integer
    Dim ilRecLen As Integer
    Dim tlCSF As CSF
    Dim tlCsfSrchKey As LONGKEY0
    Dim slComment As String
    Dim slTemp As String
    Dim blOneChar As Byte
    
    On Error GoTo ErrHand
    
    mGetCSFComment = False
    tlCsfSrchKey.lCode = lCSFCode
    tlCSF.sComment = ""
    ilRecLen = Len(tlCSF) '5011
    ilRet = btrGetEqual(hmCsfFile, tlCSF, ilRecLen, tlCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet <> BTRV_ERR_NONE Then
        tlCSF.lCode = 0
        tlCSF.sComment = ""
        'tlCSF.iStrLen = 0
        ilRet = btrClose(hmCsfFile)
        btrDestroy hmCsfFile
        Exit Function
    End If

    If blReplaceCRLF Then
        slComment = Replace(gStripChr0(tlCSF.sComment), sgCRLF, "<BR>")
    Else
        slComment = gStripChr0(tlCSF.sComment)
    End If
    'If tlCSF.iStrLen > 0 Then
    If slComment <> "" Then
        'slComment = Trim$(Left$(tlCSF.sComment, tlCSF.iStrLen))
        ' Strip off any trailing non ascii characters.
        ilLen = Len(slComment)
        ' Find the first valid ascii character from the end and assume the rest of the string is good.
        For i = ilLen To 1 Step -1
            blOneChar = Asc(Mid(slComment, i, 1))
            If blOneChar >= 32 Then
                ' The first valid ASCII character has been found.
                slTemp = Left(slComment, i)
                Exit For
            End If
        Next i
        ilActualLen = i
        ' Scan through and remove any non ASCII characters. This was causing a problem for the web site.
        slComment = ""
        For i = 1 To ilActualLen
            blOneChar = Asc(Mid(slTemp, i, 1))
            If blOneChar >= 32 Then
                slComment = slComment + Mid(slTemp, i, 1)
            Else
                slComment = slComment + " "
            End If
        Next i
    End If
    sgCopyComment = ""
    If slComment <> "" Then
        sgCopyComment = slComment
        gBuildCPYRotCom lCSFCode, slComment
        mGetCSFComment = True
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in modPervasive-mGetCSFComment: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Function


Public Function mOpenPervasiveAPI() As Integer
    
    Dim hgDB As Integer
    
    sgMDBPath = ""
    sgSDBPath = ""
    sgTDBPath = sgDBPath
    igRetrievalDB = 0
    
    'hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    'hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, "", igRetrievalDB, "") 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    Do While csiHandleValue(0, 3) = 0
        '7/6/11
        Sleep 1000
    Loop

    If hgDB <> 0 Then
        gMsgBox "CBtrvMngrInit Failed"
        mOpenPervasiveAPI = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    mOpenPervasiveAPI = True
    
End Function

Public Function mClosePervasiveAPI() As Integer

    btrStopAppl

End Function

Public Function mOpenCEFFile() As Integer

    Dim ilRet As Integer
    
    hmCefFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCefFile, "", sgDBPath & "CEF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on CEF.BTR"
        ilRet = btrClose(hmCefFile)
        btrDestroy hmCefFile
        mOpenCEFFile = False
        Exit Function
    End If
    
    mOpenCEFFile = True

End Function

Public Function mOpenCSFFile() As Integer

    Dim ilRet As Integer
    
    hmCsfFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCsfFile, "", sgDBPath & "CSF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on CSF.BTR"
        ilRet = btrClose(hmCsfFile)
        btrDestroy hmCsfFile
        mOpenCSFFile = False
        Exit Function
    End If
    
    mOpenCSFFile = True

End Function

Public Function mCloseCEFFile() As Integer

    Dim ilRet As Integer
    
    ilRet = btrClose(hmCefFile)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrClose Failed on CEF.BTR"
        btrDestroy hmCefFile
        mCloseCEFFile = False
        Exit Function
    End If
    
    btrDestroy hmCefFile
    mCloseCEFFile = True
    Exit Function

End Function
Public Function gCloseCSFFile()

    Dim ilRet As Integer
    
    ilRet = btrClose(hmCsfFile)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrClose Failed on CSF.BTR"
        btrDestroy hmCsfFile
        gCloseCSFFile = False
        Exit Function
    End If
    
    btrDestroy hmCsfFile
    gCloseCSFFile = True
    Exit Function

End Function

Public Function gOpenMKDFile(hlFile As Integer, slFileName As String) As Integer

    Dim ilRet As Integer
    
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on " & slFileName
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        gOpenMKDFile = False
        Exit Function
    End If
    
    gOpenMKDFile = True

End Function

Public Function gCloseMKDFile(hlFile As Integer, slFileName As String)

    Dim ilRet As Integer
    
    ilRet = btrClose(hlFile)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrClose Failed on " & slFileName
        btrDestroy hlFile
        gCloseMKDFile = False
        Exit Function
    End If
    
    btrDestroy hlFile
    gCloseMKDFile = True
    Exit Function

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gSQLWait                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Execute Insert or Update or     *
'*                     Delete SQL operation            *
'*                                                     *
'*******************************************************
Public Function gSQLWaitNoMsgBox(sSQLQuery As String, iDoTrans As Integer) As Long
    'Dan m 9/19/09 changed llRet and function return to Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand
    
    '12/4/12: Check if activity should be logged
    mLogActivityFileName sSQLQuery
    '12/4/12: end of change
    
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery
        If llRet = 0 Then
            If iDoTrans Then
                cnn.CommitTrans
            End If
        ElseIf (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            fStart = Timer
            Do While Timer <= fStart
                llRet = llRet
            Loop
            iCount = iCount + 1
            If iCount > igWaitCount Then
                'gMsgBox "A SQL error has occurred: " & "Error # " & llRet, vbCritical
                Exit Do
            End If
        End If
    Loop While (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT)
    gSQLWaitNoMsgBox = llRet
    If llRet <> 0 Then
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            On Error GoTo mOpenFileErr:
            'hlMsg = FreeFile
            'Open sgMsgDirectory & "AffErrorLog.txt" For Append As hlMsg
            ilRet = gFileOpen(sgMsgDirectory & "AffErrorLog.txt", "Append", hlMsg)
            If ilRet = 0 Then
                Print #hlMsg, sSQLQuery
                Print #hlMsg, "Error # " & llRet
            End If
            Close #hlMsg
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHand:
'    For Each gErrSQL In cnn.Errors
'        llRet = gErrSQL.NativeError
'        If llRet < 0 Then
'            llRet = llRet + 4999
'        End If
'        'If (llRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        If (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
'            If iDoTrans Then
'                cnn.RollbackTrans
'            End If
'            cnn.Errors.Clear
'            Resume Next
'        End If
'        'If llRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        '    gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
'        'End If
'    Next gErrSQL
'    If llRet = 0 Then
'        llRet = Err.Number
'    End If
'    If iDoTrans Then
'        cnn.RollbackTrans
'    End If
'    'cnn.Errors.Clear
    llRet = mErrHand(iDoTrans)
    Resume Next
mOpenFileErr:
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gSQLWait                        *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Execute Insert or Update or     *
'*                     Delete SQL operation            *
'*                                                     *
'*                     This routine is used by Code    *
'*                     Protect                         *
'*******************************************************
Public Function gSQLWaitNoMsgBoxEX(sSQLQuery As String, iDoTrans As Integer, slModNameLineNo As String) As Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    On Error GoTo ErrHand
    
    '12/4/12: Check if activity should be logged
    mLogActivityFileName sSQLQuery
    '12/4/12: end of change
    
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery
        If llRet = 0 Then
            If iDoTrans Then
                cnn.CommitTrans
            End If
        ElseIf (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            fStart = Timer
            Do While Timer <= fStart
                llRet = llRet
            Loop
            iCount = iCount + 1
            If iCount > igWaitCount Then
                'gMsgBox "A SQL error has occurred: " & "Error # " & llRet, vbCritical
                Exit Do
            End If
        End If
    Loop While (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT)
    gSQLWaitNoMsgBoxEX = llRet
    If llRet <> 0 Then
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            'On Error GoTo mOpenFileErr:
            'hlMsg = FreeFile
            'Open sgMsgDirectory & "AffErrorLog.Txt" For Append As hlMsg
            'Print #hlMsg, sSQLQuery
            'Print #hlMsg, slModNameLineNo & " Error # " & llRet
            'Close #hlMsg
            gLogMsg slModNameLineNo & " Error # " & llRet, "AffErrorLog.Txt", False
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHand:
'    For Each gErrSQL In cnn.Errors
'        llRet = gErrSQL.NativeError
'        If llRet < 0 Then
'            llRet = llRet + 4999
'        End If
'        'If (llRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        If (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
'            If iDoTrans Then
'                cnn.RollbackTrans
'            End If
'            cnn.Errors.Clear
'            Resume Next
'        End If
'        'If llRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
'        '    gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
'        'End If
'    Next gErrSQL
'    If llRet = 0 Then
'        llRet = Err.Number
'    End If
'    If iDoTrans Then
'        cnn.RollbackTrans
'    End If
'    'cnn.Errors.Clear
    llRet = mErrHand(iDoTrans)
    Resume Next
mOpenFileErr:
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCheckDDFDates                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check DDF dates with DDFOddst.csi*
'*                     and DDFPack.csi                 *
'*                                                     *
'*******************************************************
Public Function gCheckDDFDates() As Integer
    Dim hlFrom As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim slDateTime As String
    Dim slDDFFile As String
    Dim slDDFDateTime As String
    Dim ilPos As Integer
    Dim ilEof As Integer
    Dim slDate1 As String
    Dim slDate2 As String
    Dim slTime1 As String
    Dim llTime1S As Long
    Dim llTime1E As Long
    Dim slTime2 As String
    Dim llTime2 As Long
    Dim llLen As Long
    Dim ilLoop As Integer
    Dim slFolder As String
    Dim ilTVIFound As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffset As Integer
    Dim llRecPos As Long
    Dim slTVIDateTime As String
    
    sgDDFDateInfo = ""
    ilRet = 0
    On Error GoTo gCheckDDFDatesErr:
    llLen = FileLen(sgExeDirectory & "csi_io32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_io32.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    ilRet = 0
    slDDFFile = sgDBPath & "Field.DDF"
    slDDFDateTime = gFileDateTime(slDDFFile)
    If ilRet <> 0 Then
        gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please call Counterpoint", vbExclamation, "DDF Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
    slDate1 = Left$(slDDFDateTime, ilPos - 1)
    slFolder = sgDBPath
    ilPos = InStrRev(slFolder, "\", Len(sgDBPath) - 1, vbTextCompare)
    ilRet = 0
    slDDFFile = Left$(slFolder, ilPos) & "NewDDF\Field.DDF"
    slDDFDateTime = gFileDateTime(slDDFFile)
    If ilRet <> 0 Then
        gMsgBox "Unable to find Field.DDF in " & Left$(slFolder, ilPos) & "NewDDF" & ", please call Counterpoint", vbExclamation, "DDF Missing"
        gCheckDDFDates = False
        Exit Function
    End If

    ilTVIFound = False
    slTVIDateTime = ""
    hmFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hmFile, "", sgDBPath & "File.DDF", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imFileRecLen = Len(tmFileDDF) 'btrRecordLength(hlAdf)  'Get and save record length
    ilExtLen = Len(tmFileDDF)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hmFile, tmFileDDF, imFileRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet = BTRV_ERR_NONE Then
            Call btrExtSetBounds(hmFile, llNoRec, -1, "UC", "DDFFILEPK", DDFFILEPK) 'Set extract limits (all records)
            ilOffset = 0
            ilRet = btrExtAddField(hmFile, ilOffset, imFileRecLen)  'Extract iCode field
            If ilRet = BTRV_ERR_NONE Then
                'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                        ilExtLen = Len(tmFileDDF)  'Extract operation record size
                        'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            If StrComp(Left$(tmFileDDF.sName, 3), "TVI", vbTextCompare) = 0 Then
                                slTVIDateTime = Trim$(tmFileDDF.sName)
                                ilTVIFound = True
                                Exit Do
                            End If
                            ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            End If
        End If
    End If
    btrDestroy hmFile
    If ilTVIFound Then
        sgDDFDateInfo = Mid$(slTVIDateTime, 5, 2) & "/" & Mid$(slTVIDateTime, 7, 2) & "/" & Mid$(slTVIDateTime, 9, 2) & " at " & Mid$(slTVIDateTime, 11, 2) & ":" & Mid$(slTVIDateTime, 13, 2)
        ilTVIFound = False
        hmFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hmFile, "", Left$(slFolder, ilPos) & "NewDDF\File.DDF", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        imFileRecLen = Len(tmFileDDF) 'btrRecordLength(hlAdf)  'Get and save record length
        ilExtLen = Len(tmFileDDF)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hmFile   'Clear any previous extend operation
        ilRet = btrGetFirst(hmFile, tmFileDDF, imFileRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet = BTRV_ERR_NONE Then
                Call btrExtSetBounds(hmFile, llNoRec, -1, "UC", "DDFFILEPK", DDFFILEPK) 'Set extract limits (all records)
                ilOffset = 0
                ilRet = btrExtAddField(hmFile, ilOffset, imFileRecLen)  'Extract iCode field
                If ilRet = BTRV_ERR_NONE Then
                    'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
                    ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                            ilExtLen = Len(tmFileDDF)  'Extract operation record size
                            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                            Loop
                            Do While ilRet = BTRV_ERR_NONE
                                If StrComp(Left$(tmFileDDF.sName, 3), "TVI", vbTextCompare) = 0 Then
                                    'Compare Dates and time
                                    If StrComp(slTVIDateTime, Trim$(tmFileDDF.sName), vbTextCompare) <> 0 Then
                                        gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbExclamation, "DDF Problem"
                                        gCheckDDFDates = False
                                        Exit Function
                                    End If
                                    ilTVIFound = True
                                    Exit Do
                                End If
                                ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hmFile, tmFileDDF, ilExtLen, llRecPos)
                                Loop
                            Loop
                        End If
                    End If
                End If
            End If
        End If
        btrDestroy hmFile
    End If
    If Not ilTVIFound Then
        If Trim$(slDDFDateTime) <> "" Then
            ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
            If ilPos > 1 Then
                slDate2 = Left$(slDDFDateTime, ilPos - 1)
                If (DateValue(slDate1) <> DateValue(slDate2)) Then
                    gMsgBox "Call Counterpoint as DDF Dates are in Conflict", vbExclamation, "DDF Problem"
                    gCheckDDFDates = False
                    Exit Function
                End If
            Else
                gMsgBox "Restart Affiliate, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Restart Affiliate, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    
    'Two different csi_io32 routines.  One that uses Classic cbtrv432 and the other that
    'Jeff wrote and does not use cbtrv432.  Jeff does not require the DDFOffst file and
    'is about 900000.  The other is about 90000.
    If llLen > 200000 Then
        '2/17/15: csi_io32 if jeffs, than csi_os32 not required
        'ilRet = 0
        'llLen = FileLen(sgExeDirectory & "csi_os32.dll")
        'If ilRet <> 0 Then
        '    gMsgBox "Unable to find csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        '    gCheckDDFDates = False
        '    Exit Function
        'End If
        'If llLen < 20000 Then
        '    gMsgBox "Incompatible version of csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
        '    gCheckDDFDates = False
        '    Exit Function
        'End If
        gCheckDDFDates = True
        Exit Function
    End If
    ilRet = 0
    llLen = FileLen(sgExeDirectory & "csi_os32.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    If llLen > 20000 Then
        gMsgBox "Incompatible version of csi_os32.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    ilRet = 0
    llLen = FileLen(sgExeDirectory & "cbtrv432.dll")
    If ilRet <> 0 Then
        gMsgBox "Unable to find cbtrv432.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Missing"
        gCheckDDFDates = False
        Exit Function
    End If
    If llLen < 400000 Then
        gMsgBox "Incompatible version of cbtrv.dll in " & sgExeDirectory & ", please call Counterpoint", vbExclamation, "csi_io32 Incompatible"
        gCheckDDFDates = False
        Exit Function
    End If
    If Not ilTVIFound Then
        ilRet = 0
        slDDFFile = sgDBPath & "Field.DDF"
        slDDFDateTime = gFileDateTime(slDDFFile)
        If ilRet <> 0 Then
            gMsgBox "Unable to find Field.DDF in " & sgDBPath & ", please place File in folder and run DDFOffst.exe", vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
        ilPos = InStr(1, slDDFDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate1 = Left$(slDDFDateTime, ilPos - 1)
            slTime1 = Mid$(slDDFDateTime, ilPos + 1)
            llTime1S = gTimeToLong(slTime1, False) - 10800    '3 hours
            llTime1E = gTimeToLong(slTime1, False) + 10800    '3 hours
        Else
            gMsgBox "Restart Affiliate, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate1 = slTVIDateTime
    End If
    'Test Offset table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        'hlFrom = FreeFile
        'Open sgDBPath & "DDFOffst.csi" For Input Access Read Shared As hlFrom
        ilRet = gFileOpen(sgDBPath & "DDFOffst.csi", "Input Access Read Shared", hlFrom)
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFOffst.csi" & " Error " & Str$(ilRet), vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    Do
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        If EOF(hlFrom) Then
            Exit Do
        End If
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                ilPos = InStr(1, slLine, "'DDF Date", vbTextCompare)
                If ilPos = 1 Then
                    slDateTime = Trim$(Mid$(slLine, 11))
                    Exit Do
                End If
            End If
        End If
    Loop Until ilEof
    If slDateTime = "" Then
        Close hlFrom
        gMsgBox "Unable to find DDF Date line in DDFOffst.csi, please run DDFOffst.exe", vbExclamation, "DDFOffst.Csi"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate2 = Left$(slDateTime, ilPos - 1)
            slTime2 = Mid$(slDateTime, ilPos + 1)
            llTime2 = gTimeToLong(slTime2, False)
            If (DateValue(slDate1) <> DateValue(slDate2)) Then
                gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Restart Affiliate, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFOffst.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDFOffst.Csi"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If
    'Test Pack table date stamp
    For ilLoop = 0 To 10 Step 1
        ilRet = 0
        On Error GoTo gCheckDDFDatesErr:
        'hlFrom = FreeFile
        'Open sgDBPath & "DDFPack.csi" For Input Access Read Shared As hlFrom
        ilRet = gFileOpen(sgDBPath & "DDFPack.csi", "Input Access Read Shared", hlFrom)
        If (ilRet <> 0) And (ilLoop = 10) Then
            Close hlFrom
            gMsgBox "Unable to Open " & sgDBPath & "DDFPack.csi" & " Error " & Str$(ilRet), vbExclamation, "DDFOffst.Csi"
            gCheckDDFDates = False
            Exit Function
        ElseIf ilRet = 0 Then
            Exit For
        Else
            Close hlFrom
        End If
    Next ilLoop
    slDateTime = ""
    Do
        ilRet = 0
        'On Error GoTo gCheckDDFDatesErr:
        If EOF(hlFrom) Then
            Exit Do
        End If
        Line Input #hlFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                ilPos = InStr(1, slLine, "'DDF Date", vbTextCompare)
                If ilPos = 1 Then
                    slDateTime = Trim$(Mid$(slLine, 11))
                    Exit Do
                End If
            End If
        End If
    Loop Until ilEof
    If slDateTime = "" Then
        Close hlFrom
        gMsgBox "Unable to find DDF Date line in DDFPack.csi, please run DDFOffst.exe", vbExclamation, "DDF Offset"
        gCheckDDFDates = False
        Exit Function
    End If
    Close hlFrom
    If Not ilTVIFound Then
        ilPos = InStr(1, slDateTime, " ", vbTextCompare)
        If ilPos > 0 Then
            slDate2 = Left$(slDateTime, ilPos - 1)
            slTime2 = Mid$(slDateTime, ilPos + 1)
            llTime2 = gTimeToLong(slTime2, False)
            If (DateValue(slDate1) <> DateValue(slDate2)) Then
                gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
                gCheckDDFDates = False
                Exit Function
            End If
        Else
            gMsgBox "Restart Affiliate, if that does not fix the issue Call Counterpoint as DDF not Found", vbExclamation, "DDF Problem"
            gCheckDDFDates = False
            Exit Function
        End If
    Else
        slDate2 = Trim$(slDateTime)
        If StrComp(slDate1, slDate2, vbTextCompare) <> 0 Then
            gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
            gCheckDDFDates = False
            Exit Function
        End If
    End If
    '''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (gTimeToLong(slTime1, False) <> gTimeToLong(slTime2, False)) Then
    ''If (gDateValue(slDate1) <> gDateValue(slDate2)) Or (llTime2 < llTime1S) Or (llTime2 > llTime1E) Then
    ''Removed time test 9/6/03
    'If (gDateValue(slDate1) <> gDateValue(slDate2)) Then
    '    gMsgBox "Please run DDFOffst.exe as DDFPack.csi (" & slDate2 & ") not generated from latest DDF's (" & slDate1 & ")", vbExclamation, "DDF Pack"
    '    gCheckDDFDates = False
    '    Exit Function
    'End If

    gCheckDDFDates = True
    Exit Function
gCheckDDFDatesErr:
    ilRet = Err.Number
    Resume Next
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mFileDateTime                   *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain file time stamp          *
'*                                                     *
'*******************************************************
Function mFileDateTime(slPathFile As String) As String
    Dim ilRet As Integer

    ilRet = 0
    'On Error GoTo mFileDateTimeErr
    'mFileDateTime = FileDateTime(slPathFile)
    ilRet = gFileExist(slPathFile)
    If ilRet <> 0 Then
        mFileDateTime = Format$(Now, "m/d/yy") & " " & Format$(Now, "h:mm:ssAM/PM")
    Else
        mFileDateTime = FileDateTime(slPathFile)
    End If
    On Error GoTo 0
    Exit Function
'mFileDateTimeErr:
'    ilRet = Err.Number
'    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gExtNoRec                       *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute max # records for       *
'*                     btrieve extend operation        *
'*                                                     *
'*            Formula: # Rec = 60000/(6+RecSize)       *
'*                     6= description size added to    *
'*                        each record extracted and    *
'*                        into the return buffer       *
'*                                                     *
'*******************************************************
Function gExtNoRec(ilRecSize As Integer) As Long
    gExtNoRec = 8000 \ (6 + ilRecSize)  'Change 60000 to 8000
End Function

Public Function mPutCefComment(llCefCode As Long, slComment As String) As Long
    Dim ilRet As Integer
    Dim ilCefRecLen As Integer
    Dim tlCEF As CEF
    Dim tlCefSrchKey As LONGKEY0
    
    If llCefCode > 0 Then
        tlCefSrchKey.lCode = llCefCode
        tlCEF.sComment = ""
        ilCefRecLen = Len(tlCEF) '5011
        ilRet = btrGetEqual(hmCefFile, tlCEF, ilCefRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
        If ilRet <> BTRV_ERR_NONE Then
            tlCEF.lCode = 0
        End If
    Else
        tlCEF.lCode = 0
    End If
    'tlCEF.iStrLen = Len(Trim$(slComment))
    tlCEF.sComment = Trim$(slComment) & Chr$(0) '& Chr$(0) 'sgTB
    ilCefRecLen = Len(tlCEF)    '5 + Len(Trim$(tlCEF.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    'If ilCefRecLen - 2 > 7 Then '-2 so the control character at the end is not counted
    If Trim$(slComment) <> "" Then
        If tlCEF.lCode = 0 Then
            tlCEF.lCode = 0 'Autoincrement
            ilRet = btrInsert(hmCefFile, tlCEF, ilCefRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                gMsgBox "btrInsert Failed on CEF.BTR " & ilRet
            End If
        Else
            ilRet = btrUpdate(hmCefFile, tlCEF, ilCefRecLen)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                gMsgBox "btrUpdate Failed on CEF.BTR " & ilRet
            End If
        End If
    Else
        If tlCEF.lCode <> 0 Then
            ilRet = btrDelete(hmCefFile)
        End If
        tlCEF.lCode = 0
    End If
    mPutCefComment = tlCEF.lCode
End Function
Public Function gInsertAndReturnCode(slSQLQuery As String, slTable As String, slFieldName As String, slValueToReplace As String, Optional blTestRepeatingMax As Boolean = False) As Long
'   Dan M 9/17/09 Perform insert and return new autoincremented code.
'   I: slSqlQuery an insert command (INSERT INTO CEF_Comments_Events (cefCode,cefComments) VALUES (replace,'This is a test') )
'   I: slTable (CEF_Comments_events)
'   I: slFieldName  (cefCode)
'   I: slValueToReplace (replace) the word that will be replaced with the incremented code value
'   O: autoincremented code number--0 means error
    Dim slMaxQuery As String
    Dim llCode As Long
    Dim slNewQuery As String
    Dim ilRet As Integer
    Dim llPrevCode As Long
    On Error GoTo ErrHand
    
    '12/4/12: Check if activity should be logged
    mLogActivityFileName slSQLQuery
    '12/4/12: end of change
    llPrevCode = -1
    slMaxQuery = "SELECT MAX(" & slFieldName & ") from " & slTable
    Do
        '4/22/18
        'Set rst = cnn.Execute(slMaxQuery)
        Set rst = gSQLSelectCall(slMaxQuery)
        'Dan M 9/14/10 take care of Null
        If IsNull(rst(0).Value) Then
            llCode = 1
        Else
            If Not rst.EOF Then
                llCode = rst(0).Value + 1
            Else
                llCode = 1
            End If
        End If
        If (llCode = llPrevCode) And (blTestRepeatingMax) Then
            gInsertAndReturnCode = -1
            Exit Function
        End If
        llPrevCode = llCode
        ilRet = 0
        slNewQuery = Replace(slSQLQuery, slValueToReplace, llCode, , , vbTextCompare)
        '6991
        bgIgnoreDuplicateError = True
        If gSQLWaitNoMsgBox(slNewQuery, False) <> 0 Then
            bgIgnoreDuplicateError = False
            'dan M changed 1/10/13 from goto
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:    'changes ilRet if duplicate value
            '2/8/18:remove setting mouse
            'Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "modPervasive-gInsertAndReturnCode") Then
                gInsertAndReturnCode = -1
                Exit Function
            End If
            ilRet = 1
        End If
        bgIgnoreDuplicateError = False
    Loop While ilRet <> 0
    gInsertAndReturnCode = llCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError "", "modPervasive-gInsertAndReturnCode"
    gInsertAndReturnCode = 0
    Exit Function
ErrHand1:
    Screen.MousePointer = vbDefault
    If gHandleError4994("", "modPervasive-gInsertAndReturnCode") Then
        ilRet = 1
        Return
    End If
    gInsertAndReturnCode = 0
End Function

Public Sub gHandleError(slLogName As String, slMethodName As String)
'General routine to be used in error handler of mehtods with sql calls:
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gHandleError LOGFILE, "Export IDC-mCleanIef"
'    mCleanIef = False
'   always write to affErrorLog.txt.  Unfortunately, gmsgbox does this if igShowMsgBox = 0.
'   write to alternate if slLogName is included and not affErrorLog.txt
    Dim blIsAlternateLog As Boolean
    
'    If Len(slLogName) > 0 Then
'        blIsAlternateLog = True
'    Else
'        blIsAlternateLog = False
'    End If
    'we have an alternate log. always write it out.
    If UCase(slLogName) = "AFFERRORLOG.TXT" Or Len(slLogName) = 0 Then
        blIsAlternateLog = False
    Else
        blIsAlternateLog = True
    End If
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            sgTmfStatus = "E"
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        ElseIf gErrSQL.Number <> 0 Then
            sgTmfStatus = "E"
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        sgTmfStatus = "E"
        gMsg = "A general error has occured in " & slMethodName & ": "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
        If blIsAlternateLog Then
            gLogMsg "ERROR: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, slLogName, False
        End If
        If igShowMsgBox <> 0 Then
             gLogMsg "ERROR: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "affErrorLog.txt", False
        End If

    End If

End Sub

Public Function gHandleError4994(slLogName As String, slMethodName As String) As Boolean
    'DUPLICATE KEYS?  Return true  ttp 5217
    ' if ghandleerror4994("","frmLogin-Load") then
        ' ilret = 1
        ' return
    'End If
    'see rules in function above for log ouptut
    Dim blIsAlternateLog As Boolean
    
'    If Len(slLogName) > 0 Then
'        blIsAlternateLog = True
'    Else
'        blIsAlternateLog = False
'    End If
    If UCase(slLogName) = "AFFERRORLOG.TXT" Or Len(slLogName) = 0 Then
        blIsAlternateLog = False
    Else
        blIsAlternateLog = True
    End If
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then
            'dan 1/10/13 Dick had me add nativeerror = 5
            If gErrSQL.NativeError = -4994 Or gErrSQL.NativeError = 5 Then
                gHandleError4994 = True
                Exit Function
            End If
            gMsg = "A SQL error has occurred in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            If blIsAlternateLog Then
                gLogMsg "Error: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        ElseIf gErrSQL.Number <> 0 Then
            gMsg = "A SQL error has occured in " & slMethodName & ": "
            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, vbCritical
            If blIsAlternateLog Then
                gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, slLogName, False
            End If
            If igShowMsgBox <> 0 Then
                 gLogMsg "ERROR: " & gMsg & gErrSQL.Description & "; Error #" & gErrSQL.Number & "; Line #" & Erl, "affErrorLog.txt", False
            End If
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in " & slMethodName & ": "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        If blIsAlternateLog Then
            gLogMsg "ERROR: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, slLogName, False
        End If
        If igShowMsgBox <> 0 Then
             gLogMsg "ERROR: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "affErrorLog.txt", False
        End If
    End If
    gHandleError4994 = False
End Function

Private Sub mLogActivityFileName(slSQLQuery As String)
    '12/4/12: Routine added to capture activity for a file
    Dim slModuleName As String
    Dim blShowMessage As Boolean
    Dim slStr As String
    Dim slMsg As String
    
    On Error Resume Next
    If sgSQLTrace = "Y" Then
        gLogMsgWODT "W", hgSQLTrace, slSQLQuery
    End If
    If (sgLogActivityFileName = "") Or (sgLogActivityInto = "") Then
        Exit Sub
    End If
    blShowMessage = False
    slStr = UCase$(slSQLQuery)
    If InStr(1, slStr, "INSERT INTO " & sgLogActivityFileName, vbTextCompare) = 1 Then
        blShowMessage = True
    ElseIf InStr(1, slStr, "UPDATE " & sgLogActivityFileName, vbTextCompare) = 1 Then
        blShowMessage = True
    ElseIf InStr(1, slStr, "DELETE FROM " & sgLogActivityFileName, vbTextCompare) = 1 Then
        blShowMessage = True
    End If
    If Not blShowMessage Then
        Exit Sub
    End If
    slModuleName = sgCallStack(2)
    If slModuleName = "" Then
        slModuleName = "Unknown"
    End If
    slMsg = "Call by " & slModuleName & " Activity: " & slSQLQuery
    gLogMsg slMsg, sgLogActivityInto, False
End Sub

Public Function gSQLSelectCall(slSQLQuery As String, Optional slMsg As String = "") As ADODB.Recordset
    On Error GoTo ErrHand
    If sgSQLTrace = "Y" Then
        lgSTimeSQL = timeGetTime
    End If
    Set gSQLSelectCall = cnn.Execute(slSQLQuery)
    If sgSQLTrace = "Y" Then
        lgETimeSQL = timeGetTime
        lgTtlTimeSQL = lgTtlTimeSQL + (lgETimeSQL - lgSTimeSQL)
        gLogMsgWODT "W", hgSQLTrace, slSQLQuery
    End If
    Exit Function
ErrHand:
    If sgSQLTrace = "Y" Then
        gLogMsgWODT "W", hgSQLTrace, slSQLQuery
    End If
    If slMsg <> "" Then
        gHandleError "AffErrorLog.txt", slMsg & " " & slSQLQuery
    Else
        gHandleError "AffErrorLog.txt", slSQLQuery
    End If
End Function

Private Function mErrHand(iDoTrans As Integer) As Long
    Dim llRet As Long
    
    llRet = 0
    
    For Each gErrSQL In cnn.Errors
        llRet = gErrSQL.NativeError
        If llRet < 0 Then
            llRet = llRet + 4999
        End If
        'If (llRet = 84) And (iDoTrans) Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        If (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            If iDoTrans Then
                cnn.RollbackTrans
            End If
            cnn.Errors.Clear
            mErrHand = llRet
            Exit Function
        End If
        'If llRet <> 0 Then              'SQLSetConnectAttr vs. SQLSetOpenConnection
        '    gMsgBox "A SQL error has occurred: " & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        'End If
    Next gErrSQL
    If llRet = 0 Then
        llRet = Err.Number
    End If
    If iDoTrans Then
        cnn.RollbackTrans
    End If
    mErrHand = llRet
End Function
