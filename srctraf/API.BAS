Attribute VB_Name = "API"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of API.BAS on Wed 6/17/09 @ 12:56 PM ***
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Declares (Removed)                                                             *
'*  IsWindowVisible               GetParent                     IsWindowEnabled           *
'*                                                                                        *
'*                                                                                        *
'* Public Constants (Marked)                                                              *
'*  GWL_STYLE                     GWL_EXSTYLE                   WS_EX_TOPMOST             *
'*  WS_EX_TRANSPARENT             WNDTOPMOST                    SW_SHOWNORMAL             *
'*  SW_SHOWMAXIMIZED              SW_SHOWNOACTIVATE             WPF_RESTORETOMAXIMIZED    *
'*                                                                                        *
'*                                                                                        *
'* Public Declares (Marked)                                                               *
'*  GetLastError                  CloseHandle                   SendMessage               *
'*  SetWindowLong                 btrSetContinuousMode          btrExtGetFirst            *
'*  btrExtGetNextExt              btrUnlock                     SendMessageLong           *
'*  IsZoomed                                                                              *
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: API.BAS
'
' Release: 1.1
'
' Description:
'   This file contains the API declarations
Option Explicit

'Alert Menu Item
Public Const MF_BITMAP = &H4&
Public Const MF_BYPOSITION = &H400&
Public Const MF_ENABLED = &H0
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
' Bitmap Header Definition
Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal HWnd As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long

Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'Type TIME_ZONE_INFORMATION
'    Bias As Long
'    StandardName(31) As Integer
'    StandardDate As SYSTEMTIME
'    StandardBias As Long
'    DaylightName(31) As Integer
'    DaylightDate As SYSTEMTIME
'    DaylightBias As Long
'End Type
Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Used to get current setting of color, vertical and horizontal resoluation
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long



' These two functions work together. First call gParseGetStringSize to parse the string
' and find out what string size will be needed to store it.
' Then call gParseGetString to actually get the parsed string.

Declare Function gParseGetStringSize Lib "csi_utils.dll" Alias "csiParseGetStringSize" (slInputStr As Variant, ilItemNo As Variant, slDelimiter As Variant) As Integer

Declare Function gParseGetString Lib "csi_utils.dll" Alias "csiParseGetString" (Output As Variant) As Integer

'Declare Sub SetWindowPos Lib "User32" (ByVal h1%, ByVal h2%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%)
Declare Function GetFocus Lib "user32" () As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'SetHandleCount (default for VB is 15 files opened)
Declare Function SetHandleCount Lib "kernel32" (ByVal wNumber As Long) As Long
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize As Long, ByVal lpFileName$)
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpWriteString$, ByVal lpFileName$)
Declare Function InvalidateRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function GetModuleUsage Lib "kernel32" (ByVal hModule As Integer) As Integer

Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetLastError Lib "kernel32" () As Long 'VBC NR
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long 'VBC NR

'Declare Function SendMessage Lib "User32" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
'Declare Function SendMessage Lib "User32" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) 'VBC NR
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam&)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam$)
Declare Sub HMemCpy Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Declare Function GetFreeSystemResources Lib "User32" (ByVal fuSysResource As Integer) As Integer
Declare Function GetVersion Lib "kernel32" () As Long
'Declare Function GetWinFlags Lib "Kernel32" () As Long

Public Const GWL_STYLE = (-16) 'VBC NR
Public Const GWL_EXSTYLE = (-20) 'VBC NR
Public Const WS_EX_TOPMOST = &H8& 'VBC NR
Public Const WS_EX_TRANSPARENT = &H20& 'VBC NR
'Sample code
'UserMsg.Show vbModeless
'SetWindowLong UserMsg.hWnd, GWL_EXSTYLE, WS_EX_TOPMOST & WS_EX_TRANSPARENT
'UserMsg!lacMsg.Caption = "Message in Window"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long 'VBC NR

''VBAPI
'Declare Function GetOffSetForInt Lib "CSI.DLL" (i As Any, j As Any) As Integer
''Declare Function GetOffSetForStr Lib "CSI.DLL" (i As Any, j As Any) As Integer
'Declare Sub ReplaceInt Lib "CSI.DLL" (i As Any, ByVal j As Integer, ByVal k As Integer)
'Declare Sub ReplaceLong Lib "CSI.DLL" (i As Any, ByVal j As Integer, ByVal k As Long)

'csi_os32.dll
Declare Function csiGetOffset Lib "csi_os32.dll" (ByVal slFileName$, ByVal slFieldName$) As Integer


'csi_math.dll
'Declare Function gCompAbsNumberStr Lib "csi_math32.DLL" (ByVal slCompStr1 As String, ByVal slCompStr2 As String) As Integer
'Declare Function gCompNumberStr Lib "csi_math32.DLL" (ByVal slCompStr1 As String, ByVal slCompStr2 As String) As Integer
'Declare Function gLongToUnsignInt Lib "csi_math32.DLL" (llNumber As Long) As Integer
'Declare Function gUnsignIntToLong Lib "csi_math32.DLL" (ilNumber As Integer) As Long
Type FIXSTR
    sStr As String * 10
End Type
Declare Function PDNToStr Lib "csi_math32.DLL" (slInStr As FIXSTR, ilNoDec As Integer, ilSize As Integer, ByVal slAns As String) As Integer

'Copyright « 1991-1996 Crescent Software, Inc.
'Sort
'Setup fnAV helper function
'Const G_MAX_ARRAYDIMS = 60      'VB limit on array dimensions 1 to 60
Const G_MAX_ARRAYDIMS = 59      'VB limit on array dimensions  0 to 59

Type tagAV                      'Array and Vector in 1 compact unit
    PPSA As Long                'Address of pointer to SAFEARRAY
    NumDims As Long         'Number of dimensions
    SCode As Long               'Error info
    flags As Long               'Reserved
    'Subscripts(1 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
    Subscripts(0 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
End Type

Declare Function fnAV Lib "QPRO32.DLL" (ByRef A() As Any, ParamArray SubscriptsVector()) As tagAV
Public Declare Sub ArraySortTyp Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal ElSize As Long, ByVal MbrOff As Long, ByVal MbrSiz As Long, ByVal CaseSensitive As Long)

'Declare Sub ISortT2 Lib "QPro32.DLL" (xArray As Any, Index%, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%)
'Declare Sub ArraySortTyp Lib "QPro32.DLL" (xArray() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal ElSize%, ByVal MemberOffset%, ByVal MemberSize%, ByVal CaseSenitive%)
'Declare Sub ArraySortStr Lib "QPro32.DLL" (xArray() As Any, FirstE1 As Any, ByVal NumEls%, ByVal Direct%, ByVal CaseSenitive%)

'csi_io.dll
'Use in CBtrvTable call
Global Const TWOHANDLES = 0 'Open file on master and slave datbase path
Global Const ONEHANDLE = 1  'Open file on retrieval database path
Global Const TEMPHANDLE = -1 'Open file on temporary report file database path
'Use in btrGetEqual, btrGetFirst, btrGetNext, btrGetLast, btrGetPrevious
Global Const SETFORREADONLY = 0
Global Const SETFORWRITE = 1

Global Const FRMNOMOVE = 2
Global Const FRMNOSIZE = 1
Global Const WNDTOPMOST = -1 'VBC NR
Global Const WNDNOTOPMOST = -2

' CSI_CNT32.DLL functions.
Declare Function btrSetContinuousMode Lib "CSI_CNT32.dll" Alias "csiSetContinuousMode" (ByVal DBPath$, ByVal FileList$, ByVal Mode%, ByVal ClientID%) As Integer 'VBC NR

Declare Function btrIsANetworkPath Lib "CSI_CNT32.dll" Alias "csiIsANetworkPath" (ByVal DBPath$) As Integer

' New functions to control backup.
Type CSISvr_Rsp_GetLastBackupDate
    sLastBackupDateTime As String * 20
End Type

Type CSISvr_Rsp_Answer
    iAnswer As Integer  ' 0=No, 1=Yes
End Type

Declare Function csiGetLastBackupDate Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByRef SvrRsp As CSISvr_Rsp_GetLastBackupDate) As Integer
Declare Function csiGetLastCopyDate Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByRef SvrRsp As CSISvr_Rsp_GetLastBackupDate) As Integer
Declare Function csiStartBackup Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByVal sINIPathFileName$, ByVal BUType As Integer) As Integer
Declare Function csiIsBackupRunning Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByRef SvrRsp As CSISvr_Rsp_Answer) As Integer
Declare Function csiCheckForFilesStuckInCntMode Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByRef SvrRsp As CSISvr_Rsp_Answer) As Integer
Declare Function csiStartCopyDataToTest Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByVal sINIPathFileName$, ByVal BUType As Integer) As Integer
Declare Function csiClearReportFiles Lib "CSI_CNT32.dll" (ByVal sDBPath$, ByVal sINIPathFileName$, ByVal BUType As Integer) As Integer

Declare Sub csiSetValue Lib "csi_io32.dll" (ByVal ilWaitCount%, ByVal ilNoSecsWait%, ByVal ilWaitBias%, ByVal ilTranLog%)

Declare Function csiHandleValue Lib "csi_io32.dll" (ByVal hlFile%, ByVal ilRetrieve%) As Integer

'5666
Declare Sub btrStopApplReal Lib "csi_io32.dll" Alias "csiStopAppl" ()
'Declare Sub btrStopAppl Lib "csi_io32.dll" Alias "csiStopAppl" ()

Declare Sub btrDestroy Lib "csi_io32.dll" Alias "csiDestroy" (Ohnd As Integer)    '(ByVal Ohnd%)


Declare Sub btrExtClear Lib "csi_io32.dll" Alias "csiExtClear" (ByVal Ohnd%)


Declare Function btrExtAddField Lib "csi_io32.dll" Alias "csiExtAddField" (ByVal Ohnd%, ByVal FieldOffset%, ByVal FieldLength%) As Integer



Declare Function btrExtAddLogicConst Lib "csi_io32.dll" Alias "csiExtAddLogicConst" (ByVal Ohnd%, ByVal DataType%, ByVal FieldOffset%, ByVal FieldLength%, ByVal ComparisonCode%, ByVal AndOrLogic%, ConstField As Any, ByVal ConstSize%) As Integer



Declare Function btrExtGetFirst Lib "csi_io32.dll" Alias "csiExtGetFirst" (ByVal Ohnd%, Record As Any, RecordSize As Integer, RecordPosition As Any) As Integer 'VBC NR


Declare Function btrExtGetNext Lib "csi_io32.dll" Alias "csiExtGetNext" (ByVal Ohnd%, Record As Any, RecordSize As Integer, RecordPosition As Any) As Integer


' Use either EG or UC for header control. EG excludes current record. UC includes current record.
Declare Sub btrExtSetBounds Lib "csi_io32.dll" Alias "csiExtSetBounds" (ByVal Ohnd%, ByVal MaxRetrieved%, ByVal MaxSkipped%, ByVal HeaderControl$, ByVal PackName$, ByVal PackStr$)

'Declare Function btrOpenFiles Lib "csi_io32.dll" Alias "csiOpenFiles" (ByVal Ohnd%) As Integer

Declare Function btrRecordLength Lib "csi_io32.dll" Alias "csiRecordLength" (ByVal Ohnd%) As Integer

Declare Function btrRecords Lib "csi_io32.dll" Alias "csiRecords" (ByVal Ohnd%) As Long

Declare Function btrAbortTrans Lib "csi_io32.dll" Alias "csiAbortTrans" (ByVal Ohnd%) As Integer

Declare Function btrBeginTrans Lib "csi_io32.dll" Alias "csiBeginTrans" (ByVal Ohnd%, ByVal LockBias%) As Integer

Declare Function btrClear Lib "csi_io32.dll" Alias "csiClear" (ByVal Ohnd%) As Integer

Declare Function btrClone Lib "csi_io32.dll" Alias "csiClone" (ByVal Ohnd%, ByVal NewFileName$, ByVal FileFlag%) As Integer

Declare Function btrClose Lib "csi_io32.dll" Alias "csiClose" (ByVal Ohnd%) As Integer

Declare Function btrDelete Lib "csi_io32.dll" Alias "csiDelete" (ByVal Ohnd%) As Integer

Declare Function btrEndTrans Lib "csi_io32.dll" Alias "csiEndTrans" (ByVal Ohnd%) As Integer

'Declare Function btrExtend Lib "csi_io32.dll" Alias "csiExtend" (ByVal Ohnd%, ByVal ExtendFilePath$, ByVal fStoreFlag%) As Integer


'Declare Function btrGetCurrent Lib "csi_io32.dll" Alias "csiGetCurrent" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer) As Integer

'Declare Function btrGetCurrentKey Lib "csi_io32.dll" Alias "csiGetCurrentKey" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal LockBias%) As Integer

'Declare Function btrGetData Lib "csi_io32.dll" Alias "csiGetData" (ByVal Ohnd%, ByVal Operation%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer

Declare Function btrGetDirect Lib "csi_io32.dll" Alias "csiGetDirect" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, RecordPosition As Any, ByVal KeyNumber%, ByVal LockBias%) As Integer

Declare Function btrGetEqual Lib "csi_io32.dll" Alias "csiGetEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer

Declare Function btrGetFirst Lib "csi_io32.dll" Alias "csiGetFirst" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer


Declare Function btrGetGreaterOrEqual Lib "csi_io32.dll" Alias "csiGetGreaterOrEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer

'Declare Function btrGetKey Lib "csi_io32.dll" Alias "csiGetKey" (ByVal Ohnd%, ByVal Operation%, KeyBuffer As Any, ByVal KeyNumber%) As Integer

Declare Function btrGetLast Lib "csi_io32.dll" Alias "csiGetLast" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer


Declare Function btrGetLessOrEqual Lib "csi_io32.dll" Alias "csiGetLessOrEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%) As Integer

Declare Function btrGetNext Lib "csi_io32.dll" Alias "csiGetNext" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer

'Declare Function btrGetNextExt Lib "csi_io32.dll" Alias "csiGetNextExt" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer) As Integer

Declare Function btrExtGetNextExt Lib "csi_io32.dll" Alias "csiExtGetNextExt" (ByVal Ohnd%) As Integer 'VBC NR


'Declare Function btrGetPercentage Lib "csi_io32.dll" (ByVal Ohnd%, ByVal Percentage&, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%) As Integer

Declare Function btrGetPosition Lib "csi_io32.dll" Alias "csiGetPosition" (ByVal Ohnd%, RecordPosition As Any) As Integer

Declare Function btrGetPrevious Lib "csi_io32.dll" Alias "csiGetPrevious" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer

Declare Function btrInsert Lib "csi_io32.dll" Alias "csiInsert" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%) As Integer

Declare Function btrOpen Lib "csi_io32.dll" Alias "csiOpen" (ByVal Ohnd%, ByVal OwnerName$, ByVal fileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer


Declare Function btrStepFirst Lib "csi_io32.dll" Alias "csiStepFirst" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer

Declare Function btrStepLast Lib "csi_io32.dll" Alias "csiStepLast" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer

Declare Function btrStepNext Lib "csi_io32.dll" Alias "csiStepNext" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer

Declare Function btrStepPrevious Lib "csi_io32.dll" Alias "csiStepPrevious" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer

'Declare Function btrShutdownApi Lib "csi_io32.dll" Alias "csiShutdownApi" (ByVal Ohnd%, ByVal Abort%) As Integer

'Declare Function btrStop Lib "csi_io32.dll" Alias "csiStop" (ByVal Ohnd%) As Integer

Declare Function btrUnlock Lib "csi_io32.dll" Alias "csiUnlock" (ByVal Ohnd%, ByVal fLock%) As Integer 'VBC NR


Declare Function btrUpdate Lib "csi_io32.dll" Alias "csiUpdate" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer) As Integer

'Declare Function btrVersion Lib "csi_io32.dll" Alias "csiVersion" (ByVal Ohnd%, VersionData As Any) As Integer

'Declare Function CBtrvMngr Lib "csi_io32.dll" Alias "csiCBtrvMngr" (ByVal OptionString$) As Integer

Declare Function CBtrvMngrInit Lib "csi_io32.dll" Alias "csiCBtrvMngrInit" (ByVal InitType%, ByVal MDBPath$, ByVal sDBPath$, ByVal TDBPath$, ByVal RetrievalDB%, ByVal GDBPath$) As Integer
'Declare Function CBtrvMngrInit Lib "cbtrvbx.dll" (ByVal InitType%) As Integer

'Declare Function btrMngrsInUse Lib "csi_io32.dll" Alias "csiMngrsInUse" (ByVal Ohnd%) As Integer

'Declare Function btrInit Lib "csi_io32.dll" Alias "csiInit" (ByVal Ohnd%, ByVal OptionString$) As Integer


Declare Function CBtrvTable Lib "csi_io32.dll" Alias "csiCBtrvTable" (ByVal RetrievalOnly%) As Integer

Declare Function ddExtRecordCount Lib "csi_io32.dll" Alias "csiddExtRecordCount" (ByVal Ohnd%) As Long

Declare Function btrIndexes Lib "csi_io32.dll" Alias "csiIndexes" (ByVal Ohnd%) As Integer

' 02-17-17 CSharp conversion support to replace LSet is this csiCopyMemory function.
Declare Sub csiCopyMemory Lib "csi_io32_cs.dll" (DataBuffer_1 As Any, DataBuffer_1 As Any, nLen_1 As Integer, nLen_2 As Integer)

'csi_get.dll
Declare Function csiSetStamp Lib "csi_get32.dll" (ByVal slFileName$, ByVal slStamp$) As Integer
Declare Function csiSetAlloc Lib "csi_get32.dll" (ByVal slFileName$, ByVal ilStartIndex As Integer, ByVal ilEndIndex As Integer) As Integer

Declare Function csiSetName Lib "csi_get32.dll" (ByVal slID$, ByVal slStamp$) As Integer
Declare Function csiGetName Lib "csi_get32.dll" (ByVal slID$, ByVal slStamp$) As Integer
Declare Sub csiInitName Lib "csi_get32.dll" ()


' Internet Access Functions
Public Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Public Type POINTAPI
    X       As Long
    Y       As Long
End Type

Public Type WINDOWPLACEMENT
    Length            As Long
    flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As RECT
End Type

Public Const SW_SHOWNORMAL = 1 'VBC NR
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3 'VBC NR
Public Const SW_SHOWNOACTIVATE = 4 'VBC NR
Public Const WPF_RESTORETOMAXIMIZED = &H2 'VBC NR

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'VBC NR

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal HWnd As Long) As Long

Public Declare Function SetFocus Lib "user32" (ByVal HWnd As Long) As Long

Public Declare Function IsZoomed Lib "user32" (ByVal HWnd As Long) As Long 'VBC NR

Public Declare Function GetWindowPlacement Lib "user32" (ByVal HWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetWindowPlacement Lib "user32" (ByVal HWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long


Public Type PeekArrayType
    Ptr         As Long
    Reserved    As Currency
End Type
Public Declare Function PeekArray Lib "kernel32" Alias "RtlMoveMemory" (Arr() As Any, Optional ByVal Length As Long = 4) As PeekArrayType

'Megaphone API
Public Type HttpRequestHeader
    sKey As String
    sValue As String
End Type

Public bgApiDelay As Integer 'Milliseconds between API calls
Public Const DeliveryStatus_NA = 0
Public Const DeliveryStatus_NotPushed = 1
Public Const DeliveryStatus_Pushed = 2
Public Const DeliveryStatus_Partial = 3
Public Const DeliveryStatus_IssueEncountered = 4
Public Const DeliveryStatus_RequiresRepush = 5

