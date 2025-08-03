Attribute VB_Name = "EngrAPI"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrAPI.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the API declarations
Option Explicit

'Used to get current setting of color, vertical and horizontal resoluation
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetFocus Lib "user32" () As Long


Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize As Long, ByVal lpFileName$)


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
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpWriteString$, ByVal lpFileName$)

'Crystal declaration to create ttx
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal fileName As String, ByVal bOverWriteExistingFile As Long) As Long

Global Const TWOHANDLES = 0 'Open file on master and slave datbase path
Global Const ONEHANDLE = 1  'Open file on retrieval database path
Global Const BTRV_LOCK_NONE = 0
Global Const BTRV_OPEN_NORMAL = 1
Global Const BTRV_OPEN_NONSHARE = 0
Global Const BTRV_OPEN_VERIFY = -3
Global Const SETFORREADONLY = 0
Global Const SETFORWRITE = 1
Global Const INDEXKEY0 = 0
Global Const INDEXKEY1 = 1
Global Const INDEXKEY2 = 2
Global Const INDEXKEY3 = 3
Global Const INDEXKEY4 = 4
Global Const INDEXKEY5 = 5

Global Const BTRV_ERR_NONE = 0
Global Const BTRV_ERR_DUPLICATE_KEY = 5
Global Const BTRV_ERR_END_OF_FILE = 9
Global Const BTRV_ERR_CONFLICT = 80
Global Const BTRV_ERR_FILTER_LIMIT = 64
Global Const BTRV_ERR_REJECT_COUNT = 60

Global Const BTRV_KT_INT = 1
Global Const BTRV_EXT_EQUAL = 1
Global Const BTRV_EXT_GT = 2
Global Const BTRV_EXT_LT = 3
Global Const BTRV_EXT_NOT_EQUAL = 4
Global Const BTRV_EXT_GTE = 5
Global Const BTRV_EXT_LTE = 6
Global Const BTRV_EXT_LAST_TERM = 0
Global Const BTRV_EXT_AND = 1
Global Const BTRV_EXT_OR = 2


Declare Function CBtrvMngrInit Lib "csi_io32.dll" Alias "csiCBtrvMngrInit" (ByVal InitType%, ByVal MDBPath$, ByVal sDBPath$, ByVal TDBPath$, ByVal RetrievalDB%, ByVal GDBPath$) As Integer
Declare Function CBtrvTable Lib "csi_io32.dll" Alias "csiCBtrvTable" (ByVal RetrievalOnly%) As Integer
Declare Function btrUpdate Lib "csi_io32.dll" Alias "csiUpdate" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer) As Integer
Declare Function btrGetEqual Lib "csi_io32.dll" Alias "csiGetEqual" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, KeyBuffer As Any, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrInsert Lib "csi_io32.dll" Alias "csiInsert" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%) As Integer
Declare Function btrOpen Lib "csi_io32.dll" Alias "csiOpen" (ByVal Ohnd%, ByVal OwnerName$, ByVal fileName$, ByVal fOpenMode%, ByVal Shareable%, ByVal LockBias%) As Integer
Declare Sub btrStopAppl Lib "csi_io32.dll" Alias "csiStopAppl" ()

Declare Sub btrDestroy Lib "csi_io32.dll" Alias "csiDestroy" (Ohnd As Integer)    '(ByVal Ohnd%)
Declare Function btrRecords Lib "csi_io32.dll" Alias "csiRecords" (ByVal Ohnd%) As Long
Declare Function btrClone Lib "csi_io32.dll" Alias "csiClone" (ByVal Ohnd%, ByVal NewFileName$, ByVal FileFlag%) As Integer
Declare Function btrStepFirst Lib "csi_io32.dll" Alias "csiStepFirst" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer
Declare Function btrStepLast Lib "csi_io32.dll" Alias "csiStepLast" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer
Declare Function btrStepNext Lib "csi_io32.dll" Alias "csiStepNext" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer
Declare Function btrStepPrevious Lib "csi_io32.dll" Alias "csiStepPrevious" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal LockBias%) As Integer
Declare Function btrGetPosition Lib "csi_io32.dll" Alias "csiGetPosition" (ByVal Ohnd%, RecordPosition As Any) As Integer
Declare Function btrIndexes Lib "csi_io32.dll" Alias "csiIndexes" (ByVal Ohnd%) As Integer
Declare Function btrGetFirst Lib "csi_io32.dll" Alias "csiGetFirst" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrGetNext Lib "csi_io32.dll" Alias "csiGetNext" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrGetLast Lib "csi_io32.dll" Alias "csiGetLast" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal KeyNumber%, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer
Declare Function btrGetPrevious Lib "csi_io32.dll" Alias "csiGetPrevious" (ByVal Ohnd%, DataBuffer As Any, DataLength As Integer, ByVal Lock_GetKey%, ByVal ForUpdate%) As Integer

Declare Sub btrExtClear Lib "csi_io32.dll" Alias "csiExtClear" (ByVal Ohnd%)
Declare Function btrExtAddLogicConst Lib "csi_io32.dll" Alias "csiExtAddLogicConst" (ByVal Ohnd%, ByVal DataType%, ByVal FieldOffset%, ByVal FieldLength%, ByVal ComparisonCode%, ByVal AndOrLogic%, ConstField As Any, ByVal ConstSize%) As Integer
Declare Function btrExtAddField Lib "csi_io32.dll" Alias "csiExtAddField" (ByVal Ohnd%, ByVal FieldOffset%, ByVal FieldLength%) As Integer
Declare Sub btrExtSetBounds Lib "csi_io32.dll" Alias "csiExtSetBounds" (ByVal Ohnd%, ByVal MaxRetrieved%, ByVal MaxSkipped%, ByVal HeaderControl$, ByVal PackName$, ByVal PackStr$)
Declare Function btrExtGetNext Lib "csi_io32.dll" Alias "csiExtGetNext" (ByVal Ohnd%, Record As Any, RecordSize As Integer, RecordPosition As Any) As Integer

'csi_os32.dll
Declare Function csiGetOffset Lib "csi_os32.dll" (ByVal slFileName$, ByVal slFieldName$) As Integer


'Setup fnAV helper function
Const G_MAX_ARRAYDIMS = 60      'VB limit on array dimensions

Type tagAV                      'Array and Vector in 1 compact unit
    PPSA As Long                'Address of pointer to SAFEARRAY
    NumDims As Long         'Number of dimensions
    SCode As Long               'Error info
    Flags As Long               'Reserved
    Subscripts(1 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
End Type

Declare Function fnAV Lib "QPRO32.DLL" (ByRef A() As Any, ParamArray SubscriptsVector()) As tagAV
Public Declare Sub ArraySortTyp Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal ElSize As Long, ByVal MbrOff As Long, ByVal MbrSiz As Long, ByVal CaseSensitive As Long)

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

Declare Function btrIsANetworkPath Lib "CSI_CNT32.dll" Alias "csiIsANetworkPath" (ByVal DBPath$) As Integer

'ttp5218 spell check for email
Public Const FRMNOMOVE = 2
Public Const FRMNOSIZE = 1
Public Const WNDNOTOPMOST = -2

Public Const SW_SHOWMINIMIZED = 2
Type RECT
    Left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

Public Type POINTAPI
    x       As Long
    y       As Long
End Type

Public Type WINDOWPLACEMENT
    LENGTH            As Long
    Flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As RECT
End Type

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
    ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function SetWindowPlacement Lib "user32" _
  (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long


