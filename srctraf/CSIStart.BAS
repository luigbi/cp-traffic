Attribute VB_Name = "CSIStartSubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: API.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the API declarations
Option Explicit

Public Const MAX_PATH = 260

'  ----- VerFindFile() flags -----
Public Const VFFF_ISSHAREDFILE = &H1

Public Const VFF_CURNEDEST = &H1
Public Const VFF_FILEINUSE = &H2
Public Const VFF_BUFFTOOSMALL = &H4

'  ----- VerInstallFile() flags -----
Public Const VIFF_FORCEINSTALL = &H1
Public Const VIFF_DONTDELETEOLD = &H2

Public Const VIF_TEMPFILE = &H1&
Public Const VIF_MISMATCH = &H2&
Public Const VIF_SRCOLD = &H4&

Public Const VIF_DIFFLANG = &H8&
Public Const VIF_DIFFCODEPG = &H10&
Public Const VIF_DIFFTYPE = &H20&

Public Const VIF_WRITEPROT = &H40&
Public Const VIF_FILEINUSE = &H80&
Public Const VIF_OUTOFSPACE = &H100&
Public Const VIF_ACCESSVIOLATION = &H200&
Public Const VIF_SHARINGVIOLATION = &H400&
Public Const VIF_CANNOTCREATE = &H800&
Public Const VIF_CANNOTDELETE = &H1000&
Public Const VIF_CANNOTRENAME = &H2000&
Public Const VIF_CANNOTDELETECUR = &H4000&
Public Const VIF_OUTOFMEMORY = &H8000&

Public Const VIF_CANNOTREADSRC = &H10000
Public Const VIF_CANNOTREADDST = &H20000
Public Const VIF_BUFFTOOSMALL = &H40000

' Special HWND value for use with PostMessage and SendMessage
Public Const HWND_BROADCAST = -1
Public Const WM_FONTCHANGE = &H1D

Declare Function VerFindFile Lib "Version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Declare Function VerInstallFile Lib "Version.dll" Alias "VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long

'Moved to API.BAS 10/20/03
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


'Declare Function SendMessage& Lib "User" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
'Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
'Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type

Global Const STILL_ACTIVE = &H103
Global Const PROCESS_QUERY_INFORMATION = &H400

Public Declare Function RegComctl32 Lib "Comctl32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegComdlg32 Lib "Comdlg32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegCraxdrt Lib "Craxdrt.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegCrtslv Lib "Crtslv.dll" Alias "DllRegisterServer" () As Long
Public Declare Function Regrviewer Lib "Crviewer.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegDuzactx Lib "Duzactx.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegDzactx Lib "Dzactx.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegExportModeller Lib "ExportModeller.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegMci32 Lib "Mci32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegMfc40 Lib "Mfc40.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegMfc42 Lib "Mfc42.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegMshflxgd Lib "Mshflxgd.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegMsmask32 Lib "Msmask32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegMsComCtl Lib "MsComCtl.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegMsrdo20 Lib "Msrdo20.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegMsvcrt Lib "Msvcrt.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegMsComm32 Lib "MsComm32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegOleaut32 Lib "Oleaut32.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegOlepro32 Lib "Olepro32.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegResize32 Lib "Resize32.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function RegSsdw3b32 Lib "Ssdw3b32.ocx" Alias "DllRegisterServer" () As Long
'Public Declare Function Regtabctl32 Lib "Tabctl32.ocx" Alias "DllRegisterServer" () As Long
Public Declare Function RegCruflcsi Lib "cruflCSI.dll" Alias "DllRegisterServer" () As Long
Public Declare Function RegCrqe Lib "crqe.dll" Alias "DllRegisterServer" () As Long

Public Declare Function Regpdfsplitmerge Lib "pdfsplitmerge.dll" Alias "DllRegisterServer" () As Long
'Dan 8/30/2010 be able to read csiNetReporter fileversion
Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer
   dwStrucVersionh As Integer
   dwFileVersionMSl As Integer
   dwFileVersionMSh As Integer
   dwFileVersionLSl As Integer
   dwFileVersionLSh As Integer
   dwProductVersionMSl As Integer
   dwProductVersionMSh As Integer
   dwProductVersionLSl As Integer
   dwProductVersionLSh As Integer
   dwFileFlagsMask As Long
   dwFileFlags As Long
   dwFileOS As Long
   dwFileType As Long
   dwFileSubtype As Long
   dwFileDateMS As Long
   dwFileDateLS As Long
End Type
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)


Public Function gDriveExists(ByVal sDrive As String) As Boolean
    Dim buffer As String
    buffer = Space(64)
    ' return False if invalid argument
    If Len(sDrive) = 0 Then Exit Function
    'get the string that contains all drives
    GetLogicalDriveStrings Len(buffer), buffer
    ' check that the letter we're looking for is there
    gDriveExists = InStr(1, buffer, left$(sDrive, 1), vbTextCompare)
End Function

Public Function gShellAndWait(Frm As Form, ByVal slFilePath As String, ilWinStyle As Integer, ilWait As Integer) As Long
    Dim llProcess As Long
    Dim llReturn As Long
    Dim llRet As Long
    Dim llIDShell As Long
    
    On Error GoTo gShellAndWaitErr:
    
    llIDShell = Shell(slFilePath, ilWinStyle)
    If llIDShell <> 0 Then
        llProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, llIDShell)
        
        If llProcess <> 0 Then
            If Not ilWait Then
                gShellAndWait = 0
                Exit Function
            End If
            Frm.Enabled = False
            Do
                GetExitCodeProcess llProcess, llReturn
                Sleep 50
                DoEvents
            Loop While llReturn = STILL_ACTIVE
            llRet = CloseHandle(llProcess)
            Frm.Enabled = True
            gShellAndWait = 0
        Else
            'MsgBox "Unable to Shell to " & slFilePath, vbOkOnly, "Shell Error"
            llRet = GetLastError()
            gShellAndWait = 0   'llRet
        End If
    Else
        'MsgBox "Unable to Shell to " & slFilePath, vbOkOnly, "Shell Error"
        llRet = -1
        gShellAndWait = llRet
    End If
    Exit Function
gShellAndWaitErr:
    llProcess = 0
    Resume Next
End Function

Public Function gShellAndWaitAwhile(Frm As Form, ByVal slFilePath As String, slProcessName As String, ilWinStyle As Integer, llMinutesToWait As Integer, blAskUser As Boolean) As Boolean
    'O- process completed successfully?
    Dim llProcess As Long
    Dim llReturn As Long
    Dim slStartTime As Single
    Dim slEndTime As Single
    Dim ilRet As Integer
    Dim blFail As Boolean
    
    slStartTime = Timer
    slEndTime = slStartTime + (llMinutesToWait * 60)
    llProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(slFilePath, ilWinStyle))
    If llProcess <> 0 Then
        Frm.Enabled = False
        Do
            If slEndTime < Timer Then
                If blAskUser Then
                    ilRet = MsgBox("The installation seems to be stalled.  If you want to halt installation, please enter 'yes'. Otherwise, enter 'no'.", vbYesNo, "Cancel install file?")
                    If ilRet = vbYes Then
                        blFail = True
                        Exit Do
                    Else
                       slEndTime = Timer + (llMinutesToWait * 60)
                    End If
                Else
                    blFail = True
                    Exit Do
                End If
            End If
            GetExitCodeProcess llProcess, llReturn
            Sleep 50
            DoEvents
        Loop While llReturn = STILL_ACTIVE
        If blFail Then
            Shell "taskkill.exe /f /t /im " & slProcessName
        End If
        Frm.Enabled = True
        gShellAndWaitAwhile = Not blFail
    Else
        MsgBox "Unable to Shell to " & slFilePath, vbOK, "Shell Error"
    End If
End Function
Public Function gLoadFromIni(Section As String, Key As String, slPath As String, sValue As String) As Boolean
    'generic. from any ini with slpath pointing to ini file.
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    
    sValue = "Not Found"
    gLoadFromIni = False
    If Dir(slPath) > "" Then
        BytesCopied = GetPrivateProfileString(Section, Key, "Not Found", sBuffer, 128, slPath)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                sValue = left(sBuffer, BytesCopied)
                gLoadFromIni = True
            End If
        End If
    End If 'slPath not valid?
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function
Public Function gSetPathEndSlash(ByVal slInPath As String) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If Len(slPath) > 0 Then
        If right$(slPath, 1) <> "\" Then
            slPath = slPath + "\"
        End If
    End If
    gSetPathEndSlash = slPath
End Function
