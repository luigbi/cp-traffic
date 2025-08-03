VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "csiInstallerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'******************csiInstallerInfo***************************
' look into csiComponentCheckers('system' is one in setup; 'user' is the one in C:\csi), check msi's exist and versions match system componentCheckers. Grab
'csiNetReporter's version from inside msi..but don't bother with one in user's system32.  CsiNetReporterAlternate also.

' On error, look at FailureMessage(string) or FailureReason(integer) for what happened.
' To use:
' use Start(slPathToCsiSystem32Folder) to set folder. Most other functions won't work without it.  Can just set PathToCsiSys32, but wanted to make it clear this needed to be done first.
' CompareChecker   sets ComponentChecker values and compares them. if returns false, check failureReason:
    'failureReasons = NOPATHTOCSISYS32, COMPSYSNOEXIST, WRONGAMOUNTMSI,NOERROR (which means need to update)
'   Note:  if system component checkers not set, it will attempt to set it(as well as user's component checker).
    'failureReasons = NOPATHTOCSISYS32, NONETREPORTERVERSION, NOMSI,--(setting system component errors: COMPSYSNOEXIST, WRONGAMOUNTMSI), NOERROR (which means versions don't match)
'  to test individual msi: CompareCheckersMsi(slMsiName As String)  returns 3 way truth value:
    'note: attempts to set componentChecker values if needed, so failureReasons as above.
    '1 = match 0 = didn't match -1 couldn't find that name in systemComponentChecker..error
'  Properties SystemCompVersion and userCompVersion-- pass name of msi. return -1 if value in parameter not in dictionary.
   Dim imMsiClientVersion As Integer
   Dim imMsiCrystalVersion As Integer
   Dim smFailureMessage As String
   Dim imFailure As csiInstallFailureValue
   Dim smCsiSys32Path As String
   Dim dmUserMsis As Dictionary
   Dim dmSystemMsis As Dictionary
   Dim imSystemVersion As Integer
   Const MSICLIENT As String = "csiClientInstall"
   Const MSICRYSTAL As String = "csiCrystalInstall"
   Public Enum csiInstallFailureValue
       ' NOMSI = 5
        COMPCANTREAD = 1
        COMPSYSNOEXIST = 2
        WRONGAMOUNTMSI = 3
        'NOFOLDER = 4
        NOERROR = 0
        NOPATHTOCSISYS32 = 6
        'NONETREPORTERVERSION = 7
        NOSYSTEMVERSION = 8
        COMPUSERNOEXIST = 9
    End Enum
    Public Property Get SystemVersion() As Integer
        SystemVersion = imSystemVersion
    End Property
    Public Property Let SystemVersion(ilValue As Integer)
        If ilValue > 56 Then
            imSystemVersion = ilValue
        Else
            imSystemVersion = 0
        End If
    End Property
    Public Property Get FailureReason() As csiInstallFailureValue
         FailureReason = imFailure
    End Property
   Public Property Get FailureMessage() As String
        FailureMessage = smFailureMessage
   End Property
   Public Property Get PathToCsiSys32() As String
         PathToCsiSys32 = smCsiSys32Path
    End Property
    Public Property Let PathToCsiSys32(slValue As String)
    'dan can this be false on setting drive path?
        If mDirExists(slValue) Then
            smCsiSys32Path = gSetPathEndSlash(slValue)
        Else
            smCsiSys32Path = ""
        End If
   End Property
   Public Property Get SystemCompVersion(slName As String) As Integer
        SystemCompVersion = mGrabSystemMsiVersion(slName)
   End Property
   Public Property Get UserCompVersion(slName As String) As Integer
        If dmUserMsis.Exists(slName) Then
            UserCompVersion = dmUserMsis.Item(slName)
        Else
            UserCompVersion = -1
        End If
   End Property
    Private Sub Class_Initialize()
        Set dmSystemMsis = New Dictionary
        dmSystemMsis.CompareMode = TextCompare
        Set dmUserMsis = New Dictionary
        dmUserMsis.CompareMode = TextCompare
    End Sub
    Private Sub Class_Terminate()
        Set dmSystemMsis = Nothing
        Set dmUserMsis = Nothing
    End Sub
'public methods
    Public Function Start(slPathToCsiSystemFolder32 As String, ilVersion As Integer) As Boolean
        Dim blRet As Boolean
        
        If mDirExists(slPathToCsiSystemFolder32) Then
            smCsiSys32Path = gSetPathEndSlash(slPathToCsiSystemFolder32)
            blRet = True
        Else
            smCsiSys32Path = ""
            blRet = False
        End If
        Me.SystemVersion = ilVersion
        If blRet = True And imSystemVersion > 56 Then
            Start = True
        Else
            Start = False
        End If
    End Function
    Public Function CompareCheckers() As Boolean
        'does all--grabbing compchecker FALSE = NOPATHTOCSISYS32, COMPSYSNOEXIST, NOERROR (which means need to update)
        Dim blRet As Boolean
        mClearFailures
        If imSystemVersion > 56 Then
            If mTestSys32() Then
            'has user run GetSystemChecker? If not, run.
                If Not mSetCompsAsNeeded() Then
                    CompareCheckers = False
                    Exit Function
                End If
                If mCompareCheckers() Then
                    blRet = True
                Else
                    blRet = False
                End If
            End If
        Else
            smFailureMessage = "Must set a valid system version number"
            imFailure = NOSYSTEMVERSION
            blRet = False
        End If
        CompareCheckers = blRet
    End Function
    Public Function CompareCheckersMsi(slMsiName As String) As Integer
            '1 = match 0 = didn't match -1 couldn't find that name in systemComponentChecker..error
        Dim blRet As Boolean
        
        mClearFailures
        If mTestSys32() Then
        'has user run GetSystemChecker? If not, run.  User is run after...assumes if system not set, user isn't either.
            If Not mSetCompsAsNeeded() Then
                CompareCheckersMsi = -1
                Exit Function
            End If
            If dmSystemMsis.Exists(slMsiName) Then
                blRet = mCompareCheckersEach(slMsiName)
                If blRet Then
                    CompareCheckersMsi = 1
                Else
                    CompareCheckersMsi = 0
                End If
            Else
                CompareCheckersMsi = -1
            End If
        Else
            CompareCheckersMsi = -1
        End If
    End Function

'private methods--component stuff
    Private Function GetUserChecker() As Boolean
    'returns false for errors reading text; if no text returns true.  So use passed argument to determine if text exists.
        Dim slFullPath As String
        Dim blReturn As Boolean
        Dim slLine As String
        Dim ilVersion As Integer
        Dim oMyFileObj As FileSystemObject
        Dim MyFile As TextStream
        Dim slName As String
 
        dmUserMsis.RemoveAll
        Set oMyFileObj = New FileSystemObject
        blReturn = True
        '5676
        'slFullPath = "C:\csi\csiComponentCheckerv" & imSystemVersion & ".txt"
        slFullPath = sgRootDrive & "csi\csiComponentCheckerv" & imSystemVersion & ".txt"
        If oMyFileObj.FileExists(slFullPath) Then
            Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForReading, False)
            slLine = MyFile.ReadLine
            Do While Not MyFile.AtEndOfStream
                If Not mCompAddMsi(slLine, dmUserMsis) Then
                'false means error
                    blReturn = False
                    GoTo CLEANUP
                End If
                slLine = MyFile.ReadLine
            Loop
           MyFile.Close
           Set MyFile = Nothing
        Else
            blReturn = False
            imFailure = COMPUSERNOEXIST
            smFailureMessage = slFullPath & " does not exist."
        End If
CLEANUP:
        Set oMyFileObj = Nothing
        If Not MyFile Is Nothing Then
        On Error Resume Next
            MyFile.Close
        On Error GoTo 0
            Set MyFile = Nothing
        End If
        GetUserChecker = blReturn
    End Function
    Private Function GetSystemChecker() As Boolean
        Dim slFullPath As String
        Dim blReturn As Boolean
        Dim slLine As String
        Dim oMyFileObj As FileSystemObject
        Dim MyFile As TextStream
 
        dmSystemMsis.RemoveAll
        Set oMyFileObj = New FileSystemObject
        blReturn = True
        slFullPath = smCsiSys32Path & "csiComponentCheckerv" & imSystemVersion & ".txt"
        If oMyFileObj.FileExists(slFullPath) Then
            Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForReading, False)
            slLine = MyFile.ReadLine
            Do While Not MyFile.AtEndOfStream
                If Not mCompAddMsi(slLine, dmSystemMsis) Then
                    blReturn = False
                    GoTo CLEANUP
                End If
                slLine = MyFile.ReadLine
            Loop
           MyFile.Close
           Set MyFile = Nothing
        Else
           blReturn = False
           smFailureMessage = slFullPath & " does not exist."
           imFailure = csiInstallFailureValue.COMPSYSNOEXIST
        End If
        If blReturn Then
'            'has the 2 msi's necessary?
            blReturn = mSysCompNeed()
        End If
CLEANUP:
        Set oMyFileObj = Nothing
        If Not MyFile Is Nothing Then
        On Error Resume Next
            MyFile.Close
        On Error GoTo 0
            Set MyFile = Nothing
        End If
        GetSystemChecker = blReturn
    End Function

    Private Function mSetCompsAsNeeded() As Boolean
        Dim blRet As Boolean
        
        If mTestSysCompCount() Then
            blRet = True
        ElseIf GetSystemChecker() Then
            If GetUserChecker() Then
                blRet = True
            ElseIf imFailure = COMPUSERNOEXIST Then
                blRet = mCreateUsersChecker()
            End If
        End If
        mSetCompsAsNeeded = blRet
    End Function
    Private Function mCreateUsersChecker() As Boolean
        Dim oMyFileObj As FileSystemObject
        Dim blRet As Boolean
        Dim slPath As String
        ' only here if version file doesn't exist. copy old checker if it exists and reload the dictionary.
        slPath = sgRootDrive & "csi\csiComponentChecker.txt"
        Set oMyFileObj = New FileSystemObject
        If oMyFileObj.FileExists(slPath) Then
            oMyFileObj.CopyFile slPath, sgRootDrive & "csi\csiComponentCheckerv" & imSystemVersion & ".txt", False
            mClearFailures
            blRet = GetUserChecker()
        End If

CLEANUP:
    Set oMyFileObj = Nothing
    mCreateUsersChecker = blRet
    End Function
    Private Function mCompAddMsi(slLine As String, dlMyDictionary As Dictionary) As Boolean
    Dim blReturn As Boolean
    Dim slName As String
    Dim ilVersion As Integer
    
    blReturn = True
    If InStr(1, slLine, "MSI", vbTextCompare) > 0 Then
        If mStripNameAndVersion(slLine, slName, ilVersion) Then
            If ilVersion = 0 Then
               blReturn = False
               smFailureMessage = "CsiComponentChecker.txt's version for " & slName & " is returning 0."
               imFailure = csiInstallFailureValue.COMPCANTREAD
            Else
               dlMyDictionary.Add slName, ilVersion
            End If
        Else
            smFailureMessage = "'CsiComponentChecker.txt' is not formatted properly.  Missing ':'. "
            imFailure = csiInstallFailureValue.COMPCANTREAD
        End If
    End If
    mCompAddMsi = blReturn
    End Function
    Private Function mCompareCheckers() As Boolean
        Dim blReturn As Boolean
        Dim vlKey As Variant
        Dim ilSysVersion As Integer
        Dim ilUserVersion As Integer
        
        If dmSystemMsis.Count = dmUserMsis.Count Then
            For Each vlKey In dmSystemMsis
                If Not mCompareCheckersEach(CStr(vlKey)) Then
                    mCompareCheckers = False
                    Exit Function
                End If
            Next
        End If
        mCompareCheckers = True
    End Function
    Private Function mGrabSystemMsiVersion(slMsiName As String) As Integer
        If dmSystemMsis.Exists(slMsiName) Then
            mGrabSystemMsiVersion = dmSystemMsis.Item(slMsiName)
        Else
            mGrabSystemMsiVersion = -1
        End If
    End Function
    Private Function mCompareCheckersEach(slMsiName As String) As Boolean
        Dim ilSysVersion As Integer
        Dim ilUserVersion As Integer

        If dmUserMsis.Exists(slMsiName) Then
            ilSysVersion = dmSystemMsis.Item(slMsiName)
            ilUserVersion = dmUserMsis.Item(slMsiName)
            'Dan 8/24/11 if 'user' is 12.4 and 'system' is 12.3, they are updated even though the msi in system is old.  Don't alert.
            'If ilSysVersion <> ilUserVersion Then
            If ilSysVersion > ilUserVersion Then
                mCompareCheckersEach = False
                Exit Function
            End If
        Else
            mCompareCheckersEach = False
            Exit Function
        End If
        mCompareCheckersEach = True
    End Function
    Private Function mSysCompNeed() As Boolean
        Dim blReturn As Boolean
        
        blReturn = True
        'didn't find minimum # of msis!
        If dmSystemMsis.Count < 2 Then
            smFailureMessage = "Couldn't find at least 2 msis in 'CsiComponentChecker.txt'"
            imFailure = csiInstallFailureValue.WRONGAMOUNTMSI
            blReturn = False
            'didnt find csiCrystal or csiClient!
            ' 7/19/2012 Dan M changed line below.  By mistake, testing msiClient twice and never msiCrystal
        ElseIf Not (dmSystemMsis.Exists(MSICLIENT) And dmSystemMsis.Exists(MSICRYSTAL)) Then
            smFailureMessage = "Couldn't find needed msis in 'CsiComponentChecker.txt'"
            imFailure = csiInstallFailureValue.WRONGAMOUNTMSI
            blReturn = False
        End If
        mSysCompNeed = blReturn
    End Function
    Private Function mTestSysCompCount() As Boolean
        If dmSystemMsis.Count > 0 Then
            mTestSysCompCount = True
        Else
            mTestSysCompCount = False
        End If
    End Function
'end private method, component stuff
'private generic methods
    Private Function mTestSys32() As Boolean
        If Len(smCsiSys32Path) = 0 Then
            imFailure = csiInstallFailureValue.NOPATHTOCSISYS32
            smFailureMessage = "Set PathToCsiSys32 first."
            mTestSys32 = False
        Else
            mTestSys32 = True
        End If
    End Function

    Private Function mStripNameAndVersion(ByVal slNameAndVersion As String, ByRef slName As String, ByRef ilVersion As Integer) As Boolean
        Dim ilColon As Integer
        Dim slVersion As String
        Dim blRet As Boolean
        'MSICSICLIENTINSTALL:V4.1 becomes CSICLIENTINSTALL & 41
        blRet = True
        ilColon = InStr(1, slNameAndVersion, ":")
        If ilColon > 0 Then
            slName = Trim(Mid(slNameAndVersion, 4, ilColon - 4))
            slVersion = Trim(Mid(slNameAndVersion, ilColon + 1))
            ilVersion = ConvertVersion(slVersion)
        Else
            smFailureMessage = " a csiComponentChecker's " & slNameAndVersion & " is not formatted properly.  Missing ':'. "
            imFailure = csiInstallFailureValue.COMPCANTREAD
            blRet = False
        End If
        mStripNameAndVersion = blRet
    End Function
    Private Sub mClearFailures()
            imFailure = csiInstallFailureValue.NOERROR
            smFailureMessage = ""
    End Sub

  Private Function ConvertVersion(ByVal slVersion As String) As Integer
      Dim ilFirstDot As Integer
      Dim ilSecondDot As Integer

      slVersion = Replace(slVersion, "V", "", , , vbTextCompare)
      ilFirstDot = InStr(1, slVersion, ".")
      If ilFirstDot > 0 Then
         If ilFirstDot <> InStrRev(slVersion, ".") Then
            'crystal version or net platform. trim to '12.2'
            ilSecondDot = InStr(ilFirstDot + 1, slVersion, ".")
            slVersion = Mid(slVersion, 1, ilSecondDot - 1)
         End If
         slVersion = Replace(slVersion, ".", "")
      End If
      If IsNumeric(slVersion) Then
         ConvertVersion = CInt(slVersion)
      Else
         ConvertVersion = 0
      End If
   End Function
Private Function mDirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    mDirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
Private Function mFileExists(file As String) As Boolean
    On Error GoTo Error
    mFileExists = (GetAttr(file) And vbDirectory) = 0
Error:
    'Return False if an error occurs
End Function

Private Function gSetPathEndSlash(ByVal slInPath As String) As String
    Dim slPath As String
    slPath = Trim$(slInPath)
    If Len(slPath) > 0 Then
        If right$(slPath, 1) <> "\" Then
            slPath = slPath + "\"
        End If
    End If
    gSetPathEndSlash = slPath
End Function

'end private generic functions



