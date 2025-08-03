VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4425
   ControlBox      =   0   'False
   Icon            =   "AffLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   20
      TabIndex        =   0
      Top             =   -45
      Width           =   4380
      Begin VB.Timer tmcTerminate 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4020
         Top             =   330
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Exit Affiliate System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   750
         TabIndex        =   4
         Tag             =   "Cancel"
         Top             =   2475
         Width           =   2805
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Start Affiliate System"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   750
         TabIndex        =   3
         Tag             =   "OK"
         Top             =   1995
         Width           =   2805
      End
      Begin VB.TextBox txtPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1605
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1410
         Width           =   2325
      End
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   1605
         TabIndex        =   1
         Top             =   1020
         Width           =   2325
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   120
         Picture         =   "AffLogin.frx":08CA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4035
      End
      Begin VB.Label Label2 
         Caption         =   "®"
         Height          =   180
         Left            =   3660
         TabIndex        =   8
         Top             =   465
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "Counterpoint Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   870
         TabIndex        =   7
         Top             =   345
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   285
         TabIndex        =   6
         Tag             =   "&Password:"
         Top             =   1425
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Tag             =   "&User Name:"
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "AffLogin.frx":4268E4
         Top             =   270
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '******************************************************
'*  frmLogin - basic log-on form for SSQL server
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Dim smSpecial As String
'5/11/11
Dim imGuideUstCode As Integer
'dan ttp 4905
Dim bmNoPervasive As Boolean
'ttp 5217
Private Const LOGFILE As String = "AffErrorLog.Txt"
Private Const FORMNAME As String = "FrmLogin"
'5608 new blank user looping and getting error
Dim bmStopLostFocus As Boolean

Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim ilValue As Integer
    Dim ilValue8 As Integer
    Dim slDate As String
    Dim ilDatabase As Integer
    Dim ilLocation As Integer
    Dim ilSQL As Integer
    Dim ilForm As Integer
    Dim sMsg As String
    Dim iLoop As Integer
    Dim sCurDate As String
    Dim sAutoLogin As String
    Dim slTimeOut As String
    Dim slDSN As String
    Dim slStartIn As String
    Dim slStartStdMo As String
    Dim slTemp As String
    ReDim sWin(0 To 13) As String * 1
    Dim ilIsTntEmpty As Integer
    Dim ilIsShttEmpty As Integer
    Dim slDateTime1 As String
    Dim slDateTime2 As String
    Dim EmailExists_rst As ADODB.Recordset
    '5/11/11
    Dim blAddGuide As Boolean
    'dan 2/23/12 can't have error handler in error handler
    Dim blNeedToCloseCnn As Boolean
    Dim slXMLINIInputFile As String
    '5676
    Dim slRootPath As String
    Dim slPhotoPath As String
    Dim slBitmapPath As String
    Dim slLogoPath As String
    Dim slDBSection As String
    Dim slLocSection As String
    Dim slCommand As String
    '4/15/20
    Dim slName As String
    Dim slPassword As String
    '10000
    Dim slWebTemp As String
    'D.S. 06/20/18
    If App.PrevInstance Then
        MsgBox "Only one copy of Affiliate can be run at a time, sorry", vbOKOnly + vbInformation, "Counterpoint"
        End
    End If
    
    gCommandArgs
    sgCommand = Command$
    slCommand = UCase(sgCommand)
    blNeedToCloseCnn = False
    'Display gMsgBox
    'igShowMsgBox = True shows the gMsgBox.
    'igShowMsgBox = False does not show any gMsgBox
    
    'Warning: One thing to remember is that if you are expecting a return value from a gMsgBox
    'and you turn gMsgBox off then you need to make sure that you handle that case.
    'example:   ilRet = gMsgBox "xxxx"
    igShowMsgBox = True
 
    igDemoMode = False
    If InStr(slCommand, UCase("Demo")) Then
        igDemoMode = True
    End If
    'Used to speed-up testing exports with multiple files reduce record count needed to create a new file
    igSmallFiles = False
    If InStr(slCommand, UCase("SmallFiles")) Then
        igSmallFiles = True
    End If
    
    igAutoImport = False
    'If StrComp(sgCommand, "AutoImport") = 0 Then
    '    igAutoImport = True
    '    igShowMsgBox = False
    'End If
    
    igCompelAutoImport = False
    'If StrComp(sgCommand, "CompelAutoImport") = 0 Then
    '    igCompelAutoImport = True
    '    igShowMsgBox = False
    'End If
    If InStr(1, slCommand, UCase("CompelAutoImport"), vbTextCompare) > 0 Then
        igCompelAutoImport = True
        igShowMsgBox = False
    Else
        If InStr(1, slCommand, UCase("AutoImport"), vbTextCompare) > 0 Then
            igAutoImport = True
            igShowMsgBox = False
        End If
    End If
    slStartIn = CurDir$
    sgCurDir = CurDir$
    '10000 moved from below
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    sgWebServerSection = "WebServer"
    bgTestSystemAllowWebVendors = False
    If InStr(1, UCase(slStartIn), UCase("Test"), vbTextCompare) = 0 Then
        igTestSystem = False
        slDBSection = "Database"
        slLocSection = "Locations"
    Else
        igTestSystem = True
        igDemoMode = True
        slDBSection = "TestDatabase"
        slLocSection = "TestLocations"
        If gLoadOption("TestWebServer", "RootURL", slWebTemp) Then
            sgWebServerSection = "TestWebServer"
            igDemoMode = False
            If gLoadOption(sgWebServerSection, "AllowWebVendors", slWebTemp) Then
                If Trim$(UCase(slWebTemp)) = "YES" Or Trim$(UCase(slWebTemp)) = "Y" Then
                    bgTestSystemAllowWebVendors = True
                End If
            End If
        End If
    End If
    igShowVersionNo = 0
    If (InStr(1, UCase(slStartIn), UCase("Prod"), vbTextCompare) = 0) And (InStr(1, UCase(slStartIn), UCase("Test"), vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    
    mInit
    
    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF
   
   
    ilRet = 0
    ilLocation = False
    ilDatabase = False
    sgDatabaseName = ""
    sgReportDirectory = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgExeDirectory = ""
    sgLogoDirectory = ""
    sgPasswordAddition = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgCrystalDateForm = "yyyy,mm,dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "h:mma/p"
    sgShowTimeWSecForm = "h:mm:ssa/p"
    igWaitCount = 10
    igTimeOut = -1
    sgWallpaper = ""
    bgIgnoreDuplicateError = False
    '10000 moved above
'    sgStartupDirectory = CurDir$
'    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    sgLogoName = "rptlogo.bmp"
    sgNowDate = ""
    ilPos = InStr(1, sgCommand, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommand, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommand, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gIsDate(slDate) Then
            sgNowDate = slDate
        End If
    End If
    '5/26/13: Report Queue
    bgReportQueue = False
    ilPos = InStr(1, sgCommand, "/Q", 1)
    If ilPos > 0 Then
        bgReportQueue = True
    End If
    
    'If Not gLoadOption("Locations", "Logo", sgLogoPath) Then
    If Not gLoadOption(slLocSection, "Logo", sgLogoPath) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'Logo' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    'If Not gLoadOption("Database", "Name", sgDatabaseName) Then
    If Not gLoadOption(slDBSection, "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini " & slDBSection & " 'Name' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    End If
    
    ' NOTE:
    '   mLoadOption reads from Traffic.ini
    '   gLoadOption reads from Affiliat.ini
    If Not mLoadOption(slLocSection, "URL_Documentation", sgURL_Documentation) Then
        'gMsgBox "Affiliat.Ini " & slLocSection & " 'URL_Documentation' key is missing.", vbCritical
        'Unload frmLogin
        'Exit Sub
    End If
    
    If Not gLoadOption(slLocSection, "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'Reports' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    End If
    If Not gLoadOption(slLocSection, "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'Export' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    End If
    If Not gLoadOption(slLocSection, "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'Exe' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    End If
    If Not gLoadOption(slLocSection, "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'Logo' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    End If
    
        
    'Import is optional
    If gLoadOption(slLocSection, "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory, True)
    Else
        sgImportDirectory = ""
    End If
    
    If gLoadOption(slLocSection, "ContractPDF", sgContractPDFPath) Then
        sgContractPDFPath = gSetPathEndSlash(sgContractPDFPath, True)
    Else
        sgContractPDFPath = ""
    End If
    'TTP 10457 - ISCI Cross Reference Export
    If gLoadOption(slLocSection, "ISCIxRefExport", sgISCIxRefExportPath) Then
        sgISCIxRefExportPath = gSetPathEndSlash(sgISCIxRefExportPath, True)
    Else
        sgISCIxRefExportPath = sgExportDirectory
    End If

    '4/14/21: TTP 9052
    If Not gLoadOption(slLocSection, "WebNumber", sgWebNumber) Then
        sgWebNumber = "1"
    End If
    '5676 sgRootDrive.  If drive C doesn't exist, look for RootDrive in ini file, then test that value to make sure it exists.
    '8886
'    If Dir("c:\") = "" Then
    If gFolderExist("c:\") Then
        If gLoadOption(slLocSection, "RootDrive", sgRootDrive) Then
            sgRootDrive = gSetPathEndSlash(sgRootDrive, True)
            If Dir(sgRootDrive) = "" Then
                sgRootDrive = "C:\"
            End If
        Else
            sgRootDrive = "C:\"
        End If
    Else
        sgRootDrive = "C:\"
    End If
    '7496
    If Not gLoadOption(slLocSection, "AudioExtension", sgAudioExtension) Then
        sgAudioExtension = ".mp2"
    Else
        If Len(sgAudioExtension) = 0 Then
            sgAudioExtension = ".mp2"
        ElseIf InStr(sgAudioExtension, ".") <> 1 Then
            sgAudioExtension = "." & sgAudioExtension
        End If
    End If
    sgAudioExtension = LCase(sgAudioExtension)
    'Commented out below because I can't see why you would need a backslash
    'on the end of a DSN name
    'sgDatabaseName = gSetPathEndSlash(sgDatabaseName)
    sgReportDirectory = gSetPathEndSlash(sgReportDirectory, True)
    sgExportDirectory = gSetPathEndSlash(sgExportDirectory, True)
    sgExeDirectory = gSetPathEndSlash(sgExeDirectory, True)
    sgLogoDirectory = gSetPathEndSlash(sgLogoDirectory, True)
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", sBuffer) Then
        If sBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption(slLocSection, "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption(slLocSection, "Wallpaper", sgWallpaper)
    
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    ' JD Novelty - 09-06-22 - Added code to hide certain things if using the Novelty web system.
    gIsUsingNovelty = False
    slWebTemp = "0"
    If gLoadOption("WebServer", "UsingNovelty", slWebTemp) Then
        slWebTemp = UCase$(Left$(Trim$(slWebTemp), 1))
        If slWebTemp = "1" Or slWebTemp = "T" Or slWebTemp = "Y" Then
            gIsUsingNovelty = True
        End If
    End If
    
    If Not gLoadOption(slLocSection, "DBPath", sgDBPath) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'DBPath' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    If Not mCheckVersion() Then
        Unload frmLogin
        Exit Sub
    End If
    
    'Set Message folder
    If Not gLoadOption(slLocSection, "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini " & slLocSection & " 'DBPath' key is missing.", vbCritical
        Unload frmLogin
        Exit Sub
    Else
        sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory, True) & "Messages\"
'        sgMsgDirectory = CurDir
'        If InStr(1, sgMsgDirectory, "Data", vbTextCompare) Then
'            sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory) & "Messages\"
'        Else
'            sgMsgDirectory = sgExportDirectory
'        End If
    End If
    
    '4/21/18: SQL Trace
    sgSQLTrace = ""
    hgSQLTrace = -1
    'If gLoadOption("Database", "SQLTrace", sgSQLTrace) Then
    If gLoadOption(slDBSection, "SQLTrace", sgSQLTrace) Then
        If Trim$(sgSQLTrace) <> "" Then
            sgSQLTrace = UCase$(Left$(Trim$(sgSQLTrace), 1))
            If sgSQLTrace = "Y" Then
                gLogMsgWODT "ON", hgSQLTrace, sgMsgDirectory & "SQLTrace.txt"
            Else
                sgSQLTrace = ""
            End If
        End If
    End If
    
    
    '12/4/12: Check if Logging activity for a file
    mLogActivityFileName
    '12/4/12: end of change
    
    ' Not sure what section this next item is coming from. The original code did not specify.
    'Call gLoadOption("SQLSpec", "WaitCount", sBuffer)
    'igWaitCount = Val(sBuffer)
    
    On Error GoTo ErrHand
    Set cnn = New ADODB.Connection
    'Set env = rdoEnvironments(0)
    'cnn.CursorDriver = rdUseOdbc
    
    'Set cnn = cnn.OpenConnection(dsName:="Affiliate", Prompt:=rdDriverCompleteRequired)
    ' The default timeout is 15 seconds. This always fails on my PC the first time I run this program.


    slDSN = sgDatabaseName
    'ttp 4905.  Need to try connection. If it fails, try one more time, after sleeping.
    'cnn.Open "DSN=" & slDSN
    
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & slDSN

    On Error GoTo ErrHand
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & slDSN
    End If
    
    'Example of using a user name and password
    'cnn.Open "DSN=" & slDSN, "Master", "doug"
    Set rst = New ADODB.Recordset
    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
 
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(slDSN, Len(slDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        slDSN = Left(slDSN, Len(slDSN) - 1)
    End If
    'Set cnn = cnn.OpenConnection(dsName:=slDSN, Prompt:=rdDriverCompleteRequired)
    'If igTimeOut >= 0 Then
    '    cnn.QueryTimeout = igTimeOut
    'End If

    '5/17/13: Moved from Main to here so that DDF can be checked prior to making any SQL or API calls
    'Start the Pervasive API engine
    If Not mOpenPervasiveAPI Then
        Unload frmLogin
        Exit Sub
    End If
    
    If Not gCheckDDFDates() Then
        Unload frmLogin
        Exit Sub
    End If

    'Code modified for testing
    txtUID.Text = ""
    txtPWD.Text = ""
    
    'Debug
    'txtUID.Text = "csi"
    'txtPWD.Text = "login3428"
    
    
    
    'Test for Guide- if not added- add
    'SQLQuery = "Select MAX(ustCode) from ust"
    'Set rst = gSQLSelectCall(SQLQuery)
    ''If rst(0).Value = 0 Then
    'If IsNull(rst(0).Value) Then
    ''5/11/11
    '    blAddGuide = True
    'Else
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            blAddGuide = True
        Else
            blAddGuide = False
            imGuideUstCode = rst!ustCode
        End If
    'End If
    If blAddGuide Then
    '5/11/11
        'SQLQuery = "INSERT INTO ust(ustName, ustPassword, ustState)"
        'SQLQuery = SQLQuery & "VALUES ('Guide', 'Guide', 0)"
        sCurDate = Format(Now, sgShowDateForm)
        For iLoop = 0 To 13 Step 1
            sWin(iLoop) = "I"
        Next iLoop
        '5/11/11
        mResetGuideGlobals
        SQLQuery = "INSERT INTO ust(ustName, ustReportName, ustPassword, "
        SQLQuery = SQLQuery & "ustState, ustPassDate, ustActivityLog, ustWin1, "
        SQLQuery = SQLQuery & "ustWin2, ustWin3, ustWin4, "
        SQLQuery = SQLQuery & "ustWin5, ustWin6, ustWin7, "
        SQLQuery = SQLQuery & "ustWin8, ustWin9, ustPledge, "
        SQLQuery = SQLQuery & "ustExptSpotAlert, ustExptISCIAlert, ustTrafLogAlert, "
        SQLQuery = SQLQuery & "ustWin10, ustWin11, ustWin12, ustWin13, "
        SQLQuery = SQLQuery & "ustWin14, ustWin15, ustPhoneNo, ustCity, ustEMailCefCode, ustAllowedToBlock, "
        SQLQuery = SQLQuery & "ustWin16, "
        SQLQuery = SQLQuery & "ustUserInitials, "
        SQLQuery = SQLQuery & "ustDntCode, "
        SQLQuery = SQLQuery & "ustAllowCmmtChg, "
        SQLQuery = SQLQuery & "ustAllowCmmtDelete, "
        SQLQuery = SQLQuery & "ustWin17, "
        SQLQuery = SQLQuery & "ustChgExptPriority, "
        SQLQuery = SQLQuery & "ustExptSpec, "
        SQLQuery = SQLQuery & "ustChgRptPriority, "
        SQLQuery = SQLQuery & "ustUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "VALUES ('" & "Guide" & "', "
        SQLQuery = SQLQuery & "'" & "System" & "', '" & "Guide" & "', "
        SQLQuery = SQLQuery & 0 & ", '" & Format$(sCurDate, sgSQLDateForm) & "', '" & "V" & "', '" & sgUstWin(1) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(2) & "', '" & sgUstWin(3) & "', '" & sgUstWin(4) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(5) & "', '" & sgUstWin(6) & "', '" & sgUstWin(7) & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(8) & "', '" & sgUstWin(9) & "', '" & sgUstPledge & "', "
        SQLQuery = SQLQuery & "'" & sgExptSpotAlert & "', '" & sgExptISCIAlert & "', '" & sgTrafLogAlert & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(10) & "', '" & sgUstWin(11) & "', '" & sgUstWin(12) & "', '" & sgUstWin(13) & "', "
        SQLQuery = SQLQuery & "'" & sgUstClear & "', '" & sgUstDelete & "', "
        SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', " & 0 & ", '" & "Y" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(sgUstWin(0)) & "', "
        SQLQuery = SQLQuery & "'" & "G" & "', "
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtChg & "', "
        SQLQuery = SQLQuery & "'" & sgUstAllowCmmtDelete & "', "
        SQLQuery = SQLQuery & "'" & sgUstWin(14) & "', "
        SQLQuery = SQLQuery & "'" & sgChgExptPriority & "', "
        SQLQuery = SQLQuery & "'" & sgExptSpec & "', "
        SQLQuery = SQLQuery & "'" & sgChgRptPriority & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        cnn.BeginTrans
        blNeedToCloseCnn = True
        'cnn.ConnectionTimeout = 30  ' Increase from the default of 15 to 30 seconds.
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "Login-Form_Load"
            cnn.RollbackTrans
            tmcTerminate.Enabled = True
            Exit Sub
        End If
        cnn.CommitTrans
        blNeedToCloseCnn = False
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            imGuideUstCode = rst!ustCode
        Else
            imGuideUstCode = 0
        End If
    End If
    
    gUsingCSIBackup = False
    gUsingXDigital = False
    gWegenerExport = False
    gOLAExport = False
    sgDelNet = "N"
    ' Dan M added spfusingFeatures2
    SQLQuery = "SELECT spfGClient, spfGAlertInterval, spfGUseAffSys, spfUsingFeatures7, spfUsingFeatures2, spfUsingFeatures8, spfUsingFeatures5, spfSDelNet"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    
    If Not rst.EOF Then
        If UCase(rst!spfGUseAffSys) <> "Y" Then
            gMsgBox "The Affiliate system has not been activated.  Please call Counterpoint.", vbCritical
            Unload frmLogin
            Exit Sub
        End If
        ilValue8 = Asc(rst!spfUsingFeatures8)
        If (ilValue8 And ALLOWMSASPLITCOPY) <> ALLOWMSASPLITCOPY Then
            gUsingMSARegions = False
        Else
            gUsingMSARegions = True
        End If
        If (ilValue8 And ISCIEXPORT) <> ISCIEXPORT Then
            gISCIExport = False
        Else
            gISCIExport = True
        End If
        ilValue = Asc(rst!spfUsingFeatures7)
        If (ilValue And CSIBACKUP) <> CSIBACKUP Then
            gUsingCSIBackup = False
        Else
            gUsingCSIBackup = True
        End If
        
        If ((ilValue And XDIGITALISCIEXPORT) <> XDIGITALISCIEXPORT) And ((ilValue8 And XDIGITALBREAKEXPORT) <> XDIGITALBREAKEXPORT) Then
            gUsingXDigital = False
        Else
            gUsingXDigital = True
        End If
        If (ilValue And WEGENEREXPORT) <> WEGENEREXPORT Then
            gWegenerExport = False
        Else
            gWegenerExport = True
        End If
        If (ilValue And OLAEXPORT) <> OLAEXPORT Then
            gOLAExport = False
        Else
            gOLAExport = True
        End If
        ilValue = Asc(rst!spfusingfeatures2)
        If (ilValue And STRONGPASSWORD) <> STRONGPASSWORD Then
            bgStrongPassword = False
        Else
            bgStrongPassword = True
        End If
        ilValue = Asc(rst!spfusingfeatures5)
        If (ilValue And REMOTEEXPORT) = REMOTEEXPORT Then
            bgRemoteExport = True
        Else
            bgRemoteExport = False
        End If
        
        sgDelNet = rst!spfSDelNet
    End If
    
    If Not rst.EOF Then
        sgClientName = Trim$(rst!spfgClient)
        igAlertInterval = rst!spfGAlertInterval
    Else
        sgClientName = "Unknown"
        gMsgBox "Client name is not defined in Site Options"
        igAlertInterval = 0
    End If
    
    If InStr(1, sgCommand, "NoAlerts", vbTextCompare) > 0 Then
        'For Debug ONLY
        igAlertInterval = 0
    End If
    If Trim$(sgNowDate) = "" Then
        If InStr(1, sgClientName, "XYZ Broadcasting", vbTextCompare) > 0 Then
            sgNowDate = "12/14/2015"
        End If
    End If

    '6/3/08: Removed siteLastDateArch and siteNoMnthRetain(moved to spfRetainAffSpot)
    'SQLQuery = "SELECT siteLastDateArch, siteNoMnthRetain"
    'SQLQuery = SQLQuery + " FROM Site"
    'Set rst = gSQLSelectCall(SQLQuery)
   '
   ' If Not rst.EOF Then
   '     If rst!siteLastDateArch <> Null Then
   '         sgLastDateArch = Format(rst!siteLastDateArch, "mm/dd/yyyy")
   '         slStartStdMo = Format$(gObtainStartStd(gNow()), "mm/dd/yyyy")
   '         slTemp = DateAdd("d", 1, sgLastDateArch)
   '         sgNumMoToRetain = rst!siteNoMnthRetain
   '         igNumMoBehind = gCalcMoBehind(slStartStdMo, gObtainEndStd(slTemp))
   '     Else
   '         sgLastDateArch = ""
   '         sgNumMoToRetain = rst!siteNoMnthRetain
   '         igNumMoBehind = 0
   '     End If
   ' End If
    
'    gUsingUnivision = False
'    gUsingWeb = False
'    'sgExportISCI = "B"
'    sgExportISCI = ""
'    sgShowByVehType = "N"
'    sgRCSExportCart4 = "Y"
'    sgRCSExportCart5 = "N"
'    sgUsingStationID = "N"
'    SQLQuery = "SELECT * From Site Where siteCode = 1"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        gUsingUnivision = rst!siteMarketron
'        gUsingWeb = rst!siteWeb
'        'sgExportISCI = rst!siteISCIExport
'        sgShowByVehType = rst!siteShowVehType
'        sgRCSExportCart4 = rst!siteExportCart4
'        sgRCSExportCart5 = rst!siteExportCart5
'        sgUsingStationID = rst!siteUsingStationID
'        If gUsingWeb Then
'            While Not gVerifyWebIniSettings()
'                frmWebIniOptions.Show vbModal
'                If Not igWebIniOptionsOK Then
'                    Unload frmLogin
'                    Exit Sub
'                End If
'            Wend
'            If Not igDemoMode Then
'                If Not gTestAccessToWebServer() Then
'                    gMsgBox "WARNING!" & vbCrLf & vbCrLf & _
'                           "Web Server Access Error: The Affiliate System does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
'                    "No data will be exported to the web site." & vbCrLf & _
'                    "No data will be imported from the web site." & vbCrLf & _
'                    "Sign off system immediately and contact system administrator.", vbExclamation
'                End If
'            End If
'        End If
'    End If

    ilRet = gInitGlobals()
    If ilRet = 0 Then
        While Not gVerifyWebIniSettings()
            frmWebIniOptions.Show vbModal
            If Not igWebIniOptionsOK Then
                Unload frmLogin
                Exit Sub
            End If
        Wend
    End If
    
    'Call gLoadOption("Database", "AutoLogin", sAutoLogin)
    Call gLoadOption(slDBSection, "AutoLogin", sAutoLogin)
    If igAutoImport Then
        txtUID.Text = "csi" '"guide"
        txtPWD.Text = "login2203"
        sgUserName = "AutoImport"
        sgReportName = "AutoImport"
        Call cmdOk_Click
        Unload frmLogin
        Unload frmWebImportAiredSpot
        End
    End If
    If igCompelAutoImport Then
        txtUID.Text = "csi" '"guide"
        txtPWD.Text = "login2203"
        sgUserName = "CompelAutoImport"
        sgReportName = "CompelAutoImport"
        Call cmdOk_Click
        Unload frmLogin
        Unload frmImportWegener
        End
    End If
    On Error GoTo ErrHand
    If Not igAutoImport And Not igCompelAutoImport Then
        ilRet = mInitAPIReport()      '4-19-04
    End If
    ilRet = gTestWebVersion()
    'Move report logo to local C drive (c:\csi\rptlogo.bmp)
    ' 5676: don't assume c drive
    slRootPath = sgRootDrive & "CSI\"
    slPhotoPath = slRootPath & "RptLogo.jpg"
    slBitmapPath = slRootPath & "RptLogo.Bmp"
    slLogoPath = sgLogoPath & "RptLogo.Bmp"
    ilRet = 0
    On Error GoTo mTrafficStartUpErr:
'8-19-14 no need to copy the logos to root folder; will be retrieved from the logo ini entry
'    slDateTime1 = FileDateTime(slBitmapPath)
'    If ilRet <> 0 Then
'        ilRet = 0
'        MkDir slRootPath
'        If ilRet = 0 Then
'            FileCopy slLogoPath, slBitmapPath
'        Else
'            FileCopy slLogoPath, slBitmapPath
'        End If
'    Else
'        ilRet = 0
'        slDateTime2 = FileDateTime(slLogoPath)
'        If ilRet = 0 Then
'            If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
'                FileCopy slLogoPath, slBitmapPath
'            End If
'        End If
'    End If
'    On Error GoTo 0
'        'ttp 5260
'    slLogoPath = sgLogoPath & "RptLogo.jpg"
'    If Dir(slLogoPath) > "" Then
'        If Dir(slPhotoPath) = "" Then
'            FileCopy slLogoPath, slPhotoPath
'        'ok, both exist.  is logopath's more recent?
'        Else
'            '5461
'            slDateTime1 = FileDateTime(slLogoPath)
'            slDateTime2 = FileDateTime(slPhotoPath)
'            If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
'                FileCopy slLogoPath, slPhotoPath
'            End If
'        End If
'    End If
    
    
    
'    ilRet = 0
'    On Error GoTo mTrafficStartUpErr:
'    slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
'    If ilRet <> 0 Then
'        ilRet = 0
'        MkDir "C:\CSI"
'        If ilRet = 0 Then
'            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
'        Else
'            FileCopy sgDBPath & "RptLogo.Bmp", sgLogoPath & "RptLogo.Bmp"
'        End If
'    Else
'        ilRet = 0
'        slDateTime2 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
'        If ilRet = 0 Then
'            If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
'                FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
'            End If
'        End If
'    End If
'     'ttp 5260
'    If Dir(sgLogoPath & "RptLogo.jpg") > "" Then
'        If Dir("c:\csi\RptLogo.jpg") = "" Then
'            FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
'        'ok, both exist.  is logopath's more recent?
'        Else
'            slDateTime1 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
'            slDateTime2 = FileDateTime("C:\CSI\RptLogo.jpg")
'            If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
'                FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
'            End If
'        End If
'    End If
    'Determine number if X-Digital HeadEnds
    ReDim sgXDSSection(0 To 0) As String
    slXMLINIInputFile = gXmlIniPath(True)
    If LenB(slXMLINIInputFile) <> 0 Then
        ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    End If
    'Test to see if this function has been ran before, if so don't run it again
    igEmailNeedsConv = False
    'SQLQuery = "SELECT Max(emtCode) from EMT"
    'Set EmailExists_rst = gSQLSelectCall(SQLQuery)
    'If Not EmailExists_rst.EOF Then
    '    If IsNull(EmailExists_rst(0).Value) Then
    '        igEmailNeedsConv = True
    '    End If
    'End If
    
    '9/13/11: Doug- This test was removed
    'SQLQuery = "SELECT Count(emtCode) FROM EMT"
    'Set EmailExists_rst = gSQLSelectCall(SQLQuery)
    'If EmailExists_rst(0).Value = 0 Then
    '    SQLQuery = "SELECT Count(shttCode) FROM SHTT"
    '    Set EmailExists_rst = gSQLSelectCall(SQLQuery)
    '    If EmailExists_rst(0).Value <> 0 Then
    '        igEmailNeedsConv = True
    '    End If
    'End If
    
    'Dan M 4/30/09 setting ust here so can use for 'lostFocus' and cmdOk
    '6/23/09:  Moved to Blank test and into cmd_Ok
    'SQLQuery = "SELECT * FROM ust "
    'Set rst = gSQLSelectCall(SQLQuery)
    If InStr(1, sgCommand, UCase("FROMTRAFFIC"), vbTextCompare) > 0 Then
        ilRet = gParseItem(sgCommand, 2, "\", slName)
        ilRet = gParseItem(sgCommand, 3, "\", slPassword)
        txtUID.Text = slName
        txtPWD.Text = slPassword
        cmdOk_Click
    End If
    
    Exit Sub
    
TableDoesNotExist:
    ilRet = False
    Resume Next

mTrafficStartUpErr:
    ilRet = Err.Number
    Resume Next

mReadFileErr:
    ilRet = Err.Number
    Resume Next
    'ttp 4905
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    'ttp 5217
    gHandleError "", FORMNAME & "-Form_Load"
    'ttp 4905 need to quit app
    bmNoPervasive = True
    If blNeedToCloseCnn Then
        cnn.RollbackTrans
    End If
    'unload affiliate  ttp 4905
    tmcTerminate.Enabled = True
End Sub


Private Sub cmdCancel_Click()
    If sgSQLTrace = "Y" Then
        gLogMsgWODT "W", hgSQLTrace, "SQL Overall Time: " & gTimeString(lgTtlTimeSQL / 1000, True)
        gLogMsgWODT "C", hgSQLTrace, ""
    End If
    Unload frmLogin
    Set frmLogin = Nothing
End Sub


Private Sub cmdOk_Click()
    Dim iRet As Integer
    Dim sName As String
    Dim sPass As String
    Dim iUpper As Integer
    Dim iFound As Integer
    Dim iZoneFd As Integer
    Dim iFedFd As Integer
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim sChar As String * 1
    Dim lCode As Long
    Dim ilRet As Integer
    
    Dim ilPos As Integer
    Dim slDriveLetter As String
    Dim llSectorsPerCluster As Long
    Dim llBytesPerSector As Long
    Dim llNumberOfFreeClusters As Long
    Dim llTotalNumberOfClusters As Long
    Dim llRet As Long
    Dim flFreeGigBytes As Single
    Dim blSuccess As Boolean
    ' ttp 5608
    Dim blOk As Boolean
    Dim slStoredPassword As String
    
    'txtUID.Text = "csi"
    'txtPWD.Text = "login6531"
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    sgUserNameToPassToTraffic = Trim$(txtUID.Text)
    sgUserPasswordToPassToTraffic = Trim$(txtPWD.Text)
    sgSpecialPassword = ""
    bgInternalGuide = False
    igPasswordOk = False
    sName = Trim$(txtUID.Text)
    If StrComp(sName, "CSI", vbTextCompare) = 0 Then
        sName = "Guide"
    End If
    sPass = Trim$(txtPWD.Text)
 
    
    'Dan M 4/29/09  " " and regular guide
    If sPass = "" Then
        If Not mTestForBlankPassword(True, False) Then
            Screen.MousePointer = vbDefault
            txtPWD.SetFocus
            Exit Sub
        End If
        'txtPWD has been changed from blank to new password
        sPass = Trim$(txtPWD.Text)
    End If
    Screen.MousePointer = vbHourglass
    lgEMailCefCode = 0
    If (StrComp(sName, "Counterpoint", 1) <> 0) Or (StrComp(sPass, "JD#41", 1) <> 0) Then
        If (StrComp(sName, "Guide", 1) <> 0) Or (InStr(1, sPass, smSpecial, 1) <> 1) Or (Len(sPass) <= 5) Then
            'SQLQuery = "SELECT ustPassword, ustState, ustCode FROM ust WHERE (ustName = '" & sName & "')"
            'Set rst = gSQLSelectCall(SQLQuery)
            'If rst.EOF Then
            '    Beep
            '    Screen.MousePointer = vbDefault
            '    txtUID.SetFocus
            '    Exit Sub
            'End If
            'If (StrComp(sPass, Trim$(rst!ustPassword), 1) <> 0) Then
            '    Beep
            '    Screen.MousePointer = vbDefault
            '    txtUID.SetFocus
            '    Exit Sub
            'End If
            'If rst!ustState <> 0 Then
            '    Beep
            '    Screen.MousePointer = vbDefault
            '    txtUID.SetFocus
            '    Exit Sub
            'End If
            
            'dan 4/30/09 now setting in 'lostFocus'.  If user's password is blank go directly to NewPassword form
            '6/23/09:  Added SQL call back in so that the Name and password can be checked and removed from end of Load Form.
            SQLQuery = "SELECT * FROM ust "
            Set rst = gSQLSelectCall(SQLQuery)
            If Not igAutoImport And Not igCompelAutoImport Then
                iFound = False
            Else
                iFound = True
            End If
                Do While Not rst.EOF
                    If (StrComp(sName, Trim$(rst!ustname), 1) = 0) Then
                        'dan M 4/14/09 proper case for user
                        sName = Trim(rst!ustname)
                        slStoredPassword = Trim$(rst!ustpassword)
                        '5608 user without strong password can slip through
                        If Not igAutoImport And Not igCompelAutoImport Then
                            If bgStrongPassword And Not gStrongPassword(slStoredPassword) Then
                                If StrComp(sPass, slStoredPassword, 1) = 0 Then
                                'user must change password
                                    If mGetNewPassword(False) Then
                                        blOk = True
                                    End If
                                End If
                            ElseIf bgStrongPassword Then
                                If StrComp(sPass, slStoredPassword, vbBinaryCompare) = 0 Then
                                    blOk = True
                                End If
                            ElseIf (StrComp(sPass, slStoredPassword, 1) = 0) Then
                                    blOk = True
                            End If
                        Else
                            blOk = True
                        End If
                       ' If (StrComp(sPass, Trim$(rst!ustpassword), 1) = 0) Then
                       ' Dan M 4/14/09 blank password was tested here briefly: now in txtUid_lostFocus
                        If blOk Then
                            If rst!ustState = 0 Then
                                iFound = True
                                sgUserName = sName
                                sgReportName = rst!ustReportName
                                igUstCode = rst!ustCode
                                sgUstWin(0) = rst!ustWin16
                                sgUstWin(1) = rst!ustWin1
                                sgUstWin(2) = rst!ustWin2
                                sgUstWin(3) = rst!ustWin3
                                sgUstWin(4) = rst!ustWin4
                                sgUstWin(5) = rst!ustWin5
                                sgUstWin(6) = rst!ustWin6
                                sgUstWin(7) = rst!ustWin7
                                sgUstWin(8) = rst!ustWin8
                                sgUstWin(9) = rst!ustWin9
                                sgUstWin(10) = rst!ustWin10
                                sgUstWin(11) = rst!ustWin11
                                sgUstWin(12) = rst!ustWin12
                                sgUstWin(13) = rst!ustWin13
                                sgUstWin(14) = rst!ustWin17
                                sgUstClear = rst!ustWin14
                                sgUstActivityLog = rst!ustActivityLog
                                sgUstDelete = rst!ustWin15
                                sgUstPledge = rst!ustPledge
                                sgUstAllowCmmtChg = rst!ustAllowCmmtChg
                                sgUstAllowCmmtDelete = rst!ustAllowCmmtDelete
                                sgExptSpotAlert = rst!ustExptSpotAlert
                                sgExptISCIAlert = rst!ustExptISCIAlert
                                sgTrafLogAlert = rst!ustTrafLogAlert
                                sgPhoneNo = rst!ustPhoneNo
                                igUstSSMnfCode = rst!ustSSMnfCode
                                sgCity = rst!ustCity
                                sgAllowedToBlock = rst!ustAllowedToBlock
                                sgEMail = ""
                                lgEMailCefCode = rst!ustEmailcefcode
                                sgChgExptPriority = rst!ustChgExptPriority
                                sgExptSpec = rst!ustExptSpec
                                sgChgRptPriority = rst!ustChgRptPriority
                                igPasswordOk = True
                                
                                'Dan M 4/15/09 guide now limited functionality.  Change database if not limited
                                If (StrComp(sName, "GUIDE", vbTextCompare) = 0) Then
                                    blSuccess = mLimitGuide(rst)
                                    If Not blSuccess Then
                                        gMsg = "Error finding guide data: frmLogin-mLimitGuide"
                                        gLogMsg gMsg, "AffErrorLog.txt", False
                                        If Not igAutoImport And Not igCompelAutoImport Then
                                            gMsgBox gMsg, vbCritical
                                        End If
                                        Set rst = Nothing
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    End If
                                End If
                                Exit Do
                            End If
                        Else
                            Beep
                            Screen.MousePointer = vbDefault
                            If Not igAutoImport And Not igCompelAutoImport Then
                                txtUID.SetFocus
                            End If
                            txtUID.SetFocus
                            Exit Sub
                        End If
                    End If
                    rst.MoveNext
                Loop
                If Not iFound Then
                    Beep
                    Screen.MousePointer = vbDefault
                    txtUID.SetFocus
                    Exit Sub
                End If
                '11/8/14: Limit Remote Export users
                mLimitRemoteExportUsers
           ' End If Dan M removed 9-04-09: stopping compiling
        Else
            sgUserName = sName
            '5/11/11
            'igUstCode = 1
            igUstCode = imGuideUstCode
            For iLoop = 0 To 14 Step 1
                sgUstWin(iLoop) = "I"
            Next iLoop
            sgUstDelete = "Y"
            sgUstClear = "Y"
            sgUstActivityLog = "V"
            sgUstPledge = "Y"
            sgUstAllowCmmtChg = "Y"
            sgUstAllowCmmtDelete = "Y"
            sgExptSpotAlert = "Y"
            sgExptISCIAlert = "Y"
            sgTrafLogAlert = "Y"
            sgPhoneNo = ""
            igUstSSMnfCode = 0
            sgCity = ""
            sgAllowedToBlock = "Y"
            sgChgExptPriority = "Y"
            sgExptSpec = "Y"
            sgChgRptPriority = "Y"
            sgEMail = ""
            If Len(sPass) = 9 Then
                If StrComp(Trim$(txtUID.Text), "Guide", vbTextCompare) = 0 Then
                    Beep
                    Screen.MousePointer = vbDefault
                    txtUID.SetFocus
                    Exit Sub
                End If
                '11/4/09:  Only set special if the special password was entered
                'sgSpecialPassword = Mid$(sPass, 6)
                If (InStr(1, sPass, smSpecial, 1) = 1) Then
                    sgSpecialPassword = Mid$(sPass, 6)
                    bgInternalGuide = True
                End If
                If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
                   StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                    igPasswordOk = True
                End If
            End If
        End If
    Else
        sgUserName = sName
        '5/11/11
        'igUstCode = 0
        igUstCode = imGuideUstCode
        For iLoop = 0 To 14 Step 1
            sgUstWin(iLoop) = "I"
        Next iLoop
        sgUstPledge = "Y"
        sgExptSpotAlert = "Y"
        sgExptISCIAlert = "Y"
        sgTrafLogAlert = "Y"
        sgPhoneNo = ""
        igUstSSMnfCode = 0
        sgCity = ""
        sgAllowedToBlock = "Y"
        sgEMail = ""
        sgChgExptPriority = "Y"
        sgExptSpec = "Y"
        sgChgRptPriority = "Y"
    End If
    'SQLQuery = "SELECT cmfCode, cmfComment FROM " & """" & "CMF_Boilerplate" & """" & " cmf"
    'Set rst = gSQLSelectCall(SQLQuery, rdOpenStatic)
    'Do While Not rst.EOF
    '    lCode = rst!cmfCode
    '    sName = rst!cmfComment
    '    rst.MoveNext
    'Loop
    
    If Not mConversionCheck() Then
        Exit Sub
    End If
    
    'D.S. 02/26/13 Moved the pop routines to a common place to support unattended users that don't log in
    'such as auto import.
    If igAutoImport Or igCompelAutoImport Then
        iRet = gPopAll
    End If
    
    ReDim tgCifCpfInfo1(0 To 0) As CIFCPFINFO1
    ReDim tgCrfInfo1(0 To 0) As CRFINFO1

    SQLQuery = "SELECT spfUseCartNo, spfRemoteUsers, spfUsingFeatures2, spfUsingFeatures5, spfUsingFeatures9, spfSportInfo, spfUseProdSptScr"
    'SQLQuery = SQLQuery + " FROM SPF_Site_Options spf"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        sgSpfUseCartNo = rst!spfUseCartNo
        sgSpfRemoteUsers = rst!spfRemoteUsers
        'If rst!spfUsingFeatures2 <> Null Then
        If IsNull(rst!spfSportInfo) Or (Len(rst!spfSportInfo) = 0) Then
            sgSpfSportInfo = Chr$(0)
        Else
            sgSpfSportInfo = rst!spfSportInfo
        End If
        If IsNull(rst!spfusingfeatures2) Or (Len(rst!spfusingfeatures2) = 0) Then
            sgSpfUsingFeatures2 = Chr$(0)
        Else
            sgSpfUsingFeatures2 = rst!spfusingfeatures2
        End If
        If IsNull(rst!spfusingfeatures5) Or (Len(rst!spfusingfeatures5) = 0) Then
            sgSpfUsingFeatures5 = Chr$(0)
        Else
            sgSpfUsingFeatures5 = rst!spfusingfeatures5
        End If
        If IsNull(rst!spfUsingFeatures9) Or (Len(rst!spfUsingFeatures9) = 0) Then
            sgSpfUsingFeatures9 = Chr$(0)
        Else
            sgSpfUsingFeatures9 = rst!spfUsingFeatures9
        End If
        sgSpfUseProdSptScr = rst!spfUseProdSptScr
    Else
        sgSpfUseCartNo = "Y"
        sgSpfRemoteUsers = "N"
        sgSpfUsingFeatures2 = Chr$(0)
        sgSpfUsingFeatures5 = Chr$(0)
        sgSpfUsingFeatures9 = Chr$(0)
        sgSpfSportInfo = Chr$(0)
        sgSpfUseProdSptScr = "A"
    End If
    
    mCheckSecurity
    
    sgReplacementStamp = ""
    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        iRet = gObtainReplacments()
    Else
        ReDim tgRBofRec(0 To 0) As BOFREC
        ReDim tgSplitNetLastFill(0 To 0) As SPLITNETLASTFILL
    End If

    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) <> STATIONINTERFACE) Then
        gUsingUnivision = False
        gUsingWeb = False
    End If
    
    
    mCreateStatustype
    mCreateExportSpec
    
    DARKGREEN = RGB(0, 128, 0)
    
    'Check Disk Space
    ilPos = InStr(1, sgDBPath, ":", vbTextCompare)
    If ilPos > 0 Then
        slDriveLetter = Left$(sgDBPath, ilPos)
    Else
        ilPos = InStr(1, sgDBPath, "\", vbTextCompare)
        If ilPos > 0 Then
            ilPos = InStr(ilPos + 2, sgDBPath, "\", vbTextCompare)
            If ilPos > 0 Then
                slDriveLetter = Left$(sgDBPath, ilPos)
            End If
        End If
    End If
    
    'D.S. 5/15/09 Don't show this message in debug mode
    If igShowVersionNo <> 2 Then
        llRet = GetDiskFreeSpace(slDriveLetter, llSectorsPerCluster, llBytesPerSector, llNumberOfFreeClusters, llTotalNumberOfClusters)
        If llRet <> 0 Then
            flFreeGigBytes = (CDbl(llNumberOfFreeClusters) * llSectorsPerCluster * llBytesPerSector) / CDbl(1073741824)
            If flFreeGigBytes <= 5# Then
                llRet = 10 * flFreeGigBytes
                If Not igAutoImport And Not igCompelAutoImport Then
                    gMsgBox "Warning: only " & gLongToStrDec(llRet, 1) & " gig space remains on the server, inform your IT department that more free disk space is required", vbOKOnly + vbExclamation, "WARNING"
                End If
            End If
        End If
    End If
    'Dan 5666
       ' mCheckShortDateForm

    If igAutoImport Or igCompelAutoImport Then
        igPasswordOk = True
    End If
    
    '5/26/13: Report Queue
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        bgReportQueue = True
    End If
    
    Unload frmLogin
    Screen.MousePointer = vbDefault
    
    If igAutoImport Then
        sgUserName = "AutoImport"
        sgReportName = "AutoImport"
        If Not mOpenPervasiveAPI Then
            Exit Sub
        End If
        If igGGFlag <> 0 Then
            frmWebImportAiredSpot.Show
        End If
        gVendorToWebAllowed
'   D.S. 09/26/19 Moved Else down below
'    Else
'        If igPasswordOk Then
'            Call gCheckForContFiles
'            frmMain.Show
'        End If
    End If
    
    If igCompelAutoImport Then
        sgUserName = "CompelAutoImport"
        sgReportName = "CompelAutoImport"
        If Not mOpenPervasiveAPI Then
            Exit Sub
        End If
        If igGGFlag <> 0 Then
            frmImportWegener.Show
        End If
        gVendorToWebAllowed
    End If
    
    'D.S. Else from above moved to here
    If Not igCompelAutoImport And Not igAutoImport Then
        If igPasswordOk Then
            Call gCheckForContFiles
            frmMain.Show
        End If
    End If
    Set frmLogin = Nothing
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", FORMNAME & "-cmdOK_Click"
    Unload frmLogin
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub tmcTerminate_Timer()
    Unload Me
End Sub

Private Sub txtPWD_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtUID_LostFocus()
    If Not bmStopLostFocus Then
        mTestForBlankPassword False, True
    End If
End Sub

Private Sub txtUID_GotFocus()
'ttp 4905 dan M  pervasive not open. catch here.
'    If bmNoPervasive Then
'
'    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Function mTestForBlankPassword(blTestGuide As Boolean, blCallClick As Boolean) As Boolean
'4/29/09 if password = "" then go directly to new password form.
Dim slSQLQuery As String
Dim slName As String
Dim slRsName As String
    slName = Trim(txtUID.Text)
    If StrComp(slName, "CSI", vbTextCompare) = 0 Then
        slName = "Guide"
    End If
    '6/23/09:  Placed SQL call here and in cmd_Ok instead of Form_Load.
    'rst.MoveFirst
    SQLQuery = "SELECT * FROM ust "
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        slRsName = Trim(rst!ustname)
        If (StrComp(slName, slRsName, 1) = 0) Then
            If Trim(rst!ustpassword) = "" Then
            ' 'Guide' handled differently, so internal guide can login even if guides password is blank.  This method assumes 'guide' username cannot be changed
                If blTestGuide Then
                    If (StrComp(slRsName, "Guide", 1) = 0) Then
                        mTestForBlankPassword = mGetNewPassword(blCallClick)
                        Exit Function
                    End If
               End If
               'general user
                If Not (StrComp(slRsName, "Guide", 1) = 0) Then
                    mTestForBlankPassword = mGetNewPassword(blCallClick)
                    Exit Function
                End If
            End If
            mTestForBlankPassword = False
            Exit Function
        End If
        rst.MoveNext
    Loop
End Function
Private Function mGetNewPassword(blCallClick As Boolean) As Boolean
    sgPassUserName = Trim(txtUID.Text) ' to pass username to form affNewPW
    bmStopLostFocus = True
    AffNewPW.Show vbModal
    bmStopLostFocus = False
    If igExitAff = False Then
        txtPWD.Text = sgPasswordPasser
        mGetNewPassword = True
        If blCallClick Then 'avoid calling click event within click event
            cmdOk_Click
        End If
    Else
        mGetNewPassword = False
    End If
End Function


'           mInitAPIReport - Gather all the filenames from File.ddf.  Required
'           if converting a Btrieve report to ODBC.  If aliases on filenames are
'           used in the report, need to get the real name of the filename to
'           store in the database/tables/location.
'           4-20-04
'
Public Function mInitAPIReport() As Integer
    Dim sFileName As String * 20
    Dim ilUpper As Integer
    Dim ilPos As Integer
    Dim ddf_rst As ADODB.Recordset

    On Error GoTo ErrHand
    
    ReDim tgDDFFileNames(0 To 0) As DDFFILENAMES
    ilUpper = UBound(tgDDFFileNames)
    SQLQuery = "SELECT Xf$Name FROM X$File"
    Set ddf_rst = gSQLSelectCall(SQLQuery)
    
    If Not ddf_rst.EOF Then
        While Not ddf_rst.EOF
            If Mid(ddf_rst(0).Value, 1, 2) <> "X$" Then
                tgDDFFileNames(ilUpper).sLongName = Trim$(ddf_rst(0).Value)
                sFileName = Trim$(tgDDFFileNames(ilUpper).sLongName)
                ilPos = InStr(sFileName, "_")
                If ilPos = 0 Then
                    tgDDFFileNames(ilUpper).sShortName = Trim$(sFileName)
                Else
                    tgDDFFileNames(ilUpper).sShortName = Mid(sFileName, 1, ilPos - 1)
                End If
                ilUpper = ilUpper + 1
                ReDim Preserve tgDDFFileNames(0 To ilUpper)
            End If
            ddf_rst.MoveNext
        Wend
    Else
        If Not igAutoImport And Not igCompelAutoImport Then
            gMsgBox "DDF Open Failed"
        Else
            gLogMsg "DDF Open Failed", "AffErrorLog.txt", False
        End If
    End If
    Exit Function
    
ErrHand:
    gHandleError "", FORMNAME & "-mInitAPIReport"
    Unload frmLogin
End Function

Private Function mConversionCheck()

    'D.S. 12/20/06 Make sure that the conversion program has been run that supports
    'the multicast
    
    Dim ilRet As Integer
    Dim ilIsTntEmpty As Integer
    Dim ilIsShttEmpty As Integer
    
    mConversionCheck = False
    ilRet = True
    On Error GoTo TableDoesNotExist
    'Does the tnt file exist? If so is it empty?
    ilIsTntEmpty = True
    SQLQuery = "Select MAX(tntCode) from tnt"
    Set rst = gSQLSelectCall(SQLQuery)
    If ilRet Then
        'The Tnt file exists, now see if anything is in it
        If IsNull(rst(0).Value) Then
            ilIsTntEmpty = True
            'gMsgBox "Call Counterpoint, a database conversion program needs to be run."
            'mConversionCheck = False
            'Exit Function
        Else
            ilIsTntEmpty = False
        End If
    Else
        'The tnt file does not exist
        If Not igAutoImport And Not igCompelAutoImport Then
            gMsgBox "Call Counterpoint, either DDF reorg needs to be run or the DDF's are out of date."
        Else
            gLogMsg "Call Counterpoint, either DDF reorg needs to be run or the DDF's are out of date.", "AffErorLog.txt", False
        End If
        mConversionCheck = False
        Exit Function
    End If
    
    On Error GoTo ErrHand
    'Check if the shtt has any records
    SQLQuery = "Select MAX(shttCode) from shtt"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ilIsShttEmpty = True
    Else
        ilIsShttEmpty = False
    End If
    
    'There is no problem if both the tnt and shtt are empty, must be a new install
    If ilIsTntEmpty And ilIsShttEmpty Then
        mConversionCheck = True
        Exit Function
    End If
    
    If Not ilIsShttEmpty Then
        ilIsShttEmpty = True
        'We have station records
        SQLQuery = "Select shttPDName from shtt"
        Set rst = gSQLSelectCall(SQLQuery)
        If Trim$(rst(0).Value) = "" Then
            SQLQuery = "Select shttPDName from shtt"
            Set rst = gSQLSelectCall(SQLQuery)
                If Trim$(rst(0).Value) = "" Then
                    SQLQuery = "Select shttTDName from shtt"
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Trim$(rst(0).Value) = "" Then
                        SQLQuery = "Select shttMDName from shtt"
                        Set rst = gSQLSelectCall(SQLQuery)
                        If Trim$(rst(0).Value) = "" Then
                            ilIsShttEmpty = True
                        End If
                    Else
                        ilIsShttEmpty = False
                    End If
                Else
                    ilIsShttEmpty = False
                End If
        Else
            ilIsShttEmpty = False
        End If
    End If
    
    'We have Shtt that has data but the Tnt does not.  Run the mcastconversion.
    If Not ilIsShttEmpty And ilIsTntEmpty Then
        If Not igAutoImport And Not igCompelAutoImport Then
            gMsgBox "Call Counterpoint, a database conversion program needs to be run."
        Else
            gLogMsg "Call Counterpoint, a database conversion program needs to be run.", "AffErorLog.txt", False
        End If
        Exit Function
    End If
    
    If Not ilIsShttEmpty And ilIsTntEmpty Then
        If Not igAutoImport And Not igCompelAutoImport Then
            gMsgBox "Call Counterpoint, either DDF reorg needs to be ran or the DDF's are out of date."
        Else
            gLogMsg "Call Counterpoint, either DDF reorg needs to be ran or the DDF's are out of date.", "AffErorLog.txt", False
        End If
        mConversionCheck = False
        Exit Function
    End If
    mConversionCheck = True
    Exit Function
    
TableDoesNotExist:
    ilRet = False
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mConversionCheck"
End Function

Private Sub mInit()
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slStr As String
    
    bgDevEnv = IsDevEnv()
    slDate = Format$(Now(), "m/d/yy")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slStr = Trim$(Str$(ilValue))
    Do While Len(slStr) < 4
        slStr = "0" & slStr
    Loop
    smSpecial = "Login" & slStr
    igExportSource = 0
    
    bgStationVisible = False
    bgAgreementVisible = False
    bgEMailVisible = False
    bgLogVisible = False
    bgAffidavitVisible = False
    bgPostBuyVisible = False
    bgManagementVisible = False
    bgExportVisible = False
    bgSiteVisible = False
    bgUserVisible = False
    bgRadarVisible = False

End Sub

'Private Sub mCheckShortDateForm()
'    Dim slDate As String
'    Dim ilPos1 As Integer
'    Dim ilPos2 As Integer
'
'    slDate = Format$(Now, "m/d/yy")
'    ilPos1 = InStr(1, slDate, "/", vbTextCompare)
'    ilPos2 = InStr(ilPos1 + 1, slDate, "/", vbTextCompare)
'    If ilPos2 <= 0 Then
'        gMsgBox "Please change your Regional definition of Short Date Format to m/d/yy", vbOK + vbExclamation, "WARNING"
'    Else
'        If ilPos2 + 2 < Len(slDate) Then
'            gMsgBox "Please change your Regional definition of Short Date Format to m/d/yy", vbOK + vbExclamation, "WARNING"
'        End If
'    End If
'End Sub
Private Function mLimitGuide(ByRef rst As Recordset) As Boolean
Dim blGuideLimited As Boolean
Dim blSuccess As Boolean
On Error GoTo ERRORBOX
    bgLimitedGuide = True
    blGuideLimited = mTestIfGuideLimited(rst)
    If Not blGuideLimited Then
        blSuccess = mRecordGuideLimits(rst)
    End If
    If blGuideLimited Or blSuccess Then
        mLimitGuide = True
    Else
        mLimitGuide = False
    End If
    Exit Function
ERRORBOX:
    mLimitGuide = False
'    If Err.Number = 15001 Then
'        Err.Raise 15001, "mTestIfGuideLimited", "Couldn't find Guide"
'    Else
'        Err.Raise 15002, "mRecordGuideLimits", "Couldn't update Guide"
'    End If
End Function
Private Function mTestIfGuideLimited(ByRef rst As Recordset) As Boolean
Dim olField As Field
Dim iltestthisfield As Integer
If StrComp(Trim(rst!ustname), "GUIDE", vbTextCompare) = 0 Then
    For Each olField In rst.Fields
        iltestthisfield = mGuideFieldTest(olField)
        Select Case iltestthisfield
            Case 1             'y or n
                If StrComp(olField.Value, "Y", vbTextCompare) = 0 Then
                    mTestIfGuideLimited = False
                    Exit Function
                End If
            Case 2              'h,i,v
                If StrComp(olField.Value, "I", vbTextCompare) = 0 Then
                    mTestIfGuideLimited = False
                    Exit Function
                End If
        End Select
    Next olField
    mTestIfGuideLimited = True
Else
    Err.Raise 15001
End If
End Function
Private Function mRecordGuideLimits(ByRef rst As Recordset) As Boolean
Dim slUpdateString
On Error GoTo ERRORBOX
    slUpdateString = mUpdateGuideString
    cnn.Execute (slUpdateString)
    mResetGuideGlobals
    mRecordGuideLimits = True
    Exit Function
ERRORBOX:
    Err.Raise 15002
End Function
Private Function mUpdateGuideString() As String
Dim slUpdate As String
    slUpdate = "Update ust set ustwin1 = 'H', ustwin2 = 'H', ustwin3 = 'H', ustwin4 = 'H', ustwin5 = 'H', ustwin6 = 'H', ustwin7 = 'H', ustwin8 = 'H'"
    slUpdate = slUpdate & " ,ustwin9 = 'I', ustwin10 = 'I', ustwin11 = 'H', ustwin12 = 'H', ustwin13 = 'H', ustwin14 = 'N', ustwin15 = 'N', ustwin16 = 'H', ustwin17 = 'H'"
    slUpdate = slUpdate & " ,ustpledge = 'N', ustexptspotalert = 'N', ustexptiscialert = 'N', usttraflogalert = 'N', ustchgExptPriority = 'N', ustExptSpec = 'N' where ustname = 'Guide'"
    mUpdateGuideString = slUpdate
End Function
Private Function mGuideFieldTest(ByRef olField As Field) As Integer
'return 1 for y or n field, 2 for hiv field, 0 for not as tested field
On Error GoTo ERRORBOX
    If InStr(1, olField.Name, "ustWin", vbTextCompare) > 0 And _
        StrComp(olField.Name, "ustWin9", vbTextCompare) <> 0 And StrComp(olField.Name, "ustWin10", vbTextCompare) <> 0 Then
        'ustWin in field, but not with 9 or 10...those fields don't get tested
        If StrComp(olField.Name, "ustWin14", vbTextCompare) = 0 Or StrComp(olField.Name, "ustWin15", vbTextCompare) = 0 Then
            mGuideFieldTest = 1 'win14 and 15 are y or n
        Else
            mGuideFieldTest = 2 'ustwin are usually h, i, or v
        End If
    ElseIf StrComp(olField.Name, "ustPledge", vbTextCompare) = 0 Or StrComp(olField.Name, "ustExptSpotAlert", vbTextCompare) = 0 _
        Or StrComp(olField.Name, "ustExptISCIAlert", vbTextCompare) = 0 Or StrComp(olField.Name, "ustTrafLogAlert", vbTextCompare) = 0 Then
        mGuideFieldTest = 1
    Else
        mGuideFieldTest = 0
    End If
    Exit Function
ERRORBOX:
    mGuideFieldTest = 0
End Function
Private Sub mResetGuideGlobals()
    Dim c As Integer
    For c = 0 To UBound(sgUstWin)
        sgUstWin(c) = "H"
    Next c
    sgUstWin(9) = "I"
    sgUstWin(10) = "I"
    sgUstClear = "N"
    sgUstActivityLog = "H"
    sgUstDelete = "N"
    sgUstPledge = "N"
    sgUstAllowCmmtChg = "N"
    sgUstAllowCmmtDelete = "N"
    sgExptSpotAlert = "N"
    sgExptISCIAlert = "N"
    sgTrafLogAlert = "N"
    sgChgExptPriority = "N"
    sgExptSpec = "N"
    sgChgRptPriority = "N"
End Sub

Private Sub mCreateExportSpec()
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) = STATIONINTERFACE) Then
        ReDim tgSpecInfo(0 To 10) As SPECINFO
        tgSpecInfo(0).sName = "Aff Logs"
        tgSpecInfo(0).sType = "A"
        tgSpecInfo(0).sFullName = "Affiliate Logs"
        tgSpecInfo(0).sCheckDateSpan = "N"
        tgSpecInfo(1).sName = "C & C"
        tgSpecInfo(1).sType = "C"
        tgSpecInfo(1).sFullName = "Clearance and Compensation"
        tgSpecInfo(1).sCheckDateSpan = "N"
        tgSpecInfo(2).sName = "IDC"
        tgSpecInfo(2).sType = "D"
        tgSpecInfo(2).sFullName = "IDC"
        tgSpecInfo(2).sCheckDateSpan = "Y"
        tgSpecInfo(3).sName = "ISCI"
        tgSpecInfo(3).sType = "I"
        tgSpecInfo(3).sFullName = "ISCI"
        tgSpecInfo(3).sCheckDateSpan = "N"
        tgSpecInfo(4).sName = "ISCI C/R"
        tgSpecInfo(4).sType = "R"
        tgSpecInfo(4).sFullName = "ISCI Cross Reference"
        tgSpecInfo(4).sCheckDateSpan = "Y"
        tgSpecInfo(5).sName = "RCS 4"
        tgSpecInfo(5).sType = "4"
        tgSpecInfo(5).sFullName = "RCS 4 Digit Cart #'s"
        tgSpecInfo(5).sCheckDateSpan = "N"
        tgSpecInfo(6).sName = "RCS 5"
        tgSpecInfo(6).sType = "5"
        tgSpecInfo(6).sFullName = "RCS 5 Digit Cart #'s"
        tgSpecInfo(6).sCheckDateSpan = "N"
        tgSpecInfo(7).sName = "StarGd"
        tgSpecInfo(7).sType = "S"
        tgSpecInfo(7).sFullName = "StarGuide"
        tgSpecInfo(7).sCheckDateSpan = "Y"
        tgSpecInfo(8).sName = "Compel"
        tgSpecInfo(8).sType = "W"
        tgSpecInfo(8).sFullName = "Wegener Compel"
        tgSpecInfo(8).sCheckDateSpan = "Y"
        tgSpecInfo(9).sName = "X-Digital"
        tgSpecInfo(9).sType = "X"
        tgSpecInfo(9).sFullName = "X-Digital"
        tgSpecInfo(9).sCheckDateSpan = "Y"
        tgSpecInfo(10).sName = "IPump"
        tgSpecInfo(10).sType = "P"
        tgSpecInfo(10).sFullName = "Wegener IPump"
        tgSpecInfo(10).sCheckDateSpan = "Y"

    Else
        ReDim tgSpecInfo(0 To 9) As SPECINFO
        tgSpecInfo(0).sName = "C & C"
        tgSpecInfo(0).sType = "C"
        tgSpecInfo(0).sFullName = "Clearance and Compensation"
        tgSpecInfo(0).sCheckDateSpan = "N"
        tgSpecInfo(1).sName = "IDC"
        tgSpecInfo(1).sType = "D"
        tgSpecInfo(1).sFullName = "IDC"
        tgSpecInfo(1).sCheckDateSpan = "Y"
        tgSpecInfo(2).sName = "ISCI"
        tgSpecInfo(2).sType = "I"
        tgSpecInfo(2).sFullName = "ISCI"
        tgSpecInfo(2).sCheckDateSpan = "N"
        tgSpecInfo(3).sName = "ISCI C/R"
        tgSpecInfo(3).sType = "R"
        tgSpecInfo(3).sFullName = "ISCI Cross Reference"
        tgSpecInfo(3).sCheckDateSpan = "Y"
        tgSpecInfo(4).sName = "RCS 4"
        tgSpecInfo(4).sType = "4"
        tgSpecInfo(4).sFullName = "RCS 4 Digit Cart #'s"
        tgSpecInfo(4).sCheckDateSpan = "N"
        tgSpecInfo(5).sName = "RCS 5"
        tgSpecInfo(5).sType = "5"
        tgSpecInfo(5).sFullName = "RCS 5 Digit Cart #'s"
        tgSpecInfo(5).sCheckDateSpan = "N"
        tgSpecInfo(6).sName = "StarGd"
        tgSpecInfo(6).sType = "S"
        tgSpecInfo(6).sFullName = "StarGuide"
        tgSpecInfo(6).sCheckDateSpan = "Y"
        tgSpecInfo(7).sName = "Compel"
        tgSpecInfo(7).sType = "W"
        tgSpecInfo(7).sFullName = "Wegener"
        tgSpecInfo(7).sCheckDateSpan = "Y"
        tgSpecInfo(8).sName = "X-Digital"
        tgSpecInfo(8).sType = "X"
        tgSpecInfo(8).sFullName = "X-Digital"
        tgSpecInfo(8).sCheckDateSpan = "Y"
        tgSpecInfo(9).sName = "IPump"
        tgSpecInfo(9).sType = "P"
        tgSpecInfo(9).sFullName = "Wegener IPump"
        tgSpecInfo(9).sCheckDateSpan = "Y"
    End If
End Sub

Private Sub mCreateStatustype()
    'Agreement only shows status- 1:; 2:; 5: and 9:
    'All other screens show all the status
    tgStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
    tgStatusTypes(0).iPledged = 0
    tgStatusTypes(0).iStatus = 0
    tgStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(1).iPledged = 1
    tgStatusTypes(1).iStatus = 1
    tgStatusTypes(2).sName = "3-Not Aired Tech Diff"
    tgStatusTypes(2).iPledged = 2
    tgStatusTypes(2).iStatus = 2
    tgStatusTypes(3).sName = "4-Not Aired Blackout"
    tgStatusTypes(3).iPledged = 2
    tgStatusTypes(3).iStatus = 3
    tgStatusTypes(4).sName = "5-Not Aired Other"
    tgStatusTypes(4).iPledged = 2
    tgStatusTypes(4).iStatus = 4
    tgStatusTypes(5).sName = "6-Not Aired Product"
    tgStatusTypes(5).iPledged = 2
    tgStatusTypes(5).iStatus = 5
    tgStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
    tgStatusTypes(6).iPledged = 3
    tgStatusTypes(6).iStatus = 6
    tgStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
    tgStatusTypes(7).iPledged = 3
    tgStatusTypes(7).iStatus = 7
    'D.S. 11/6/08 remove the "or Aired" from the status 9 description
    'Affiliate Meeting Decisions item 5) f-iv
    'tgStatusTypes(8).sName = "9-Not Carried or Aired"
    tgStatusTypes(8).sName = "9-Not Carried"
    tgStatusTypes(8).iPledged = 2
    tgStatusTypes(8).iStatus = 8
    tgStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(9).iPledged = 1
    tgStatusTypes(9).iStatus = 9
    tgStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
    tgStatusTypes(10).iPledged = 1
    tgStatusTypes(10).iStatus = 10
    tgStatusTypes(ASTEXTENDED_MG).sName = "MG"
    tgStatusTypes(ASTEXTENDED_MG).iPledged = 3
    tgStatusTypes(ASTEXTENDED_MG).iStatus = ASTEXTENDED_MG
    tgStatusTypes(ASTEXTENDED_BONUS).sName = "Bonus"
    tgStatusTypes(ASTEXTENDED_BONUS).iPledged = 3
    tgStatusTypes(ASTEXTENDED_BONUS).iStatus = ASTEXTENDED_BONUS
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).sName = "Replacement"
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iPledged = 3
    tgStatusTypes(ASTEXTENDED_REPLACEMENT).iStatus = ASTEXTENDED_REPLACEMENT
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub

Private Function mCheckVersion() As Integer
    Dim ilRet As Integer
    Dim hlVersion As Integer
    Dim slVersion As String
    Dim ilPos As Integer
    Dim slChar As String
    Dim slLine As String
    Dim slTemp As String

    slVersion = App.Major & "." & App.Minor
    'ilRet = 0
    'On Error GoTo mWrongVersion:
    'hlVersion = FreeFile
    'Open sgDBPath & "Version.Csi" For Input Access Read As hlVersion
    slTemp = sgDBPath & "Version.Csi"
    ilRet = gFileOpen(slTemp, "Input Access Read", hlVersion)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Version.Csi missing from " & sgDBPath, vbOKOnly + vbCritical, "Error"
        mCheckVersion = False
        Exit Function
    End If
    On Error GoTo mWrongVersion:
    Line Input #hlVersion, slLine
    On Error GoTo 0
    Close hlVersion
    If (Asc(slLine) = 26) Or (ilRet <> 0) Or (StrComp(slVersion, slLine, 1) <> 0) Then    'Ctrl Z
        Screen.MousePointer = vbDefault
        MsgBox "Programs and Database versions don't match", vbOKOnly + vbCritical, "Error"
        mCheckVersion = False
        Exit Function
    End If
    mCheckVersion = True
    Exit Function
mWrongVersion:
    On Error GoTo 0
    ilRet = 1
    Resume Next
End Function

Private Sub mCheckSecurity()
    Dim slName As String
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim slStr As String
    Dim llNow As Long
    Dim llDate As Long
    
    igGGFlag = -1
    igRptGGFlag = 1
    If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
        Exit Sub
    End If
    If bgInternalGuide Then
        Exit Sub
    End If
    SQLQuery = "Select safName From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        slName = Trim$(rst!safName)
        ilField1 = Asc(slName)
        slStr = Mid$(slName, 2, 5)
        llDate = Val(slStr)
        llNow = gDateValue(Format$(gNow(), "m/d/yy"))
        ilField2 = Asc(Mid$(slName, 11, 1))
        If (ilField1 = 0) And (ilField2 = 1) Then
            If llDate <= llNow Then
                ilField2 = 0
            End If
        End If
        If (ilField1 = 0) And (ilField2 = 0) Then
            mResetGuideGlobals
            sgUstWin(9) = "H"
            sgUstWin(10) = "H"
            sgSpfUseCartNo = "Y"
            sgSpfRemoteUsers = "N"
            sgSpfUsingFeatures2 = Chr$(0)
            sgSpfUsingFeatures5 = Chr$(0)
            sgSpfUsingFeatures9 = Chr$(0)
            sgSpfSportInfo = Chr$(0)
            sgSpfUseProdSptScr = "A"
            gUsingWeb = False
            gUsingUnivision = False
            gISCIExport = False
            sgRCSExportCart4 = "N"
            sgRCSExportCart5 = "N"
            gUsingXDigital = False
            gWegenerExport = False
            gOLAExport = False
            gWegenerExport = False
            igGGFlag = 0
            gSetRptGGFlag slName
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "mCheckSecurity"
    Exit Sub

End Sub

Private Sub mLogActivityFileName()
    '12/4/12: Routine added to determine if activity to a file should be captured
    Dim hlLogActivityFileName As Integer
    Dim slSection As String
    sgLogActivityFileName = ""
    sgLogActivityInto = ""
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    If gLoadOption(slSection, "LogActivityFileName", sgLogActivityFileName) Then
        sgLogActivityFileName = UCase(Trim$(sgLogActivityFileName))
        sgLogActivityInto = "LogActivityFor_" & sgLogActivityFileName & ".txt"
        gLogMsgWODT "OD", hlLogActivityFileName, sgMsgDirectory & sgLogActivityInto
        gLogMsgWODT "C", hlLogActivityFileName, ""
    End If
End Sub

Private Sub mLimitRemoteExportUsers()
    If bgRemoteExport <> True Then
        Exit Sub
    End If
    sgUstWin(0) = "H"
    'sgUstWin(1) = rst!ustWin1   'Station
    sgUstWin(2) = "H"
    sgUstWin(3) = "H"
    sgUstWin(4) = "H"
    sgUstWin(5) = "H"
    sgUstWin(6) = "H"
    sgUstWin(7) = "H"
    sgUstWin(8) = "H"
    'sgUstWin(9) = rst!ustWin9  'Option
    'sgUstWin(10) = rst!ustWin10 'Site
    sgUstWin(11) = "H"
    sgUstWin(12) = "H"
    sgUstWin(13) = "H"
    sgUstWin(14) = "H"
    sgUstClear = rst!ustWin14
    sgUstActivityLog = rst!ustActivityLog
    sgUstDelete = rst!ustWin15
    sgUstPledge = rst!ustPledge
    sgUstAllowCmmtChg = rst!ustAllowCmmtChg
    sgUstAllowCmmtDelete = rst!ustAllowCmmtDelete
    sgExptSpotAlert = rst!ustExptSpotAlert
    sgExptISCIAlert = rst!ustExptISCIAlert
    sgTrafLogAlert = rst!ustTrafLogAlert
    sgPhoneNo = rst!ustPhoneNo
    igUstSSMnfCode = rst!ustSSMnfCode
    sgCity = rst!ustCity
    sgAllowedToBlock = rst!ustAllowedToBlock
    sgEMail = ""
    lgEMailCefCode = rst!ustEmailcefcode
    sgChgExptPriority = rst!ustChgExptPriority
    sgExptSpec = rst!ustExptSpec
    sgChgRptPriority = rst!ustChgRptPriority
    
    gUsingWeb = False
    gUsingUnivision = False
    gISCIExport = False
    sgRCSExportCart4 = "N"
    sgRCSExportCart5 = "N"
    gUsingXDigital = False
    gWegenerExport = False
    gOLAExport = False
    gWegenerExport = False
    
End Sub
