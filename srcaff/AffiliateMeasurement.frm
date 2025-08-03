VERSION 5.00
Begin VB.Form AffiliateMeasurement 
   Caption         =   "Affiliate Measurement"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "AffiliateMeasurement.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   7845
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7215
      Top             =   2205
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6630
      Top             =   2445
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6105
      Top             =   2580
   End
   Begin VB.CommandButton cmcGenerate 
      Caption         =   "Generate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1770
      TabIndex        =   0
      Top             =   2535
      Width           =   1575
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4455
      TabIndex        =   1
      Top             =   2535
      Width           =   1575
   End
   Begin VB.Label lacBuilding 
      Height          =   285
      Index           =   1
      Left            =   4590
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label lacClearing 
      Height          =   285
      Index           =   1
      Left            =   4590
      TabIndex        =   7
      Top             =   1095
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label lacBuilding 
      Caption         =   "Step 3: Building First Week from Agreements................"
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lacClearing 
      Caption         =   "Step 2: Clear First Weeks.............................................."
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1095
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lacAging 
      Height          =   285
      Index           =   1
      Left            =   4590
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label lacAging 
      Caption         =   "Step 1: Aging Weeks....................................................."
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   660
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lacDates 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   450
      TabIndex        =   2
      Top             =   135
      Width           =   6345
   End
End
Attribute VB_Name = "AffiliateMeasurement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'Grid Controls
Private bmAutoMode As Boolean

Private imGuideUstCode As Integer
Private bmNoPervasive As Boolean

Private imGenerating As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer
'Private lmSmtCode As Long
Private smNewestWeekDate As String
Private smSuNewestWeekDate As String
Private smOldestWeekDate As String
Private smClientAbbr As String


Private lm1970 As Long

Private Const FORMNAME As String = "AffiliateMeasurement"
Private smt_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private cptt_rst As ADODB.Recordset
Private webl_rst As ADODB.Recordset


Private Sub cmcGenerate_Click()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If imGenerating = True Then
        Exit Sub
    End If
    imGenerating = True
    Screen.MousePointer = vbHourglass
    If bmAutoMode Then
        gUpdateTaskMonitor 1, "AMB"
    End If
    ilRet = mGenerate()
    imGenerating = False
    If ilRet Then
        cmcCancel.Caption = "&Done"
    End If
    Screen.MousePointer = vbDefault
    If bmAutoMode Then
        gUpdateTaskMonitor 2, "AMB"
        cmcCancel_Click
    End If
    imTerminate = False
    Exit Sub
cmcSaveErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement: cmcGenerate_Click"
End Sub

Private Sub cmcCancel_Click()
    If imGenerating Then
        imTerminate = True
        Exit Sub
    End If
    Unload AffiliateMeasurement
End Sub

Private Sub Form_Activate()
    If imFirstTime Then
        imFirstTime = False
    End If
End Sub

Private Sub Form_GotFocus()
    cmcGenerate.Caption = "&Generate"
    cmcCancel.Caption = "&Cancel"
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim ilRet As Integer
        
    Screen.MousePointer = vbHourglass
    imTerminate = False
    imGenerating = False
    imFirstTime = True
    gCenterStdAlone Me
    mInit
    If InStr(1, sgCommand, "/UserInput", vbTextCompare) > 0 Then
        bmAutoMode = False
        Me.WindowState = vbNormal
    Else
        bmAutoMode = True
        Me.WindowState = vbMinimized
    End If
    tmcStart.Enabled = True
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imGenerating Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    tmcSetTime.Enabled = False
    
    smt_rst.Close
    att_rst.Close
    ast_rst.Close
    cptt_rst.Close
    webl_rst.Close
    
    cnn.Close
    
    btrStopAppl
    Set AffiliateMeasurement = Nothing   'Remove data segment
    End
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "AMB"
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    Screen.MousePointer = vbHourglass
    mInitAffiliateMeasurement
    cmcGenerate.Enabled = True
    DoEvents
    Screen.MousePointer = vbDefault
    If bmAutoMode Then
        DoEvents
        sgTimeZone = Left$(gGetLocalTZName(), 1)
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
        igShowMsgBox = False
        cmcGenerate_Click
    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload AffiliateMeasurement
End Sub


Private Sub mInit()
    Dim sBuffer As String
    Dim ilRet As Integer
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
    ReDim sWin(0 To 13) As String * 1
    '5/11/11
    Dim blAddGuide As Boolean
    'dan 2/23/12 can't have error handler in error handler
    Dim blNeedToCloseCnn As Boolean
    
    sgCommand = Command$
    blNeedToCloseCnn = False
    igShowMsgBox = True
    
    'igDemoMode = False
    'If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    'End If
    
    'Used to speed-up testing exports with multiple files reduce record count needed to create a new file
    igSmallFiles = False
    If InStr(sgCommand, "SmallFiles") Then
        igSmallFiles = True
    End If
    
    igAutoImport = False
    slStartIn = CurDir$
    sgCurDir = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommand, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
        
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
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
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
    
    If Not gLoadOption("Locations", "Logo", sgLogoPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini [Database] 'Name' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    End If
    
        
    'Import is optional
    If gLoadOption("Locations", "Import", sgImportDirectory) Then
        sgImportDirectory = gSetPathEndSlash(sgImportDirectory, True)
    Else
        sgImportDirectory = ""
    End If
    
    If gLoadOption("Locations", "ContractPDF", sgContractPDFPath) Then
        sgContractPDFPath = gSetPathEndSlash(sgContractPDFPath, True)
    Else
        sgContractPDFPath = ""
    End If
    
    
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
    If gLoadOption("Locations", "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Locations", "Wallpaper", sgWallpaper)
    
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    If Not gLoadOption("Locations", "DBPath", sgDBPath) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is Missing.", vbCritical
        Unload AffiliateMeasurement
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is Missing.", vbCritical
        Unload AffiliateMeasurement
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
    'Code modified for testing
    
    
    If Not mOpenPervasiveAPI Then
        Unload AffiliateMeasurement
        Exit Sub
    End If
    
    
    'Test for Guide- if not added- add
    'SQLQuery = "Select MAX(ustCode) from ust"
    'Set rst = cnn.Execute(SQLQuery)
    ''If rst(0).Value = 0 Then
    'If IsNull(rst(0).Value) Then
    ''5/11/11
    '    blAddGuide = True
    'Else
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
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
        'mResetGuideGlobals
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
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        cnn.BeginTrans
        blNeedToCloseCnn = True
        'cnn.ConnectionTimeout = 30  ' Increase from the default of 15 to 30 seconds.
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "", FORMNAME & "-Form_Load"
            bmNoPervasive = True
            On Error Resume Next
            If blNeedToCloseCnn Then
                cnn.RollbackTrans
            End If
            tmcTerminate.Enabled = True
            Exit Sub
        End If
        cnn.CommitTrans
        blNeedToCloseCnn = False
        SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
        Set rst = cnn.Execute(SQLQuery)
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
    ' Dan M added spfusingFeatures2
    SQLQuery = "SELECT spfGClient, spfGAlertInterval, spfGUseAffSys, spfUsingFeatures7, spfUsingFeatures2, spfUsingFeatures8"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = cnn.Execute(SQLQuery)
    
    If Not rst.EOF Then
        If UCase(rst!spfGUseAffSys) <> "Y" Then
            gMsgBox "The Affiliate system has not been activated.  Please call Counterpoint.", vbCritical
            Unload AffiliateMeasurement
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
    End If
    
    If Not rst.EOF Then
        sgClientName = Trim$(rst!spfGClient)
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
            sgNowDate = "12/15/1999"
        End If
    End If


    ilRet = gInitGlobals()
    If ilRet = 0 Then
        'While Not gVerifyWebIniSettings()
        '    frmWebIniOptions.Show vbModal
        '    If Not igWebIniOptionsOK Then
        '        Unload AffiliateMeasurement
        '        Exit Sub
        '    End If
        'Wend
    End If
    
    Call gLoadOption("Database", "AutoLogin", sAutoLogin)
    
    
    On Error GoTo ErrHand
    'If Not igAutoImport Then
    '    ilRet = mInitAPIReport()      '4-19-04
    'End If
    
    
    ilRet = gTestWebVersion()
    'Move report logo to local C drice (c:\csi\rptlogo.bmp)
    ilRet = 0
    On Error GoTo mStartUpErr:
    'slDateTime1 = FileDateTime("C:\CSI\RptLogo.Bmp")
    'If ilRet <> 0 Then
    '    ilRet = 0
    '    MkDir "C:\CSI"
    '    If ilRet = 0 Then
    '        FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '    Else
    '        FileCopy sgDBPath & "RptLogo.Bmp", sgLogoPath & "RptLogo.Bmp"
    '    End If
    'Else
    '    ilRet = 0
    '    slDateTime2 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '    If ilRet = 0 Then
    '        If StrComp(slDateTime1, slDateTime2, 1) <> 0 Then
    '            FileCopy sgLogoPath & "RptLogo.Bmp", "C:\CSI\RptLogo.Bmp"
    '        End If
    '    End If
    'End If
     'ttp 5260
    'If Dir(sgLogoPath & "RptLogo.jpg") > "" Then
    '    If Dir("c:\csi\RptLogo.jpg") = "" Then
    '        FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '    'ok, both exist.  is logopath's more recent?
    '    Else
    '        slDateTime1 = FileDateTime(sgLogoPath & "RptLogo.Bmp")
    '        slDateTime2 = FileDateTime("C:\CSI\RptLogo.jpg")
    '        If StrComp(slDateTime1, slDateTime2, vbBinaryCompare) <> 0 Then
     '           FileCopy sgLogoPath & "RptLogo.jpg", "C:\csi\RptLogo.jpg"
    '        End If
    '    End If
    'End If
    'Determine number if X-Digital HeadEnds
    ReDim sgXDSSection(0 To 0) As String
    'slXMLINIInputFile = gXmlIniPath(True)
    'If LenB(slXMLINIInputFile) <> 0 Then
    '    ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    'End If
    'Test to see if this function has been ran before, if so don't run it again
    igEmailNeedsConv = False
    mCreateStatustype
    ilRet = gPopMarkets()
    ilRet = gPopMSAMarkets()         'MSA markets
    ilRet = gPopMntInfo("T", tgTerritoryInfo())
    ilRet = gPopMntInfo("C", tgCityInfo())
    ilRet = gPopOwnerNames()
    ilRet = gPopStations()
    ilRet = gPopVehicleOptions()
    ilRet = gPopVehicles()
    ilRet = gPopSellingVehicles()
    ilRet = gPopAdvertisers()
    ilRet = gPopReportNames()
    ilRet = gGetLatestRatecard()
    ilRet = gPopTimeZones()
    ilRet = gPopStates()
    ilRet = gPopFormats()
    ilRet = gPopAvailNames()
    ilRet = gPopMediaCodes()
    
    Exit Sub

mStartUpErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
'    gMsg = ""
'    For Each gErrSQL In cnn.Errors
'        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
'            gMsg = "A SQL error has occured: "
'            gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
'        End If
'    Next gErrSQL
'    On Error Resume Next
'    cnn.RollbackTrans
'    On Error GoTo 0
'    If gMsg = "" Then
'        gMsgBox "Error at Start-up " & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
'    End If
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
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"          '4-13-17 Missed-mg bypassed
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub

Private Function mAddSmt(ilVefCode As Integer, ilShttCode As Integer) As Integer
    'Dim llSmtCode As Long
    Dim ilMktRepUstCode As Integer
    Dim ilServRepUstCode As Integer
    
    On Error GoTo ErrHand
    mAddSmt = False
    If Not mGetReps(ilShttCode, ilMktRepUstCode, ilServRepUstCode) Then
        ilMktRepUstCode = 0
        ilServRepUstCode = 0
    End If
    SQLQuery = "Insert Into smt ( "
    'SQLQuery = SQLQuery & "smtCode, "
    SQLQuery = SQLQuery & "smtWk1StartDate, "
    SQLQuery = SQLQuery & "smtClientAbbr, "
    SQLQuery = SQLQuery & "smtVefCode, "
    SQLQuery = SQLQuery & "smtShttCode, "
    SQLQuery = SQLQuery & "smtMktRepUstCode, "
    SQLQuery = SQLQuery & "smtServRepUstCode, "
    SQLQuery = SQLQuery & "smtGenDate, "
    SQLQuery = SQLQuery & "smtWeeksAired1, "
    SQLQuery = SQLQuery & "smtWeeksAired2, "
    SQLQuery = SQLQuery & "smtWeeksAired3, "
    SQLQuery = SQLQuery & "smtWeeksAired4, "
    SQLQuery = SQLQuery & "smtWeeksAired5, "
    SQLQuery = SQLQuery & "smtWeeksAired6, "
    SQLQuery = SQLQuery & "smtWeeksAired7, "
    SQLQuery = SQLQuery & "smtWeeksAired8, "
    SQLQuery = SQLQuery & "smtWeeksAired9, "
    SQLQuery = SQLQuery & "smtWeeksAired10, "
    SQLQuery = SQLQuery & "smtWeeksAired11, "
    SQLQuery = SQLQuery & "smtWeeksAired12, "
    SQLQuery = SQLQuery & "smtWeeksAired13, "
    SQLQuery = SQLQuery & "smtWeeksAired14, "
    SQLQuery = SQLQuery & "smtWeeksAired15, "
    SQLQuery = SQLQuery & "smtWeeksAired16, "
    SQLQuery = SQLQuery & "smtWeeksAired17, "
    SQLQuery = SQLQuery & "smtWeeksAired18, "
    SQLQuery = SQLQuery & "smtWeeksAired19, "
    SQLQuery = SQLQuery & "smtWeeksAired20, "
    SQLQuery = SQLQuery & "smtWeeksAired21, "
    SQLQuery = SQLQuery & "smtWeeksAired22, "
    SQLQuery = SQLQuery & "smtWeeksAired23, "
    SQLQuery = SQLQuery & "smtWeeksAired24, "
    SQLQuery = SQLQuery & "smtWeeksAired25, "
    SQLQuery = SQLQuery & "smtWeeksAired26, "
    SQLQuery = SQLQuery & "smtWeeksAired27, "
    SQLQuery = SQLQuery & "smtWeeksAired28, "
    SQLQuery = SQLQuery & "smtWeeksAired29, "
    SQLQuery = SQLQuery & "smtWeeksAired30, "
    SQLQuery = SQLQuery & "smtWeeksAired31, "
    SQLQuery = SQLQuery & "smtWeeksAired32, "
    SQLQuery = SQLQuery & "smtWeeksAired33, "
    SQLQuery = SQLQuery & "smtWeeksAired34, "
    SQLQuery = SQLQuery & "smtWeeksAired35, "
    SQLQuery = SQLQuery & "smtWeeksAired36, "
    SQLQuery = SQLQuery & "smtWeeksAired37, "
    SQLQuery = SQLQuery & "smtWeeksAired38, "
    SQLQuery = SQLQuery & "smtWeeksAired39, "
    SQLQuery = SQLQuery & "smtWeeksAired40, "
    SQLQuery = SQLQuery & "smtWeeksAired41, "
    SQLQuery = SQLQuery & "smtWeeksAired42, "
    SQLQuery = SQLQuery & "smtWeeksAired43, "
    SQLQuery = SQLQuery & "smtWeeksAired44, "
    SQLQuery = SQLQuery & "smtWeeksAired45, "
    SQLQuery = SQLQuery & "smtWeeksAired46, "
    SQLQuery = SQLQuery & "smtWeeksAired47, "
    SQLQuery = SQLQuery & "smtWeeksAired48, "
    SQLQuery = SQLQuery & "smtWeeksAired49, "
    SQLQuery = SQLQuery & "smtWeeksAired50, "
    SQLQuery = SQLQuery & "smtWeeksAired51, "
    SQLQuery = SQLQuery & "smtWeeksAired52, "
    SQLQuery = SQLQuery & "smtWeeksMissing1, "
    SQLQuery = SQLQuery & "smtWeeksMissing2, "
    SQLQuery = SQLQuery & "smtWeeksMissing3, "
    SQLQuery = SQLQuery & "smtWeeksMissing4, "
    SQLQuery = SQLQuery & "smtWeeksMissing5, "
    SQLQuery = SQLQuery & "smtWeeksMissing6, "
    SQLQuery = SQLQuery & "smtWeeksMissing7, "
    SQLQuery = SQLQuery & "smtWeeksMissing8, "
    SQLQuery = SQLQuery & "smtWeeksMissing9, "
    SQLQuery = SQLQuery & "smtWeeksMissing10, "
    SQLQuery = SQLQuery & "smtWeeksMissing11, "
    SQLQuery = SQLQuery & "smtWeeksMissing12, "
    SQLQuery = SQLQuery & "smtWeeksMissing13, "
    SQLQuery = SQLQuery & "smtWeeksMissing14, "
    SQLQuery = SQLQuery & "smtWeeksMissing15, "
    SQLQuery = SQLQuery & "smtWeeksMissing16, "
    SQLQuery = SQLQuery & "smtWeeksMissing17, "
    SQLQuery = SQLQuery & "smtWeeksMissing18, "
    SQLQuery = SQLQuery & "smtWeeksMissing19, "
    SQLQuery = SQLQuery & "smtWeeksMissing20, "
    SQLQuery = SQLQuery & "smtWeeksMissing21, "
    SQLQuery = SQLQuery & "smtWeeksMissing22, "
    SQLQuery = SQLQuery & "smtWeeksMissing23, "
    SQLQuery = SQLQuery & "smtWeeksMissing24, "
    SQLQuery = SQLQuery & "smtWeeksMissing25, "
    SQLQuery = SQLQuery & "smtWeeksMissing26, "
    SQLQuery = SQLQuery & "smtWeeksMissing27, "
    SQLQuery = SQLQuery & "smtWeeksMissing28, "
    SQLQuery = SQLQuery & "smtWeeksMissing29, "
    SQLQuery = SQLQuery & "smtWeeksMissing30, "
    SQLQuery = SQLQuery & "smtWeeksMissing31, "
    SQLQuery = SQLQuery & "smtWeeksMissing32, "
    SQLQuery = SQLQuery & "smtWeeksMissing33, "
    SQLQuery = SQLQuery & "smtWeeksMissing34, "
    SQLQuery = SQLQuery & "smtWeeksMissing35, "
    SQLQuery = SQLQuery & "smtWeeksMissing36, "
    SQLQuery = SQLQuery & "smtWeeksMissing37, "
    SQLQuery = SQLQuery & "smtWeeksMissing38, "
    SQLQuery = SQLQuery & "smtWeeksMissing39, "
    SQLQuery = SQLQuery & "smtWeeksMissing40, "
    SQLQuery = SQLQuery & "smtWeeksMissing41, "
    SQLQuery = SQLQuery & "smtWeeksMissing42, "
    SQLQuery = SQLQuery & "smtWeeksMissing43, "
    SQLQuery = SQLQuery & "smtWeeksMissing44, "
    SQLQuery = SQLQuery & "smtWeeksMissing45, "
    SQLQuery = SQLQuery & "smtWeeksMissing46, "
    SQLQuery = SQLQuery & "smtWeeksMissing47, "
    SQLQuery = SQLQuery & "smtWeeksMissing48, "
    SQLQuery = SQLQuery & "smtWeeksMissing49, "
    SQLQuery = SQLQuery & "smtWeeksMissing50, "
    SQLQuery = SQLQuery & "smtWeeksMissing51, "
    SQLQuery = SQLQuery & "smtWeeksMissing52, "
    SQLQuery = SQLQuery & "smtSpotPosted1, "
    SQLQuery = SQLQuery & "smtSpotPosted2, "
    SQLQuery = SQLQuery & "smtSpotPosted3, "
    SQLQuery = SQLQuery & "smtSpotPosted4, "
    SQLQuery = SQLQuery & "smtSpotPosted5, "
    SQLQuery = SQLQuery & "smtSpotPosted6, "
    SQLQuery = SQLQuery & "smtSpotPosted7, "
    SQLQuery = SQLQuery & "smtSpotPosted8, "
    SQLQuery = SQLQuery & "smtSpotPosted9, "
    SQLQuery = SQLQuery & "smtSpotPosted10, "
    SQLQuery = SQLQuery & "smtSpotPosted11, "
    SQLQuery = SQLQuery & "smtSpotPosted12, "
    SQLQuery = SQLQuery & "smtSpotPosted13, "
    SQLQuery = SQLQuery & "smtSpotPosted14, "
    SQLQuery = SQLQuery & "smtSpotPosted15, "
    SQLQuery = SQLQuery & "smtSpotPosted16, "
    SQLQuery = SQLQuery & "smtSpotPosted17, "
    SQLQuery = SQLQuery & "smtSpotPosted18, "
    SQLQuery = SQLQuery & "smtSpotPosted19, "
    SQLQuery = SQLQuery & "smtSpotPosted20, "
    SQLQuery = SQLQuery & "smtSpotPosted21, "
    SQLQuery = SQLQuery & "smtSpotPosted22, "
    SQLQuery = SQLQuery & "smtSpotPosted23, "
    SQLQuery = SQLQuery & "smtSpotPosted24, "
    SQLQuery = SQLQuery & "smtSpotPosted25, "
    SQLQuery = SQLQuery & "smtSpotPosted26, "
    SQLQuery = SQLQuery & "smtSpotPosted27, "
    SQLQuery = SQLQuery & "smtSpotPosted28, "
    SQLQuery = SQLQuery & "smtSpotPosted29, "
    SQLQuery = SQLQuery & "smtSpotPosted30, "
    SQLQuery = SQLQuery & "smtSpotPosted31, "
    SQLQuery = SQLQuery & "smtSpotPosted32, "
    SQLQuery = SQLQuery & "smtSpotPosted33, "
    SQLQuery = SQLQuery & "smtSpotPosted34, "
    SQLQuery = SQLQuery & "smtSpotPosted35, "
    SQLQuery = SQLQuery & "smtSpotPosted36, "
    SQLQuery = SQLQuery & "smtSpotPosted37, "
    SQLQuery = SQLQuery & "smtSpotPosted38, "
    SQLQuery = SQLQuery & "smtSpotPosted39, "
    SQLQuery = SQLQuery & "smtSpotPosted40, "
    SQLQuery = SQLQuery & "smtSpotPosted41, "
    SQLQuery = SQLQuery & "smtSpotPosted42, "
    SQLQuery = SQLQuery & "smtSpotPosted43, "
    SQLQuery = SQLQuery & "smtSpotPosted44, "
    SQLQuery = SQLQuery & "smtSpotPosted45, "
    SQLQuery = SQLQuery & "smtSpotPosted46, "
    SQLQuery = SQLQuery & "smtSpotPosted47, "
    SQLQuery = SQLQuery & "smtSpotPosted48, "
    SQLQuery = SQLQuery & "smtSpotPosted49, "
    SQLQuery = SQLQuery & "smtSpotPosted50, "
    SQLQuery = SQLQuery & "smtSpotPosted51, "
    SQLQuery = SQLQuery & "smtSpotPosted52, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC1, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC2, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC3, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC4, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC5, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC6, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC7, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC8, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC9, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC10, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC11, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC12, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC13, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC14, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC15, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC16, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC17, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC18, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC19, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC20, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC21, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC22, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC23, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC24, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC25, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC26, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC27, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC28, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC29, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC30, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC31, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC32, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC33, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC34, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC35, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC36, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC37, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC38, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC39, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC40, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC41, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC42, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC43, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC44, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC45, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC46, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC47, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC48, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC49, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC50, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC51, "
    SQLQuery = SQLQuery & "smtSpotPostedSNC52, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC1, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC2, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC3, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC4, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC5, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC6, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC7, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC8, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC9, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC10, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC11, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC12, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC13, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC14, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC15, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC16, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC17, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC18, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC19, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC20, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC21, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC22, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC23, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC24, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC25, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC26, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC27, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC28, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC29, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC30, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC31, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC32, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC33, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC34, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC35, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC36, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC37, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC38, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC39, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC40, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC41, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC42, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC43, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC44, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC45, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC46, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC47, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC48, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC49, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC50, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC51, "
    SQLQuery = SQLQuery & "smtSpotPostedNNC52, "
    SQLQuery = SQLQuery & "smtDaysSubmitted1, "
    SQLQuery = SQLQuery & "smtDaysSubmitted2, "
    SQLQuery = SQLQuery & "smtDaysSubmitted3, "
    SQLQuery = SQLQuery & "smtDaysSubmitted4, "
    SQLQuery = SQLQuery & "smtDaysSubmitted5, "
    SQLQuery = SQLQuery & "smtDaysSubmitted6, "
    SQLQuery = SQLQuery & "smtDaysSubmitted7, "
    SQLQuery = SQLQuery & "smtDaysSubmitted8, "
    SQLQuery = SQLQuery & "smtDaysSubmitted9, "
    SQLQuery = SQLQuery & "smtDaysSubmitted10, "
    SQLQuery = SQLQuery & "smtDaysSubmitted11, "
    SQLQuery = SQLQuery & "smtDaysSubmitted12, "
    SQLQuery = SQLQuery & "smtDaysSubmitted13, "
    SQLQuery = SQLQuery & "smtDaysSubmitted14, "
    SQLQuery = SQLQuery & "smtDaysSubmitted15, "
    SQLQuery = SQLQuery & "smtDaysSubmitted16, "
    SQLQuery = SQLQuery & "smtDaysSubmitted17, "
    SQLQuery = SQLQuery & "smtDaysSubmitted18, "
    SQLQuery = SQLQuery & "smtDaysSubmitted19, "
    SQLQuery = SQLQuery & "smtDaysSubmitted20, "
    SQLQuery = SQLQuery & "smtDaysSubmitted21, "
    SQLQuery = SQLQuery & "smtDaysSubmitted22, "
    SQLQuery = SQLQuery & "smtDaysSubmitted23, "
    SQLQuery = SQLQuery & "smtDaysSubmitted24, "
    SQLQuery = SQLQuery & "smtDaysSubmitted25, "
    SQLQuery = SQLQuery & "smtDaysSubmitted26, "
    SQLQuery = SQLQuery & "smtDaysSubmitted27, "
    SQLQuery = SQLQuery & "smtDaysSubmitted28, "
    SQLQuery = SQLQuery & "smtDaysSubmitted29, "
    SQLQuery = SQLQuery & "smtDaysSubmitted30, "
    SQLQuery = SQLQuery & "smtDaysSubmitted31, "
    SQLQuery = SQLQuery & "smtDaysSubmitted32, "
    SQLQuery = SQLQuery & "smtDaysSubmitted33, "
    SQLQuery = SQLQuery & "smtDaysSubmitted34, "
    SQLQuery = SQLQuery & "smtDaysSubmitted35, "
    SQLQuery = SQLQuery & "smtDaysSubmitted36, "
    SQLQuery = SQLQuery & "smtDaysSubmitted37, "
    SQLQuery = SQLQuery & "smtDaysSubmitted38, "
    SQLQuery = SQLQuery & "smtDaysSubmitted39, "
    SQLQuery = SQLQuery & "smtDaysSubmitted40, "
    SQLQuery = SQLQuery & "smtDaysSubmitted41, "
    SQLQuery = SQLQuery & "smtDaysSubmitted42, "
    SQLQuery = SQLQuery & "smtDaysSubmitted43, "
    SQLQuery = SQLQuery & "smtDaysSubmitted44, "
    SQLQuery = SQLQuery & "smtDaysSubmitted45, "
    SQLQuery = SQLQuery & "smtDaysSubmitted46, "
    SQLQuery = SQLQuery & "smtDaysSubmitted47, "
    SQLQuery = SQLQuery & "smtDaysSubmitted48, "
    SQLQuery = SQLQuery & "smtDaysSubmitted49, "
    SQLQuery = SQLQuery & "smtDaysSubmitted50, "
    SQLQuery = SQLQuery & "smtDaysSubmitted51, "
    SQLQuery = SQLQuery & "smtDaysSubmitted52, "
    SQLQuery = SQLQuery & "smtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    'SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & Format$(smNewestWeekDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & smClientAbbr & "', "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & ilMktRepUstCode & ", "
    SQLQuery = SQLQuery & ilServRepUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(Now, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    'llSmtCode = gInsertAndReturnCode(SQLQuery, "smt", "smtCode", "Replace")
    'If llSmtCode <= 0 Then
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mAddSmt"
        mAddSmt = False
        Exit Function
    End If
    mAddSmt = True
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mAddSmt"
    Exit Function
'ErrHand1:
'    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mAddSmt"
'    Return
End Function
Private Function mBuildWk1Smt(ilVefCode As Integer, ilShttCode As Integer, ilWeeksAired As Integer, ilWeeksMissing As Integer, llSpotPosted As Long, llSpotPostedSNC As Long, llSpotPostedNNC As Long, llDaysSubmitted As Long) As Integer
    Dim ilMktRepUstCode As Integer
    Dim ilServRepUstCode As Integer
    
    On Error GoTo ErrHand
    
    mBuildWk1Smt = False
    SQLQuery = "Select smtCode, smtWk1StartDate from smt where smtVefCode = " & ilVefCode & " AND smtShttCode = " & ilShttCode
    Set rst = cnn.Execute(SQLQuery)
    If rst.EOF Then
        'lmSmtCode = mAddSmt(ilVefCode, ilShttCode)
        'If lmSmtCode <= 0 Then
        If Not mAddSmt(ilVefCode, ilShttCode) Then
            Exit Function
        End If
    'Else
    '    lmSmtCode = rst!smtCode
    End If
    SQLQuery = "Update smt Set "
    SQLQuery = SQLQuery & "smtWk1StartDate = '" & Format$(smNewestWeekDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtClientAbbr = '" & gFixQuote(smClientAbbr) & "', "
    If mGetReps(ilShttCode, ilMktRepUstCode, ilServRepUstCode) Then
        SQLQuery = SQLQuery & "smtMktRepUstCode = " & ilMktRepUstCode & ", "
        SQLQuery = SQLQuery & "smtServRepUstCode = " & ilServRepUstCode & ", "
    End If
    SQLQuery = SQLQuery & "smtGenDate = '" & Format$(Now, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtWeeksAired1 = smtWeeksAired1 + " & ilWeeksAired & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing1 = smtWeeksMissing1 + " & ilWeeksMissing & ", "
    SQLQuery = SQLQuery & "smtSpotPosted1 = smtSpotPosted1 + " & llSpotPosted & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC1 = smtSpotPostedSNC1 + " & llSpotPostedSNC & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC1 = smtSpotPostedNNC1 + " & llSpotPostedNNC & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted1 = smtDaysSubmitted1 + " & llDaysSubmitted
    SQLQuery = SQLQuery & " Where smtVefCode = " & ilVefCode & " AND smtShttCode = " & ilShttCode
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mBuildWk1Smt"
        mBuildWk1Smt = False
        Exit Function
    End If
    mBuildWk1Smt = True
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mBuildWk1Smt"
    Exit Function
ErrHand1:
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mBuildWk1Smt"
    Return
End Function
Private Function mClearWk1Smt() As Integer
    On Error GoTo ErrHand
    
    mClearWk1Smt = False
    SQLQuery = "Update smt Set "
    SQLQuery = SQLQuery & "smtWk1StartDate = '" & Format$(smNewestWeekDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtClientAbbr = '" & gFixQuote(smClientAbbr) & "', "
    SQLQuery = SQLQuery & "smtGenDate = '" & Format$(Now, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtWeeksAired1 = 0 " & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing1 = 0 " & ", "
    SQLQuery = SQLQuery & "smtSpotPosted1 = 0 " & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC1 = 0 " & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC1 = 0 " & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted1 = 0 "
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mClearWk1Smt"
        mClearWk1Smt = False
        Exit Function
    End If
    mClearWk1Smt = True
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mClearWk1Smt"
    Exit Function
ErrHand1:
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mClearWk1Smt"
    Return
End Function
Private Function mAgeSmt() As Integer
    
    On Error GoTo ErrHand
    mAgeSmt = False
    SQLQuery = "Update smt Set "
    SQLQuery = SQLQuery & "smtWk1StartDate = '" & Format$(smNewestWeekDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtClientAbbr = '" & gFixQuote(smClientAbbr) & "', "
    SQLQuery = SQLQuery & "smtGenDate = '" & Format$(Now, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "smtWeeksAired52 = smtWeeksAired51" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired51 = smtWeeksAired50" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired50 = smtWeeksAired49" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired49 = smtWeeksAired48" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired48 = smtWeeksAired47" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired47 = smtWeeksAired46" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired46 = smtWeeksAired45" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired45 = smtWeeksAired44" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired44 = smtWeeksAired43" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired43 = smtWeeksAired42" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired42 = smtWeeksAired41" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired41 = smtWeeksAired40" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired40 = smtWeeksAired39" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired39 = smtWeeksAired38" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired38 = smtWeeksAired37" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired37 = smtWeeksAired36" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired36 = smtWeeksAired35" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired35 = smtWeeksAired34" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired34 = smtWeeksAired33" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired33 = smtWeeksAired32" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired32 = smtWeeksAired31" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired31 = smtWeeksAired30" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired30 = smtWeeksAired29" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired29 = smtWeeksAired28" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired28 = smtWeeksAired27" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired27 = smtWeeksAired26" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired26 = smtWeeksAired25" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired25 = smtWeeksAired24" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired24 = smtWeeksAired23" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired23 = smtWeeksAired22" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired22 = smtWeeksAired21" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired21 = smtWeeksAired20" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired20 = smtWeeksAired19" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired19 = smtWeeksAired18" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired18 = smtWeeksAired17" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired17 = smtWeeksAired16" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired16 = smtWeeksAired15" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired15 = smtWeeksAired14" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired14 = smtWeeksAired13" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired13 = smtWeeksAired12" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired12 = smtWeeksAired11" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired11 = smtWeeksAired10" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired10 = smtWeeksAired9" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired9 = smtWeeksAired8" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired8 = smtWeeksAired7" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired7 = smtWeeksAired6" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired6 = smtWeeksAired5" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired5 = smtWeeksAired4" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired4 = smtWeeksAired3" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired3 = smtWeeksAired2" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired2 = smtWeeksAired1" & ", "
    SQLQuery = SQLQuery & "smtWeeksAired1 = 0" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing52 = smtWeeksMissing51" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing51 = smtWeeksMissing50" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing50 = smtWeeksMissing49" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing49 = smtWeeksMissing48" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing48 = smtWeeksMissing47" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing47 = smtWeeksMissing46" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing46 = smtWeeksMissing45" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing45 = smtWeeksMissing44" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing44 = smtWeeksMissing43" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing43 = smtWeeksMissing42" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing42 = smtWeeksMissing41" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing41 = smtWeeksMissing30" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing40 = smtWeeksMissing39" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing39 = smtWeeksMissing38" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing38 = smtWeeksMissing37" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing37 = smtWeeksMissing36" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing36 = smtWeeksMissing35" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing35 = smtWeeksMissing34" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing34 = smtWeeksMissing33" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing33 = smtWeeksMissing32" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing32 = smtWeeksMissing31" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing31 = smtWeeksMissing30" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing30 = smtWeeksMissing29" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing29 = smtWeeksMissing28" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing28 = smtWeeksMissing27" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing27 = smtWeeksMissing26" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing26 = smtWeeksMissing25" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing25 = smtWeeksMissing24" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing24 = smtWeeksMissing23" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing23 = smtWeeksMissing22" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing22 = smtWeeksMissing21" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing21 = smtWeeksMissing20" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing20 = smtWeeksMissing19" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing19 = smtWeeksMissing18" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing18 = smtWeeksMissing17" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing17 = smtWeeksMissing16" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing16 = smtWeeksMissing15" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing15 = smtWeeksMissing14" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing14 = smtWeeksMissing13" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing13 = smtWeeksMissing12" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing12 = smtWeeksMissing11" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing11 = smtWeeksMissing10" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing10 = smtWeeksMissing9" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing9 = smtWeeksMissing8" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing8 = smtWeeksMissing7" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing7 = smtWeeksMissing6" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing6 = smtWeeksMissing5" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing5 = smtWeeksMissing4" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing4 = smtWeeksMissing3" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing3 = smtWeeksMissing2" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing2 = smtWeeksMissing1" & ", "
    SQLQuery = SQLQuery & "smtWeeksMissing1 = 0" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted52 = smtSpotPosted51" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted51 = smtSpotPosted50" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted50 = smtSpotPosted49" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted49 = smtSpotPosted48" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted48 = smtSpotPosted47" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted47 = smtSpotPosted46" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted46 = smtSpotPosted45" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted45 = smtSpotPosted44" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted44 = smtSpotPosted43" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted43 = smtSpotPosted42" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted42 = smtSpotPosted41" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted41 = smtSpotPosted40" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted40 = smtSpotPosted39" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted39 = smtSpotPosted38" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted38 = smtSpotPosted37" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted37 = smtSpotPosted36" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted36 = smtSpotPosted35" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted35 = smtSpotPosted34" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted34 = smtSpotPosted33" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted33 = smtSpotPosted32" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted32 = smtSpotPosted31" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted31 = smtSpotPosted30" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted30 = smtSpotPosted29" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted29 = smtSpotPosted28" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted28 = smtSpotPosted27" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted27 = smtSpotPosted26" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted26 = smtSpotPosted25" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted25 = smtSpotPosted24" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted24 = smtSpotPosted23" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted23 = smtSpotPosted22" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted22 = smtSpotPosted21" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted21 = smtSpotPosted20" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted20 = smtSpotPosted19" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted19 = smtSpotPosted18" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted18 = smtSpotPosted17" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted17 = smtSpotPosted16" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted16 = smtSpotPosted15" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted15 = smtSpotPosted14" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted14 = smtSpotPosted13" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted13 = smtSpotPosted12" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted12 = smtSpotPosted11" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted11 = smtSpotPosted10" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted10 = smtSpotPosted9" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted9 = smtSpotPosted8" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted8 = smtSpotPosted7" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted7 = smtSpotPosted6" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted6 = smtSpotPosted5" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted5 = smtSpotPosted4" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted4 = smtSpotPosted3" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted3 = smtSpotPosted2" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted2 = smtSpotPosted1" & ", "
    SQLQuery = SQLQuery & "smtSpotPosted1 = 0" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC52 = smtSpotPostedSNC51" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC51 = smtSpotPostedSNC50" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC50 = smtSpotPostedSNC49" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC49 = smtSpotPostedSNC48" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC48 = smtSpotPostedSNC47" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC47 = smtSpotPostedSNC46" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC46 = smtSpotPostedSNC45" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC45 = smtSpotPostedSNC44" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC44 = smtSpotPostedSNC43" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC43 = smtSpotPostedSNC42" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC42 = smtSpotPostedSNC41" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC41 = smtSpotPostedSNC40" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC40 = smtSpotPostedSNC39" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC39 = smtSpotPostedSNC38" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC38 = smtSpotPostedSNC37" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC37 = smtSpotPostedSNC36" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC36 = smtSpotPostedSNC35" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC35 = smtSpotPostedSNC34" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC34 = smtSpotPostedSNC33" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC33 = smtSpotPostedSNC32" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC32 = smtSpotPostedSNC31" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC31 = smtSpotPostedSNC30" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC30 = smtSpotPostedSNC29" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC29 = smtSpotPostedSNC28" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC28 = smtSpotPostedSNC27" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC27 = smtSpotPostedSNC26" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC26 = smtSpotPostedSNC25" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC25 = smtSpotPostedSNC24" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC24 = smtSpotPostedSNC23" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC23 = smtSpotPostedSNC22" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC22 = smtSpotPostedSNC21" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC21 = smtSpotPostedSNC20" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC20 = smtSpotPostedSNC19" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC19 = smtSpotPostedSNC18" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC18 = smtSpotPostedSNC17" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC17 = smtSpotPostedSNC16" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC16 = smtSpotPostedSNC15" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC15 = smtSpotPostedSNC14" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC14 = smtSpotPostedSNC13" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC13 = smtSpotPostedSNC12" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC12 = smtSpotPostedSNC11" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC11 = smtSpotPostedSNC10" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC10 = smtSpotPostedSNC9" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC9 = smtSpotPostedSNC8" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC8 = smtSpotPostedSNC7" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC7 = smtSpotPostedSNC6" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC6 = smtSpotPostedSNC5" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC5 = smtSpotPostedSNC4" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC4 = smtSpotPostedSNC3" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC3 = smtSpotPostedSNC2" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC2 = smtSpotPostedSNC1" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedSNC1 = 0" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC52 = smtSpotPostedNNC51" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC51 = smtSpotPostedNNC50" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC50 = smtSpotPostedNNC49" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC49 = smtSpotPostedNNC48" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC48 = smtSpotPostedNNC47" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC47 = smtSpotPostedNNC46" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC46 = smtSpotPostedNNC45" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC45 = smtSpotPostedNNC44" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC44 = smtSpotPostedNNC43" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC43 = smtSpotPostedNNC42" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC42 = smtSpotPostedNNC41" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC41 = smtSpotPostedNNC40" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC40 = smtSpotPostedNNC39" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC39 = smtSpotPostedNNC38" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC38 = smtSpotPostedNNC37" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC37 = smtSpotPostedNNC36" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC36 = smtSpotPostedNNC35" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC35 = smtSpotPostedNNC34" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC34 = smtSpotPostedNNC33" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC33 = smtSpotPostedNNC32" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC32 = smtSpotPostedNNC31" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC31 = smtSpotPostedNNC30" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC30 = smtSpotPostedNNC29" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC29 = smtSpotPostedNNC28" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC28 = smtSpotPostedNNC27" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC27 = smtSpotPostedNNC26" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC26 = smtSpotPostedNNC25" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC25 = smtSpotPostedNNC24" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC24 = smtSpotPostedNNC23" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC23 = smtSpotPostedNNC22" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC22 = smtSpotPostedNNC21" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC21 = smtSpotPostedNNC20" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC20 = smtSpotPostedNNC19" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC19 = smtSpotPostedNNC18" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC18 = smtSpotPostedNNC17" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC17 = smtSpotPostedNNC16" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC16 = smtSpotPostedNNC15" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC15 = smtSpotPostedNNC14" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC14 = smtSpotPostedNNC13" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC13 = smtSpotPostedNNC12" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC12 = smtSpotPostedNNC11" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC11 = smtSpotPostedNNC10" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC10 = smtSpotPostedNNC9" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC9 = smtSpotPostedNNC8" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC8 = smtSpotPostedNNC7" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC7 = smtSpotPostedNNC6" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC6 = smtSpotPostedNNC5" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC5 = smtSpotPostedNNC4" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC4 = smtSpotPostedNNC3" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC3 = smtSpotPostedNNC2" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC2 = smtSpotPostedNNC1" & ", "
    SQLQuery = SQLQuery & "smtSpotPostedNNC1 = 0" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted52 = smtDaysSubmitted51" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted51 = smtDaysSubmitted50" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted50 = smtDaysSubmitted49" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted49 = smtDaysSubmitted48" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted48 = smtDaysSubmitted47" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted47 = smtDaysSubmitted46" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted46 = smtDaysSubmitted45" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted45 = smtDaysSubmitted44" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted44 = smtDaysSubmitted43" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted43 = smtDaysSubmitted42" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted42 = smtDaysSubmitted41" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted41 = smtDaysSubmitted40" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted40 = smtDaysSubmitted39" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted39 = smtDaysSubmitted38" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted38 = smtDaysSubmitted37" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted37 = smtDaysSubmitted36" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted36 = smtDaysSubmitted35" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted35 = smtDaysSubmitted34" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted34 = smtDaysSubmitted33" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted33 = smtDaysSubmitted32" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted32 = smtDaysSubmitted31" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted31 = smtDaysSubmitted30" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted30 = smtDaysSubmitted29" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted29 = smtDaysSubmitted28" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted28 = smtDaysSubmitted27" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted27 = smtDaysSubmitted26" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted26 = smtDaysSubmitted25" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted25 = smtDaysSubmitted24" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted24 = smtDaysSubmitted23" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted23 = smtDaysSubmitted22" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted22 = smtDaysSubmitted21" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted21 = smtDaysSubmitted20" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted20 = smtDaysSubmitted19" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted19 = smtDaysSubmitted18" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted18 = smtDaysSubmitted17" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted17 = smtDaysSubmitted16" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted16 = smtDaysSubmitted15" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted15 = smtDaysSubmitted14" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted14 = smtDaysSubmitted13" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted13 = smtDaysSubmitted12" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted12 = smtDaysSubmitted11" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted11 = smtDaysSubmitted10" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted10 = smtDaysSubmitted9" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted9 = smtDaysSubmitted8" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted8 = smtDaysSubmitted7" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted7 = smtDaysSubmitted6" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted6 = smtDaysSubmitted5" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted5 = smtDaysSubmitted4" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted4 = smtDaysSubmitted3" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted3 = smtDaysSubmitted2" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted2 = smtDaysSubmitted1" & ", "
    SQLQuery = SQLQuery & "smtDaysSubmitted1 = 0" & ", "
    SQLQuery = SQLQuery & "smtUnused = '" & "" & "' "
    SQLQuery = SQLQuery & " Where smtWk1StartDate <> '" & Format$(smNewestWeekDate, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mAgeSmt"
        mAgeSmt = False
        Exit Function
    End If
    mAgeSmt = True
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mAgeSmt"
    Exit Function
ErrHand1:
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mAgeSmt"
    Return
End Function

Private Function mGenerate() As Boolean
    Dim ilRet As Integer
    Dim ilWeeksAired As Integer
    Dim ilWeeksMissing As Integer
    Dim llSpotPosted As Long
    Dim llSpotPostedSNC As Long
    Dim llSpotPostedNNC As Long
    Dim llDaysSubmitted As Long
    Dim llAttCount As Long
    Dim llCount As Long
    Dim slStart As String
    Dim slEnd As String
    
    On Error GoTo ErrHand
    
    gLogMsg "Affiliate Measurement Started: " & Trim$(lacDates.Caption), "AffiliateMeasurement.Txt", False

    DoEvents
    lacAging(0).Visible = True
    lacAging(1).Visible = True
    lacAging(1).Caption = ""
    lacClearing(0).Visible = True
    lacClearing(1).Visible = True
    lacClearing(1).Caption = ""
    lacBuilding(0).Visible = True
    lacBuilding(1).Visible = True
    lacBuilding(1).Caption = ""
    DoEvents
    'Age Weeks
    lacAging(1).Caption = "Started"
    DoEvents
    ilRet = mAgeSmt()
    lacAging(1).Caption = "Completed"
    gLogMsg "Aging Affiliate Measurements Completed", "AffiliateMeasurement.Txt", False
    DoEvents
    
    'Clear First Week
    lacClearing(1).Caption = "Started"
    DoEvents
    ilRet = mClearWk1Smt()
    lacClearing(1).Caption = "Completed"
    gLogMsg "Clearing First Week Affiliate Measurements Completed", "AffiliateMeasurement.Txt", False
    DoEvents
    
    'Build First Week
    lacBuilding(1).Caption = "Started"
    DoEvents
    SQLQuery = "SELECT Count(*)"
    SQLQuery = SQLQuery & " FROM att"
    SQLQuery = SQLQuery & " WHERE (attOnAir <= '" & Format$(smNewestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attOffAir >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attDropDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attOnAir <= attOffAir"
    SQLQuery = SQLQuery & " AND attOnAir <= attDropDate"
    SQLQuery = SQLQuery & " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery & " AND attExportToWeb = 'Y'"
    SQLQuery = SQLQuery & ")"
    Set att_rst = cnn.Execute(SQLQuery)
    If Not att_rst.EOF Then
        llAttCount = att_rst(0).Value
    Else
        llAttCount = 0
    End If
    llCount = 0
    SQLQuery = "SELECT attCode, attVefCode, attShfCode, attExportToWeb, attWebInterface"
    SQLQuery = SQLQuery & " FROM att"
    SQLQuery = SQLQuery & " WHERE (attOnAir <= '" & Format$(smNewestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attOffAir >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attDropDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND attOnAir <= attOffAir"
    SQLQuery = SQLQuery & " AND attOnAir <= attDropDate"
    SQLQuery = SQLQuery & " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery & " AND attExportToWeb = 'Y'"
    SQLQuery = SQLQuery & ")"
    Set att_rst = cnn.Execute(SQLQuery)
    Do While Not att_rst.EOF
        DoEvents
        SQLQuery = "SELECT Count(*) FROM cptt WHERE (cpttAtfCode = " & att_rst!attCode
        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpttStartDate <= '" & Format$(smNewestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & ")"
        Set cptt_rst = cnn.Execute(SQLQuery)
        If Not cptt_rst.EOF Then
            ilWeeksAired = cptt_rst(0).Value
        Else
            ilWeeksAired = 0
        End If
        DoEvents
        SQLQuery = "SELECT Count(*) FROM cptt WHERE (cpttAtfCode = " & att_rst!attCode
        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpttStartDate <= '" & Format$(smNewestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpttPostingStatus <> 2"
        SQLQuery = SQLQuery & ")"
        Set cptt_rst = cnn.Execute(SQLQuery)
        If Not cptt_rst.EOF Then
            ilWeeksMissing = cptt_rst(0).Value
        Else
            ilWeeksMissing = 0
        End If
        DoEvents
    
        SQLQuery = "Select Count(*) FROM ast WHERE"
        SQLQuery = SQLQuery + " astAtfCode = " & att_rst!attCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astStatus <> 8"
        SQLQuery = SQLQuery & " AND astCPStatus <> 0"
        SQLQuery = SQLQuery & ")"
        Set ast_rst = cnn.Execute(SQLQuery)
        If Not ast_rst.EOF Then
            llSpotPosted = ast_rst(0).Value
        Else
            llSpotPosted = 0
        End If
        DoEvents
        
        SQLQuery = "Select Count(*) FROM ast WHERE"
        SQLQuery = SQLQuery + " astAtfCode = " & att_rst!attCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astStatus <> 8"
        SQLQuery = SQLQuery & " AND astCPStatus <> 0"
        SQLQuery = SQLQuery & " AND astStationCompliant <> 'Y'"
        SQLQuery = SQLQuery & ")"
        Set ast_rst = cnn.Execute(SQLQuery)
        If Not ast_rst.EOF Then
            llSpotPostedSNC = ast_rst(0).Value
        Else
            llSpotPostedSNC = 0
        End If
        DoEvents
        
        SQLQuery = "Select Count(*) FROM ast WHERE"
        SQLQuery = SQLQuery + " astAtfCode = " & att_rst!attCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astFeedDate <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND astStatus <> 8"
        SQLQuery = SQLQuery & " AND astCPStatus <> 0"
        SQLQuery = SQLQuery & " AND astAgencyCompliant <> 'Y'"
        SQLQuery = SQLQuery & ")"
        Set ast_rst = cnn.Execute(SQLQuery)
        If Not ast_rst.EOF Then
            llSpotPostedNNC = ast_rst(0).Value
        Else
            llSpotPostedNNC = 0
        End If
        DoEvents
        
        llDaysSubmitted = 0
        'SQLQuery = "SELECT Count(*) FROM Webl WHERE "
        'SQLQuery = SQLQuery & " weblType = 1 And weblAttCode = " & att_rst!attCode
        'SQLQuery = SQLQuery & " AND weblPostDay >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & " AND weblPostDay <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & " AND WeblDate In ("
        'SQLQuery = SQLQuery & "SELECT Distinct weblDate FROM webl WHERE"
        'SQLQuery = SQLQuery & " weblType = 1 And weblAttCode = " & att_rst!attCode
        'SQLQuery = SQLQuery & " AND weblPostDay >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & " AND weblPostDay <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & ")"
        SQLQuery = "SELECT Count(Distinct weblDate) FROM Webl WHERE "
        SQLQuery = SQLQuery & " weblType = 1 And weblAttCode = " & att_rst!attCode
        SQLQuery = SQLQuery & " AND weblPostDay >= '" & Format$(smOldestWeekDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND weblPostDay <= '" & Format$(smSuNewestWeekDate, sgSQLDateForm) & "'"
        Set webl_rst = cnn.Execute(SQLQuery)
        If Not webl_rst.EOF Then
            llDaysSubmitted = webl_rst(0).Value
        End If
        ilRet = mBuildWk1Smt(att_rst!attvefCode, att_rst!attshfCode, ilWeeksAired, ilWeeksMissing, llSpotPosted, llSpotPostedSNC, llSpotPostedNNC, llDaysSubmitted)
        llCount = llCount + 1
        lacBuilding(1) = llCount & " of " & llAttCount
        DoEvents
        att_rst.MoveNext
    Loop
    lacBuilding(1).Caption = "Completed"
    gLogMsg "Building First Week Affiliate Measurements Completed", "AffiliateMeasurement.Txt", False
        DoEvents
    mGenerate = True
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mGenerate"
    Resume Next
ErrHand1:
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mGenerate"
    Return
End Function



Private Function mGetReps(ilShttCode As Integer, ilMktRepUstCode As Integer, ilServRepUstCode As Integer) As Boolean
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    slSQLQuery = "Select shttMktRepUstCode, shttServRepUstCode from shtt where shttcode = " & ilShttCode
    Set rst = cnn.Execute(slSQLQuery)
    If Not rst.EOF Then
        ilMktRepUstCode = rst!shttMktRepUstCode
        ilServRepUstCode = rst!shttServRepUstCode
        mGetReps = True
    Else
        mGetReps = False
    End If
    Exit Function
ErrHand:
    gHandleError "AffiliateMeasurement.Txt", "AffiliateMeasurement-mGetReps"
    Resume Next
ErrHand1:
    gHandleError "AffiliateMeasurement.txt", "AffiliateMeasurement-mGetReps"
    Return
End Function

Private Sub mInitAffiliateMeasurement()
    Dim slDate As String

    SQLQuery = "SELECT mnfName"
    SQLQuery = SQLQuery & " FROM SPF_Site_Options, MNF_Multi_Names"
    SQLQuery = SQLQuery & " WHERE spfCode = 1"
    SQLQuery = SQLQuery & " AND spfMnfClientAbbr = mnfCode"
    Set rst = cnn.Execute(SQLQuery)
    If Not rst.EOF Then
        smClientAbbr = Trim$(rst!mnfName)
    Else
        smClientAbbr = sgClientName
    End If
    slDate = DateAdd("d", -1, Format(Now, "m/d/yy"))
    smSuNewestWeekDate = gObtainPrevSunday(slDate)
    smNewestWeekDate = gObtainPrevMonday(smSuNewestWeekDate)
    smOldestWeekDate = gObtainPrevMonday(DateAdd("ww", -51, smNewestWeekDate))
    lacDates.Caption = "From " & smOldestWeekDate & " Thru " & smNewestWeekDate
End Sub
