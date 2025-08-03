VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Station Spot Builder"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "AffStationSpotBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.PictureBox plcSignon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      Picture         =   "AffStationSpotBuilder.frx":08CA
      ScaleHeight     =   6435
      ScaleWidth      =   10125
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10185
      Begin VB.Timer tmcSetTime 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   1305
         Top             =   5730
      End
      Begin VB.Timer tmcRestartTask 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1695
         Top             =   5175
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   690
         Top             =   5220
      End
      Begin VB.CommandButton cmcMin 
         Caption         =   "Minimize"
         Height          =   330
         Left            =   4380
         TabIndex        =   2
         Top             =   5415
         Width           =   1380
      End
      Begin VB.CommandButton cmcStop 
         Caption         =   "Stop"
         Height          =   330
         Left            =   4380
         TabIndex        =   3
         Top             =   5940
         Width           =   1380
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   -45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   615
         Width           =   105
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
         Height          =   4155
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   825
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7329
         _Version        =   393216
         Rows            =   4
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorUnpopulated=   -2147483634
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************************************************
'*  EngrDayName - enters affiliate representative information
'*
'*
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imCancelled As Integer
Private imClosed As Integer
Private lmSleepTime As Long
Private imFirstTime As Integer

Private smDate As String
Private imVefCode As Integer
Private imShttCode As Integer
Private imAdfCode As Integer
Private smSpotCopyDays As String
Private smRegionDays As String

Private lmProcessRow As Long

Private hmAst As Integer

Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO

Private rst_abf As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset
Private rst_Ent As ADODB.Recordset
Private rst_Site As ADODB.Recordset

Private Const VEHICLEINDEX = 0
Private Const STATUSINDEX = 1
Private Const SOURCEINDEX = 2
Private Const GENDATEINDEX = 3
'Private Const SPOTCOPYCHGINDEX = 3
'Private Const REGIONCHGINDEX = 4
Private Const ENTEREDINDEX = 4  '5
'Private Const ENTEREDTIMEINDEX = 5  '6
Private Const ABFCODEINDEX = 6  '7





Private Sub cmcMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmcStop_Click()
    imCancelled = True
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    
    If imFirstTime Then
        igAlertInterval = 0
        mSetGridColumns
        mSetGridTitles
        'gGrid_IntegralHeight grdStatus
        mClearGrid
        gGrid_IntegralHeight grdStatus
        gGrid_FillWithRows grdStatus
        mSetStatusGridColor
        
        imFirstTime = False
    End If

End Sub

Private Sub Form_Load()
    Dim ilPos As Integer
    
    sgCommand = Command$
    If App.PrevInstance Then
        MsgBox "Only one copy of Station Spot Builder can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of Station Spot Builder path: " & App.Path & " from " & Trim$(gGetComputerName()), "StationSpotBuilder.txt", False
        End
    End If
    imFirstTime = True
    igExportSource = 2
    sgUserName = "StationSpotBuilder"
    gCenterStdAlone Me
    tmcStart.Enabled = True
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRet As Integer
    
    If imClosed = True Then
        Exit Sub
    End If
    tmcRestartTask.Enabled = False
    tmcStart.Enabled = False
    If imCancelled Then
        ilRet = MsgBox("Stop the Station Spot Builder", vbQuestion + vbYesNo, "Stop Service")
        If ilRet = vbNo Then
            Cancel = 1
            imCancelled = False
            tmcRestartTask.Enabled = True
            Exit Sub
        End If
    End If
    imClosed = True
End Sub

Private Sub Form_Resize()
    'If Me.WindowState = vbNormal Then
    '    Me.Left = Screen.Width / 2 - Me.Width / 2
    '    Me.Top = Screen.Height / 2 - Me.Height / 2
    'End If
End Sub

Private Sub mStartUp()
    Dim ilRet As Integer
    Dim slTime As String
    Dim slBuffer As String
    Dim slTimeOut As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    '6548
    Dim slXMLINIInputFile As String
    tmcStart.Enabled = False
    
    
    'igDemoMode = False
    'If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    'End If
    
    sgCurDir = CurDir$

    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF

    lmSleepTime = 1000 ' 5 seconds '300000    '5 Minutes
    imCancelled = False
    imClosed = False
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
    If InStr(1, sgStartupDirectory, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
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
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        gMsgBox "Affiliat.Ini [Database] 'Name' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Unload frmMain
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
    If gLoadOption("SQLSpec", "System", slBuffer) Then
        If slBuffer = "P7" Then
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
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Unload frmMain
        Exit Sub
    Else
        sgMsgDirectory = gSetPathEndSlash(sgMsgDirectory, True) & "Messages\"
    End If

    sgDSN = sgDatabaseName
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(sgDSN, Len(sgDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        sgDSN = Left(sgDSN, Len(sgDSN) - 1)
    End If
    
    Set cnn = New ADODB.Connection
    On Error GoTo ERRNOPERVASIVE
    ilRet = 0
    cnn.Open "DSN=" & sgDSN
    
    On Error GoTo ErrHand
    If ilRet = 1 Then
        Sleep 2000
        cnn.Open "DSN=" & sgDSN
    End If
    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
    ''6548 Dan M
    ReDim sgXDSSection(0 To 0) As String
    'slXMLINIInputFile = gXmlIniPath(True)
    'If LenB(slXMLINIInputFile) <> 0 Then
    '    ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    'End If
    
    ilRet = mOpenPervasiveAPI()
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    igShowMsgBox = False
    
    mGetGuideCode
    ilRet = gInitGlobals()
    'mClearAnyProcessing
    mPopArrays
    mPopulate
    Exit Sub
    
mReadFileErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffStationSpotBuilderLog", "Form-mStartUp"
    Unload frmMain
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    tmcSetTime.Enabled = False
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    Erase tmCPDat
    Erase tmAstInfo
    rst_abf.Close
    rst_att.Close
    rst_Cptt.Close
    rst_Ent.Close
    rst_Site.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    btrStopAppl
    Set frmMain = Nothing   'Remove data segment
    End
End Sub


Private Sub tmcRestartTask_Timer()
    tmcRestartTask.Enabled = False
    mTaskLoop
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "SSB"
    mUpdateTimeRecord
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mStartUp
    gLogMsg "Station Spot Builder path: " & App.Path & " from " & Trim$(gGetComputerName()), "StationSpotBuilder.txt", False
'    tmcTask.Interval = CInt(lmSleepTime)
'    tmcTask.Enabled = True
    sgTimeZone = Left$(gGetLocalTZName(), 1)
    tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
    tmcSetTime.Enabled = True
    mTaskLoop
End Sub



Private Sub mPopulate()
    Dim llVef As Long
    Dim llRow As Long
    Dim slDays As String * 7
    
    On Error GoTo ErrHand
    lmProcessRow = -1
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
'    SQLQuery = "SELECT * FROM AUF_Alert_User WHERE aufType = 'B' AND aufStatus = 'R' ORDER BY aufEnteredDate, aufEnteredTime"
'    Set rst_Auf = cnn.Execute(SQLQuery)
'    Do While Not rst_Auf.EOF
'        If llRow >= grdStatus.Rows Then
'            grdStatus.AddItem ""
'        End If
'        grdStatus.Row = llRow
'        llVef = gBinarySearchVef(CLng(rst_Auf!aufVefCode))
'        If llVef <> -1 Then
'            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
'        Else
'            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = "Vehicle Code = " & rst_Auf!aufVefCode
'        End If
'        If rst_Auf!aufSubType = "P" Then
'            lmProcessRow = llRow
'            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Processing"
'        Else
'            If lmProcessRow = -1 Then
'                lmProcessRow = llRow
'            End If
'            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Ready"
'        End If
'
'        grdStatus.TextMatrix(llRow, GENDATEINDEX) = Format(rst_Auf!aufMoWeekDate, sgShowDateForm)
'        slDays = "-------"
'        If rst_Auf!aufSpotCopyChg1 = "Y" Then
'            Mid(slDays, 1, 1) = "M"
'        End If
'        If rst_Auf!aufSpotCopyChg2 = "Y" Then
'            Mid(slDays, 2, 1) = "T"
'        End If
'        If rst_Auf!aufSpotCopyChg3 = "Y" Then
'            Mid(slDays, 3, 1) = "W"
'        End If
'        If rst_Auf!aufSpotCopyChg4 = "Y" Then
'            Mid(slDays, 4, 1) = "T"
'        End If
'        If rst_Auf!aufSpotCopyChg5 = "Y" Then
'            Mid(slDays, 5, 1) = "F"
'        End If
'        If rst_Auf!aufSpotCopyChg6 = "Y" Then
'            Mid(slDays, 6, 1) = "S"
'        End If
'        If rst_Auf!aufSpotCopyChg7 = "Y" Then
'            Mid(slDays, 7, 1) = "S"
'        End If
'        grdStatus.TextMatrix(llRow, SPOTCOPYCHGINDEX) = slDays
'
'        slDays = "-------"
'        If rst_Auf!aufRegionCopyChg1 = "Y" Then
'            Mid(slDays, 1, 1) = "M"
'        End If
'        If rst_Auf!aufRegionCopyChg2 = "Y" Then
'            Mid(slDays, 2, 1) = "T"
'        End If
'        If rst_Auf!aufRegionCopyChg3 = "Y" Then
'            Mid(slDays, 3, 1) = "W"
'        End If
'        If rst_Auf!aufRegionCopyChg4 = "Y" Then
'            Mid(slDays, 4, 1) = "T"
'        End If
'        If rst_Auf!aufRegionCopyChg5 = "Y" Then
'            Mid(slDays, 5, 1) = "F"
'        End If
'        If rst_Auf!aufRegionCopyChg6 = "Y" Then
'            Mid(slDays, 6, 1) = "S"
'        End If
'        If rst_Auf!aufRegionCopyChg7 = "Y" Then
'            Mid(slDays, 7, 1) = "S"
'        End If
'        grdStatus.TextMatrix(llRow, REGIONCHGINDEX) = slDays
'        grdStatus.TextMatrix(llRow, ENTEREDDATEINDEX) = Format(rst_Auf!aufEnteredDate, sgShowDateForm)
'        grdStatus.TextMatrix(llRow, ENTEREDTIMEINDEX) = Format(rst_Auf!aufEnteredTime, sgShowTimeWSecForm)
'        grdStatus.TextMatrix(llRow, AUFCODEINDEX) = rst_Auf!aufCode
'        llRow = llRow + 1
'        rst_Auf.MoveNext
'    Loop
    SQLQuery = "SELECT * FROM abf_AST_Build_Queue WHERE abfStatus = 'G' Or abfStatus = 'P'  Or abfStatus = 'H' ORDER BY abfStatus, abfGenStartDate"
    Set rst_abf = gSQLSelectCall(SQLQuery)
    Do While Not rst_abf.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
        llVef = gBinarySearchVef(CLng(rst_abf!abfVefCode))
        If llVef <> -1 Then
            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = Trim$(tgVehicleInfo(llVef).sVehicle)
        Else
            grdStatus.TextMatrix(llRow, VEHICLEINDEX) = "Vehicle Code = " & rst_abf!abfVefCode
        End If
        If rst_abf!abfSource = "A" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Agreement"
        ElseIf rst_abf!abfSource = "P" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Post Log"
        ElseIf rst_abf!abfSource = "F" Then
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Fast Add"
        Else
            grdStatus.TextMatrix(llRow, SOURCEINDEX) = "Log"
        End If
'        If rst_Auf!aufSubType = "P" Then
'            lmProcessRow = llRow
'            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Processing"
'        Else
'            If lmProcessRow = -1 Then
'                lmProcessRow = llRow
'            End If
'            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Ready"
'        End If
        
        If rst_abf!abfStatus = "P" Then
            lmProcessRow = llRow
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Processing"
        ElseIf rst_abf!abfStatus = "H" Then
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Not Ready"
        Else
            If lmProcessRow = -1 Then
                lmProcessRow = llRow
            End If
            grdStatus.TextMatrix(llRow, STATUSINDEX) = "Ready"
        End If
        grdStatus.TextMatrix(llRow, GENDATEINDEX) = Format(rst_abf!abfGenStartDate, sgShowDateForm) & "-" & Format(rst_abf!abfGenEndDate, sgShowDateForm)
        grdStatus.TextMatrix(llRow, ENTEREDINDEX) = Format(rst_abf!abfEnteredDate, sgShowDateForm) & " " & Format(rst_abf!abfEnteredTime, sgShowTimeWSecForm)
        grdStatus.TextMatrix(llRow, ABFCODEINDEX) = rst_abf!abfCode
        llRow = llRow + 1
        rst_abf.MoveNext
    Loop
    mSetStatusGridColor
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "StationSpotBuilder.txt", "Station Spot Builder-mPopulate"
    grdStatus.Redraw = True
End Sub

Private Sub mEraseArrays()
End Sub

Private Sub mTaskLoop()
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilTaskCount As Integer
    Dim llRow As Long
    Dim llPriorityRow As Long
    Dim blAnyProcessing As Boolean
    Dim ilPriority As Integer
    Dim llEqtCode As Long
    Dim blMarketron As Boolean
    Dim blUnivision As Boolean
    Dim blCumulus As Boolean
    Dim blCSIWeb As Boolean
    Dim ilEqtStatus As Integer '0=Update, 1=Set as in Error and don't update LDE it was defind to mean Delete
    Dim llAbfCode As Long
    
    On Error GoTo ErrHand
    ilTaskCount = -1
    Do
        Sleep lmSleepTime
        If imCancelled Then
            Unload frmMain
            Exit Sub
        End If
        For ilLoop = 0 To 100 Step 1
            DoEvents
        Next ilLoop
        If (ilTaskCount = -1) Or (ilTaskCount = 60) Then
            'Determine if any task needs to be performed
            'mUpdateTimeRecord
            SQLQuery = "UPDATE Site SET "
            SQLQuery = SQLQuery & "siteSSBDate = '" & Format(Now, sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "siteSSBTime = '" & Format(Now, sgSQLTimeForm) & "'"
            'SQLQuery = SQLQuery & "WHERE siteCode = " & "1"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                gSetMousePointer grdStatus, grdStatus, vbDefault
                gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mTaskLoop"
                grdStatus.Redraw = True
                Exit Sub
            End If
            On Error GoTo ErrHand
            mPopulate
            DoEvents
            slDateTime = gNow()
            slNowDate = Format$(slDateTime, "m/d/yy")
            slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
            blAnyProcessing = False
            If lmProcessRow >= grdStatus.FixedRows Then
                mPopArrays
                smDate = grdStatus.TextMatrix(lmProcessRow, GENDATEINDEX)
                llAbfCode = grdStatus.TextMatrix(lmProcessRow, ABFCODEINDEX)
                'SQLQuery = "SELECT * FROM AUF_Alert_User WHERE aufCode = " & grdStatus.TextMatrix(lmProcessRow, AUFCODEINDEX)
                SQLQuery = "SELECT * FROM ABF_AST_Build_Queue WHERE abfCode = " & llAbfCode
                Set rst_abf = gSQLSelectCall(SQLQuery)
                If Not rst_abf.EOF Then
                    'SQLQuery = "UPDATE AUF_Alert_User SET "
                    'SQLQuery = SQLQuery & "aufSubType = '" & "P" & "' "
                    'SQLQuery = SQLQuery & "WHERE aufCode = " & grdStatus.TextMatrix(lmProcessRow, AUFCODEINDEX)
                    SQLQuery = "UPDATE ABF_Ast_Build_Queue SET "
                    SQLQuery = SQLQuery & "abfStatus = '" & "P" & "' "
                    SQLQuery = SQLQuery & "WHERE abfCode = " & llAbfCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand1:
                        gSetMousePointer grdStatus, grdStatus, vbDefault
                        gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mTaskLoop"
                        grdStatus.Redraw = True
                        Exit Sub
                    End If
                    On Error GoTo ErrHand
                    imVefCode = rst_abf!abfVefCode
                    imShttCode = rst_abf!abfShttCode
                    'smSpotCopyDays = grdStatus.TextMatrix(lmProcessRow, SPOTCOPYCHGINDEX)
                    'smRegionDays = grdStatus.TextMatrix(lmProcessRow, REGIONCHGINDEX)
                    gSetMousePointer grdStatus, grdStatus, vbHourglass
                    gUpdateTaskMonitor 1, "SSB"
                    ilRet = mBuildStationSpots()
                    gUpdateTaskMonitor 2, "SSB"
                    ''Update AUF record
                    'SQLQuery = "UPDATE AUF_Alert_User SET "
                    'If ilRet Then
                    '    SQLQuery = SQLQuery & "aufStatus = 'C'" & ", "  'Completed
                    'Else
                    '    SQLQuery = SQLQuery & "aufStatus = 'E'" & ", "  'Error
                    'End If
                    'SQLQuery = SQLQuery & "aufSubType = 'D'" & ", " 'Done
                    'SQLQuery = SQLQuery & "aufClearDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    'SQLQuery = SQLQuery & "aufClearTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                    'SQLQuery = SQLQuery & "aufClearUstCode = " & igUstCode & " "
                    'SQLQuery = SQLQuery & "WHERE aufCode = " & grdStatus.TextMatrix(lmProcessRow, AUFCODEINDEX)
                    'Update AUF record
                    SQLQuery = "UPDATE ABF_AST_Build_Queue SET "
                    If ilRet Then
                        SQLQuery = SQLQuery & "abfStatus = 'C'" & ", "  'Completed
                    Else
                        SQLQuery = SQLQuery & "abfStatus = 'E'" & ", "  'Error
                    End If
                    SQLQuery = SQLQuery & "abfCompletedDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "abfCompletedTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                    SQLQuery = SQLQuery & "abfUstCode = " & igUstCode & " "
                    SQLQuery = SQLQuery & "WHERE abfCode = " & llAbfCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand1:
                        gSetMousePointer grdStatus, grdStatus, vbDefault
                        gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mTaskLoop"
                        grdStatus.Redraw = True
                        Exit Sub
                    End If
                    On Error GoTo ErrHand
                    
                    'Create Affiliate Alert
                    
                    
                    gSetMousePointer grdStatus, grdStatus, vbDefault
                End If
            End If
            'ilTaskCount = -1    'Force to check for another without waiting
            If lmProcessRow = -1 Then
                ilTaskCount = 0
            Else
                ilTaskCount = -1
            End If
        Else
            ilTaskCount = ilTaskCount + 1
        End If
   Loop
   Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mTaskLoop"
    grdStatus.Redraw = True
    Exit Sub
'ErrHand1:
'    gSetMousePointer grdStatus, grdStatus, vbDefault
'    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mTaskLoop"
'    grdStatus.Redraw = True
'    Return
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(ABFCODEINDEX) = 0
    'grdStatus.ColWidth(VEHICLEINDEX) = grdStatus.Width * 0.2
    grdStatus.ColWidth(SOURCEINDEX) = grdStatus.Width * 0.11
    grdStatus.ColWidth(STATUSINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(GENDATEINDEX) = grdStatus.Width * 0.18
    'grdStatus.ColWidth(SPOTCOPYCHGINDEX) = grdStatus.Width * 0.13
    'grdStatus.ColWidth(REGIONCHGINDEX) = grdStatus.Width * 0.13
    grdStatus.ColWidth(ENTEREDINDEX) = grdStatus.Width * 0.18
    'grdStatus.ColWidth(ENTEREDTIMEINDEX) = grdStatus.Width * 0.11

    grdStatus.ColWidth(VEHICLEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = VEHICLEINDEX To ENTEREDINDEX Step 1
        If ilCol <> VEHICLEINDEX Then
            grdStatus.ColWidth(VEHICLEINDEX) = grdStatus.ColWidth(VEHICLEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus

End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, VEHICLEINDEX) = "Vehicle Name"
    grdStatus.TextMatrix(0, SOURCEINDEX) = "Source"
    grdStatus.TextMatrix(0, STATUSINDEX) = "Status"
    grdStatus.TextMatrix(0, GENDATEINDEX) = "Spot Dates"
    'grdStatus.TextMatrix(0, SPOTCOPYCHGINDEX) = "Spot/Copy"
    'grdStatus.TextMatrix(1, SPOTCOPYCHGINDEX) = "Day Changes"
    'grdStatus.TextMatrix(0, REGIONCHGINDEX) = "Region"
    'grdStatus.TextMatrix(1, REGIONCHGINDEX) = "Day Changes"
    grdStatus.TextMatrix(0, ENTEREDINDEX) = "Entered"
    'grdStatus.TextMatrix(0, ENTEREDTIMEINDEX) = "Entered"
    'grdStatus.TextMatrix(1, ENTEREDTIMEINDEX) = "Time"

End Sub

Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdStatus, True
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = VEHICLEINDEX To ENTEREDINDEX Step 1
            grdStatus.Row = llRow
            grdStatus.Col = llCol
            grdStatus.CellBackColor = LIGHTYELLOW
        Next llCol
    Next llRow
End Sub

Private Sub mClearGrid()
    
    gGrid_Clear grdStatus, True
End Sub


Private Sub mPopArrays()
    Dim ilRet As Integer
    
    ilRet = gPopCpf()
    'Note:  Market and Territory populates must be prior to Station Populate
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
    ilRet = gPopVff()
    
         
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
        If IsNull(rst!spfUsingFeatures5) Or (Len(rst!spfUsingFeatures5) = 0) Then
            sgSpfUsingFeatures5 = Chr$(0)
        Else
            sgSpfUsingFeatures5 = rst!spfUsingFeatures5
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
    
    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        ilRet = gObtainReplacments()
    Else
        ReDim tgRBofRec(0 To 0) As BOFREC
        ReDim tgSplitNetLastFill(0 To 0) As SPLITNETLASTFILL
    End If

    
    mCreateStatustype

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



Private Sub mGetGuideCode()
    On Error GoTo ErrHand
    SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        igUstCode = 1
    Else
        igUstCode = rst!ustCode
    End If
    Exit Sub
ErrHand:
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mGetGuideCode"
End Sub

Private Sub mAdjustPriority()
    Dim ilMin As Integer
    On Error GoTo ErrHand
'    SQLQuery = "SELECT Min(eqtPriority) FROM eqt_Export_Queue WHERE (eqtPriority > 0) AND (eqtStatus = 'R' or eqtStatus = 'P')"
'    Set rst_Eqt = cnn.Execute(SQLQuery)
'    If Not rst_Eqt.EOF Then
'        If rst_Eqt(0).Value <> vbNull Then
'            ilMin = rst_Eqt(0).Value
'            If ilMin > 1 Then
'                ilMin = ilMin - 1
'                SQLQuery = "UPDATE eqt_Export_Queue SET "
'                SQLQuery = SQLQuery & "eqtPriority = eqtPriority - " & ilMin
'                SQLQuery = SQLQuery & "WHERE eqtPriority > 1 AND eqtStatus = 'R'"
'                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                    GoSub ErrHand1:
'                End If
'            End If
'        End If
'    End If
    Exit Sub
ErrHand:
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mAdjustPriority"
    Exit Sub
ErrHand1:
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mAdjustPriority"
    
End Sub

Private Function mBuildStationSpots() As Integer
    Dim slMoDate As String
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim ilFirstDay As Integer
    Dim ilLastDay As Integer
    Dim ilPos As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    On Error GoTo ErrHand
    ilPos = InStr(1, smDate, "-", vbTextCompare)
    If ilPos > 0 Then
        slStartDate = Trim$(Left(smDate, ilPos - 1))
        slEndDate = Trim$(Mid(smDate, ilPos + 1))
    Else
        slStartDate = smDate
        slEndDate = smDate
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    'slMoDate = gObtainPrevMonday(smDate)
    slMoDate = gObtainPrevMonday(slStartDate)
    
    Do
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate"
        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
        If imShttCode > 0 Then
            SQLQuery = SQLQuery & " AND cpttShfCode = " & imShttCode
        End If
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        While Not rst_Cptt.EOF
            'Force spots to be created
            If rst_Cptt!cpttAstStatus = "C" Then
                SQLQuery = "UPDATE cptt SET "
                SQLQuery = SQLQuery + "cpttAstStatus = " & "'R'"
                SQLQuery = SQLQuery + " WHERE (cpttCode = " & rst_Cptt!cpttCode & ")"
                cnn.BeginTrans
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mBuildStationSpots"
                    cnn.RollbackTrans
                    mBuildStationSpots = False
                    Exit Function
                End If
                cnn.CommitTrans
            End If
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = rst_Cptt!cpttCode
            tgCPPosting(0).iStatus = rst_Cptt!cpttStatus
            tgCPPosting(0).iPostingStatus = rst_Cptt!cpttPostingStatus
            tgCPPosting(0).lAttCode = rst_Cptt!cpttatfCode
            tgCPPosting(0).iAttTimeType = rst_Cptt!attTimeType
            tgCPPosting(0).iVefCode = imVefCode
            tgCPPosting(0).iShttCode = rst_Cptt!shttCode
            tgCPPosting(0).sZone = rst_Cptt!shttTimeZone
            tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
            tgCPPosting(0).sAstStatus = "R" 'rst_Cptt!cpttAstStatus
            ''Determine days to generate
            ''Find first and last day
            'ilFirstDay = -1
            'For ilDay = 1 To 7 Step 1
            '    If Mid(smSpotCopyDays, ilDay, 1) = "Y" Then
            '        If ilFirstDay = -1 Then
            '            ilFirstDay = ilDay
            '        End If
            '        ilLastDay = ilDay
            '    End If
            '    If Mid(smRegionDays, ilDay, 1) = "Y" Then
            '        If ilFirstDay = -1 Then
            '            ilFirstDay = ilDay
            '        End If
            '        ilLastDay = ilDay
            '    End If
            'Next ilDay
            imAdfCode = -1
            If gDateValue(slMoDate) = gDateValue(gObtainPrevMonday(slStartDate)) Then
                ilFirstDay = Weekday(slStartDate, vbMonday)
            Else
                ilFirstDay = 1
            End If
            If gDateValue(slMoDate) = gDateValue(gObtainPrevMonday(slEndDate)) Then
                ilLastDay = Weekday(slEndDate, vbMonday)
            Else
                ilLastDay = 7
            End If
            bgTaskBlocked = False
            sgTaskBlockedName = "Station Spot Builder"

            If (ilFirstDay = 1) And (ilLastDay = 7) Then
                igTimes = 1 'By Week
                ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True, , , , , , True)
            Else
                tgCPPosting(0).iNumberDays = ilLastDay - ilFirstDay + 1
                tgCPPosting(0).sDate = DateAdd("d", ilFirstDay - 1, tgCPPosting(0).sDate)
                igTimes = 3 'By Date
                ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True, , , , , , True)
            End If
            bgTaskBlocked = False
            sgTaskBlockedName = ""
            
            If ilRet Then
                'Update cptt.mkd and either update or create ent.mkd
                SQLQuery = "SELECT Count(*) as Total FROM abf_AST_Build_Queue WHERE (abfStatus = 'G' Or abfStatus = 'H') "
                SQLQuery = SQLQuery & "AND abfVefCode = " & imVefCode & " AND abfShttCode = " & rst_Cptt!shttCode
                SQLQuery = SQLQuery & " AND abfMondayDate = '" & Format(slMoDate, sgSQLDateForm) & "'"
                Set rst_abf = gSQLSelectCall(SQLQuery)
                If rst_abf!Total <= 0 Then
                    SQLQuery = "UPDATE cptt SET "
                    SQLQuery = SQLQuery & "cpttAstStatus = '" & "C" & "' "
                    SQLQuery = SQLQuery & "WHERE cpttCode = " & rst_Cptt!cpttCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand1:
                        gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mBuildStationSpots"
                        mBuildStationSpots = False
                        Exit Function
                    End If
                End If
                On Error GoTo ErrHand
            End If
            rst_Cptt.MoveNext
        Wend
        gClearASTInfo True
        slMoDate = DateAdd("d", 7, slMoDate)
    Loop While gDateValue(slMoDate) < llEndDate
    mBuildStationSpots = True
    Exit Function
ErrHand:
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mBuildStationSpots"
    mBuildStationSpots = False
    Exit Function
ErrHand1:
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mBuildStationSpots"
    mBuildStationSpots = False
End Function

Private Sub mUpdateTimeRecord()
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    '12/5/14: Replaced by Task Monitor (tmf).
    Exit Sub
    
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")

    slSQLQuery = "UPDATE site SET "
    slSQLQuery = slSQLQuery & "siteSSBDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "siteSSBTime = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
    slSQLQuery = slSQLQuery & "WHERE siteCode = " & 1
    'cnn.Execute slSQL_AlertClear, rdExecDirect
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mUpdateTimeRecord"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mUpdateTimeRecord"
    Exit Sub
ErrHand1:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "StationSpotBuilder.txt", "StationSpotBuilder-mUpdateTimeRecord"
    Return
End Sub
Public Sub gAllowedExportsImportsInMenu(blIsOn As Boolean, ilVendor As Vendors)
    '8156
End Sub
