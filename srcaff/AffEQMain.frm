VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Affiliate Export Queue"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "AffEQMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7575
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
      Height          =   6105
      Left            =   0
      Picture         =   "AffEQMain.frx":08CA
      ScaleHeight     =   6045
      ScaleWidth      =   7500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7560
      Begin VB.Timer tmcSetTime 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   5805
         Top             =   5505
      End
      Begin VB.Timer tmcRestartTask 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6195
         Top             =   4950
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5175
         Top             =   4995
      End
      Begin VB.CommandButton cmcMin 
         Caption         =   "Minimize"
         Height          =   330
         Left            =   3075
         TabIndex        =   2
         Top             =   4815
         Width           =   1380
      End
      Begin VB.CommandButton cmcStop 
         Caption         =   "Stop"
         Height          =   330
         Left            =   3075
         TabIndex        =   3
         Top             =   5340
         Width           =   1380
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   45
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   930
         Width           =   105
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdStatus 
         Height          =   3345
         Left            =   210
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   5900
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
Private lmEqtCode As Long

Dim imLastHourGGChecked As Integer

Private rst_Eht As ADODB.Recordset
Private rst_Evt As ADODB.Recordset
Private rst_Ect As ADODB.Recordset
Private rst_Eqt As ADODB.Recordset
Private rst_Ust As ADODB.Recordset

Private Const SEXPORTTYPEINDEX = 0
Private Const SEXPORTNAMEINDEX = 1
Private Const SVEHICLEINDEX = 2
Private Const SUSERINDEX = 3
Private Const STIMEREQUESTINDEX = 4
Private Const SPRORITYINDEX = 5
Private Const SEQTCODEINDEX = 6





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
        MsgBox "Only one copy of Export Queue can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of Export Queue path: " & App.Path & " from " & Trim$(gGetComputerName()), "ExportQueueLog.Txt", False
        End
    End If
    imFirstTime = True
    lmEqtCode = -1
    igExportSource = 2
    sgUserName = "ExportQueue"
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
    ilRet = MsgBox("Stop the Export Queue", vbQuestion + vbYesNo, "Stop Service")
    If ilRet = vbNo Then
        Cancel = 1
        imCancelled = False
        tmcRestartTask.Enabled = True
        Exit Sub
    End If
    imClosed = True
    imCancelled = True
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
    
    igGGFlag = 1
    imLastHourGGChecked = -1
    
    igDemoMode = False
    If InStr(sgCommand, "Demo") Then
        igDemoMode = True
    End If
    
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
    '6548 Dan M
    ReDim sgXDSSection(0 To 0) As String
    slXMLINIInputFile = gXmlIniPath(True)
    If LenB(slXMLINIInputFile) <> 0 Then
        ilRet = gSearchFile(slXMLINIInputFile, "[XDigital", True, 1, sgXDSSection())
    End If
    
    ilRet = mOpenPervasiveAPI()
    
    igShowMsgBox = False
    
    mGetGuideCode
    ilRet = gInitGlobals()
    mClearAnyProcessing
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
    gHandleError "AffExportQueueLog", "Form-mStartUp"
    imCancelled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    tmcSetTime.Enabled = False
    rst_Eht.Close
    rst_Evt.Close
    rst_Ect.Close
    rst_Eqt.Close
    rst_Ust.Close
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
    gUpdateTaskMonitor 0, "AEQ"
    mUpdateTimeRecord
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mStartUp
    gLogMsg "Export Queue path: " & App.Path & " from " & Trim$(gGetComputerName()), "ExportQueueLog.Txt", False
'    tmcTask.Interval = CInt(lmSleepTime)
'    tmcTask.Enabled = True
    sgTimeZone = Left$(gGetLocalTZName(), 1)
    tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
    tmcSetTime.Enabled = True
    mTaskLoop
End Sub



Private Sub mPopulate()
    Dim llRow As Long
    
    On Error GoTo ErrHand
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtPriority > 0 AND (eqtStatus = 'R' or eqtStatus = 'P') ORDER BY eqtPriority, eqtDateEntered, eqtTimeEntered"
    Set rst_Eqt = cnn.Execute(SQLQuery)
    Do While Not rst_Eqt.EOF
        '9435 this really doesn't do anything, because Marketron is actually a subset of "A"
        If rst_Eqt!eqtType = "1" And gIsWebVendor(22) Then
            'skip
        Else
            If llRow >= grdStatus.Rows Then
                grdStatus.AddItem ""
            End If
            grdStatus.Row = llRow
            Select Case rst_Eqt!eqtType
                Case "A"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Aff Logs"
                Case "C"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "C & C"
                Case "D"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "IDC"
                Case "I"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "ISCI"
                Case "R"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "ISCI C/R"
                Case "4"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "RCS 4"
                Case "5"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "RCS 5"
                Case "S"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "StarGuide"
                Case "W"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Compel"
                Case "X"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "X-Digital"
                Case "1"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Marketron"
                Case "2"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "Univision"
                Case "3"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "CSI Web"
                'iPump
                Case "P"
                    grdStatus.TextMatrix(llRow, SEXPORTTYPEINDEX) = "IPump"
            End Select
            SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & rst_Eqt!eqtEhtCode
            Set rst_Eht = cnn.Execute(SQLQuery)
            If Not rst_Eht.EOF Then
                grdStatus.TextMatrix(llRow, SEXPORTNAMEINDEX) = Trim$(rst_Eht!ehtExportName)
            Else
                grdStatus.TextMatrix(llRow, SEXPORTNAMEINDEX) = ""
            End If
            SQLQuery = "SELECT ustname, ustReportName, ustUserInitials FROM Ust Where ustCode = " & rst_Eqt!eqtUstCode
            Set rst_Ust = cnn.Execute(SQLQuery)
            If Not rst_Ust.EOF Then
                If Trim$(rst_Ust!ustUserInitials) <> "" Then
                    grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustUserInitials)
                Else
                    If Trim$(rst_Ust!ustReportName) <> "" Then
                        grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustReportName)
                    Else
                        grdStatus.TextMatrix(llRow, SUSERINDEX) = Trim$(rst_Ust!ustname)
                    End If
                End If
            End If
            grdStatus.TextMatrix(llRow, STIMEREQUESTINDEX) = Format(rst_Eqt!eqtDateEntered, sgShowDateForm) & " " & Format(rst_Eqt!eqtTimeEntered, sgShowTimeWSecForm)
            If (rst_Eqt!eqtPriority <= 0) Or (rst_Eht!ehtSubType = "C") Then
                grdStatus.TextMatrix(llRow, SPRORITYINDEX) = "Custom"
            ElseIf rst_Eqt!eqtStatus = "P" Then
                grdStatus.TextMatrix(llRow, SPRORITYINDEX) = "Processing"
            Else
                grdStatus.TextMatrix(llRow, SPRORITYINDEX) = rst_Eqt!eqtPriority
            End If
            SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtCode
            Set rst_Evt = cnn.Execute(SQLQuery)
            If Not rst_Evt.EOF Then
                grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = rst_Evt(0).Value
            Else
                grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = ""
            End If
            If rst_Eht!ehtStandardEhtCode > 0 Then
                SQLQuery = "SELECT Count(evtCode) FROM evt_Export_Vehicles WHERE evtEhtCode = " & rst_Eht!ehtStandardEhtCode
                Set rst_Evt = cnn.Execute(SQLQuery)
                If Not rst_Evt.EOF Then
                    grdStatus.TextMatrix(llRow, SVEHICLEINDEX) = grdStatus.TextMatrix(llRow, SVEHICLEINDEX) & " of " & rst_Evt(0).Value
                End If
            End If
            grdStatus.TextMatrix(llRow, SEQTCODEINDEX) = rst_Eqt!eqtCode
            llRow = llRow + 1
        End If
        rst_Eqt.MoveNext
    Loop
    mSetStatusGridColor
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ExportQueueLog.txt", "ExportQueue-mPopulate"
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
            DoEvents
            mCheckGG
            If igGGFlag = 0 Then
                Unload frmMain
                Exit Sub
            End If
            mPopulate
            DoEvents
            slDateTime = gNow()
            slNowDate = Format$(slDateTime, "m/d/yy")
            slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
            blAnyProcessing = False
            ilPriority = 32000
            llPriorityRow = -1
            For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
                If grdStatus.TextMatrix(llRow, SPRORITYINDEX) <> "" Then
                    If grdStatus.TextMatrix(llRow, SPRORITYINDEX) <> "Custom" Then
                        If grdStatus.TextMatrix(llRow, SPRORITYINDEX) = "Processing" Then
                            blAnyProcessing = True
                            Exit For
                        Else
                            If Val(grdStatus.TextMatrix(llRow, SPRORITYINDEX)) < ilPriority Then
                                llPriorityRow = llRow
                                ilPriority = Val(grdStatus.TextMatrix(llRow, SPRORITYINDEX))
                            End If
                        End If
                    End If
                End If
            Next llRow
            If (Not blAnyProcessing) And (llPriorityRow <> -1) Then
                mPopArrays
                llEqtCode = Val(grdStatus.TextMatrix(llPriorityRow, SEQTCODEINDEX))
                SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtCode = " & llEqtCode
                Set rst_Eqt = cnn.Execute(SQLQuery)
                
                grdStatus.TextMatrix(llPriorityRow, SPRORITYINDEX) = "Processing"
                DoEvents
                SQLQuery = "UPDATE eqt_Export_Queue SET "
                SQLQuery = SQLQuery & "eqtStatus = 'P'" & ", "
                SQLQuery = SQLQuery & "eqtDateStarted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "eqtTimeStarted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
                SQLQuery = SQLQuery & "WHERE eqtCode = " & llEqtCode
                'cnn.Execute slSQL_AlertClear, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gSetMousePointer grdStatus, grdStatus, vbDefault
                    gHandleError "ExportQueueLog.txt", "ExportQueue-mTaskLoop"
                    grdStatus.Redraw = True
                    Exit Sub
                End If
                gUpdateTaskMonitor 1, "AEQ"
                igExportSource = 2
                lgExportEhtCode = rst_Eqt!eqtEhtCode
                sgExporStartDate = Format(rst_Eqt!eqtStartDate, sgShowDateForm)
                sgExportEndDate = Format(rst_Eqt!eqtEndDate, sgShowDateForm)
                igExportDays = Val(rst_Eqt!eqtNumberDays)
                sgExportTypeChar = rst_Eqt!eqtType
                lgExportEqtCode = rst_Eqt!eqtCode
                igExportReturn = 0
                ilEqtStatus = 0
                ilRet = gPopAll()
                Select Case rst_Eqt!eqtType
                    Case "A"
                        'Determine which exports requested
                        blMarketron = False
                        blUnivision = False
                        blCumulus = False
                        blCSIWeb = False
                        SQLQuery = "SELECT * FROM ect_Export_Criteria WHERE ectEhtCode = " & lgExportEhtCode
                        Set rst_Ect = cnn.Execute(SQLQuery)
                        Do While Not rst_Ect.EOF
                            If Trim$(rst_Ect!ectLogType) = "M" Then
                                If Trim$(rst_Ect!ectFieldName) = "ckcMarketron" Then
                                    '9435
                                    If Not gIsWebVendor(22) Then
                                        If rst_Ect!ectFieldValue = vbChecked Then
                                            blMarketron = True
                                        End If
                                    End If
                                End If
                            End If
                            If Trim$(rst_Ect!ectLogType) = "U" Then
                                If Trim$(rst_Ect!ectFieldName) = "ckcUnivision" Then
                                    If rst_Ect!ectFieldValue = vbChecked Then
                                        blUnivision = True
                                    End If
                                End If
                            End If
                            If Trim$(rst_Ect!ectLogType) = "C" Then
                                If Trim$(rst_Ect!ectFieldName) = "ckcCumulus" Then
                                    If rst_Ect!ectFieldValue = vbChecked Then
                                        blCumulus = True
                                    End If
                                End If
                            End If
                            If Trim$(rst_Ect!ectLogType) = "W" Then
                                If Trim$(rst_Ect!ectFieldName) = "ckcCSIWeb" Then
                                    If rst_Ect!ectFieldValue = vbChecked Then
                                        blCSIWeb = True
                                    End If
                                End If
                            End If
                            rst_Ect.MoveNext
                        Loop
                        If blMarketron Then
                            igExportTypeNumber = 1      'Marketron
                            FrmExportMarketron.Show vbModal
                        End If
                        If blUnivision Then
                            igExportTypeNumber = 2      'Univision
                            frmExportSchdSpot.Show vbModal
                        End If
                        If blCumulus Or blCSIWeb Then
                            ilRet = gTestWebVersion()
                            If ilRet = -1 Then
                                igExportTypeNumber = 3      'CSI Web
                                If blCumulus And blCSIWeb Then
                                    sgWebExport = "B"
                                ElseIf blCumulus Then
                                    sgWebExport = "C"
                                Else
                                    sgWebExport = "W"
                                End If
                                frmWebExportSchdSpot.Show vbModal
                            Else
                                ilEqtStatus = 1
                            End If
                        End If
                    Case "C"
                        igExportTypeNumber = 6
                        frmExportCnCSpots.Show vbModal
                    Case "D"    'IDC
                        igExportTypeNumber = 7
                        FrmExportIDC.Show vbModal
                    Case "I"    'ISCI
                        igExportTypeNumber = 8
                        frmExportISCI.Show vbModal
                    Case "R"    'ISCI Cross Reference
                        igExportTypeNumber = 9
                        frmExportISCIXRef.Show vbModal
                    Case "4"    'RCS 4
                        igExportTypeNumber = 4
                        igRCSExportBy = 4
                        frmExportRCS.Show vbModal
                    Case "5"    'RCS 5
                        igExportTypeNumber = 5
                        igRCSExportBy = 5
                        frmExportRCS.Show vbModal
                    Case "S"    'StarGuide
                        igExportTypeNumber = 10
                        frmExportStarGuide.Show vbModal
                    Case "W"    'Wegener
                        igExportTypeNumber = 11
                        FrmExportWegener.Show vbModal
                    Case "X"    'X-Digital
                        igExportTypeNumber = 12
                        FrmExportXDigital.Show vbModal
                    Case "P"    'ipump
                        igExportTypeNumber = 13
                        FrmExportiPump.Show vbModal
                    Case "1"
                    Case "2"
                    Case "3"
                End Select
                If igExportReturn = 2 Then
                    sgTmfStatus = "E"
                End If
                gUpdateTaskMonitor 2, "AEQ"
                mSetStatus ilEqtStatus
                mSetLastExportDate ilEqtStatus
                mAdjustPriority
                '6394 clear values
                sgExportResultName = ""
                DoEvents
            End If
            ilTaskCount = 1
        Else
            ilTaskCount = ilTaskCount + 1
        End If
   Loop
   Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ExportQueueLog.txt", "ExportQueue-mTaskLoop"
    grdStatus.Redraw = True
    Exit Sub
'ErrHand1:
'    gSetMousePointer grdStatus, grdStatus, vbDefault
'    gHandleError "ExportQueueLog.txt", "ExportQueue-mTaskLoop"
'    grdStatus.Redraw = True
'    Return
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(SEQTCODEINDEX) = 0
    grdStatus.ColWidth(SEXPORTTYPEINDEX) = grdStatus.Width * 0.13
    grdStatus.ColWidth(SUSERINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(STIMEREQUESTINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(SPRORITYINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(SVEHICLEINDEX) = grdStatus.Width * 0.11

    grdStatus.ColWidth(SEXPORTNAMEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = SEXPORTTYPEINDEX To SPRORITYINDEX Step 1
        If ilCol <> SEXPORTNAMEINDEX Then
            grdStatus.ColWidth(SEXPORTNAMEINDEX) = grdStatus.ColWidth(SEXPORTNAMEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus

End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, SEXPORTTYPEINDEX) = "Type"
    grdStatus.TextMatrix(0, SEXPORTNAMEINDEX) = "Name"
    grdStatus.TextMatrix(0, SUSERINDEX) = "User"
    grdStatus.TextMatrix(0, STIMEREQUESTINDEX) = "Requested"
    grdStatus.TextMatrix(0, SPRORITYINDEX) = "Priority"
    grdStatus.TextMatrix(0, SVEHICLEINDEX) = "Vehicle"


End Sub

Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdStatus, True
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = SEXPORTTYPEINDEX To SPRORITYINDEX Step 1
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
    
         
    ReDim tgCifCpfInfo1(0 To 0) As CIFCPFINFO1
    ReDim tgCrfInfo1(0 To 0) As CRFINFO1

    SQLQuery = "SELECT spfUseCartNo, spfRemoteUsers, spfUsingFeatures2, spfUsingFeatures5, spfUsingFeatures9, spfSportInfo, spfUseProdSptScr"
    'SQLQuery = SQLQuery + " FROM SPF_Site_Options spf"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = cnn.Execute(SQLQuery)
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
    mCreateExportSpec

End Sub

Private Sub mCreateExportSpec()
    If ((Asc(sgSpfUsingFeatures5) And STATIONINTERFACE) = STATIONINTERFACE) Then
        ReDim tgSpecInfo(0 To 10) As SPECINFO
        tgSpecInfo(0).sName = "Aff Logs"
        tgSpecInfo(0).sType = "A"
        tgSpecInfo(0).sFullName = "Affiliate Logs"
        tgSpecInfo(1).sName = "C & C"
        tgSpecInfo(1).sType = "C"
        tgSpecInfo(1).sFullName = "Clearance and Compensation"
        tgSpecInfo(2).sName = "IDC"
        tgSpecInfo(2).sType = "D"
        tgSpecInfo(2).sFullName = "IDC"
        tgSpecInfo(3).sName = "ISCI"
        tgSpecInfo(3).sType = "I"
        tgSpecInfo(3).sFullName = "ISCI"
        tgSpecInfo(4).sName = "ISCI C/R"
        tgSpecInfo(4).sType = "R"
        tgSpecInfo(4).sFullName = "ISCI Cross Reference"
        tgSpecInfo(5).sName = "RCS 4"
        tgSpecInfo(5).sType = "4"
        tgSpecInfo(5).sFullName = "RCS 4 Digit Cart #'s"
        tgSpecInfo(6).sName = "RCS 5"
        tgSpecInfo(6).sType = "5"
        tgSpecInfo(6).sFullName = "RCS 5 Digit Cart #'s"
        tgSpecInfo(7).sName = "StarGd"
        tgSpecInfo(7).sType = "S"
        tgSpecInfo(7).sFullName = "StarGuide"
        tgSpecInfo(8).sName = "Compel"
        tgSpecInfo(8).sType = "W"
        tgSpecInfo(8).sFullName = "Wegener Compel"
        tgSpecInfo(9).sName = "X-Digital"
        tgSpecInfo(9).sType = "X"
        tgSpecInfo(9).sFullName = "X-Digital"
        tgSpecInfo(10).sName = "Wegener IPump"
        tgSpecInfo(10).sType = "P"
        tgSpecInfo(10).sFullName = "IPump"

    Else
        ReDim tgSpecInfo(0 To 9) As SPECINFO
        tgSpecInfo(0).sName = "C & C"
        tgSpecInfo(0).sType = "C"
        tgSpecInfo(0).sFullName = "Clearance and Compensation"
        tgSpecInfo(1).sName = "IDC"
        tgSpecInfo(1).sType = "D"
        tgSpecInfo(1).sFullName = "IDC"
        tgSpecInfo(2).sName = "ISCI"
        tgSpecInfo(2).sType = "I"
        tgSpecInfo(2).sFullName = "ISCI"
        tgSpecInfo(3).sName = "ISCI C/R"
        tgSpecInfo(3).sType = "R"
        tgSpecInfo(3).sFullName = "ISCI Cross Reference"
        tgSpecInfo(4).sName = "RCS 4"
        tgSpecInfo(4).sType = "4"
        tgSpecInfo(4).sFullName = "RCS 4 Digit Cart #'s"
        tgSpecInfo(5).sName = "RCS 5"
        tgSpecInfo(5).sType = "5"
        tgSpecInfo(5).sFullName = "RCS 5 Digit Cart #'s"
        tgSpecInfo(6).sName = "StarGd"
        tgSpecInfo(6).sType = "S"
        tgSpecInfo(6).sFullName = "StarGuide"
        tgSpecInfo(7).sName = "Compel"
        tgSpecInfo(7).sType = "W"
        tgSpecInfo(7).sFullName = "Wegener Compel"
        tgSpecInfo(8).sName = "X-Digital"
        tgSpecInfo(8).sType = "X"
        tgSpecInfo(8).sFullName = "X-Digital"
        tgSpecInfo(9).sName = "IPump"
        tgSpecInfo(9).sType = "P"
        tgSpecInfo(9).sFullName = "Wegener IPump"

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

Private Sub mUpdateTimeRecord()
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    
    On Error GoTo ErrHand
    '12/5/14: Replaced by Task Monitor (tmf)
    Exit Sub
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "m/d/yy")
    slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
    If lmEqtCode <= 0 Then
        SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtType = 'T'"
        Set rst_Eqt = cnn.Execute(SQLQuery)
        If rst_Eqt.EOF Then
            SQLQuery = "Insert Into eqt_Export_Queue ( "
            SQLQuery = SQLQuery & "eqtCode, "
            SQLQuery = SQLQuery & "eqtEhtCode, "
            SQLQuery = SQLQuery & "eqtPriority, "
            SQLQuery = SQLQuery & "eqtDateEntered, "
            SQLQuery = SQLQuery & "eqtTimeEntered, "
            SQLQuery = SQLQuery & "eqtStatus, "
            SQLQuery = SQLQuery & "eqtDateStarted, "
            SQLQuery = SQLQuery & "eqtTimeStarted, "
            SQLQuery = SQLQuery & "eqtDateCompleted, "
            SQLQuery = SQLQuery & "eqtTimeCompleted, "
            SQLQuery = SQLQuery & "eqtUstCode, "
            SQLQuery = SQLQuery & "eqtResultFile, "
            SQLQuery = SQLQuery & "eqtType, "
            SQLQuery = SQLQuery & "eqtStartDate, "
            SQLQuery = SQLQuery & "eqtNumberDays, "
            SQLQuery = SQLQuery & "eqtEndDate, "
            SQLQuery = SQLQuery & "eqtProcesingVefCode, "
            SQLQuery = SQLQuery & "eqtToBeProcessed, "
            SQLQuery = SQLQuery & "eqtBeenProcessed, "
            SQLQuery = SQLQuery & "eqtUnused "
            SQLQuery = SQLQuery & ") "
            SQLQuery = SQLQuery & "Values ( "
            SQLQuery = SQLQuery & "Replace" & ", "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Format$(slNowDate, sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$(slNowTime, sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & "C" & "', "
            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & igUstCode & ", "
            SQLQuery = SQLQuery & "'" & "" & "', "
            SQLQuery = SQLQuery & "'" & "T" & "', "
            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & 0 & ", "
            SQLQuery = SQLQuery & "'" & "" & "' "
            SQLQuery = SQLQuery & ") "
            lmEqtCode = gInsertAndReturnCode(SQLQuery, "eqt_export_queue", "eqtCode", "Replace")
            If lmEqtCode <= 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                gSetMousePointer grdStatus, grdStatus, vbDefault
                gHandleError "ExportQueueLog.txt", "ExportQueue-mUpdateTimeRecord"
                Exit Sub
            End If
            On Error GoTo ErrHand
            Exit Sub
        Else
            lmEqtCode = rst_Eqt!eqtCode
        End If
    End If
    SQLQuery = "UPDATE eqt_Export_Queue SET "
    SQLQuery = SQLQuery & "eqtDateEntered = '" & Format$(slNowDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "eqtTimeEntered = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
    SQLQuery = SQLQuery & "WHERE eqtCode = " & lmEqtCode
    'cnn.Execute slSQL_AlertClear, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gSetMousePointer grdStatus, grdStatus, vbDefault
        gHandleError "ExportQueueLog.txt", "ExportQueue-mUpdateTimeRecord"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ExportQueueLog.txt", "ExportQueue-mUpdateTimeRecord"
    Exit Sub
'ErrHand1:
'    gSetMousePointer grdStatus, grdStatus, vbDefault
'    gHandleError "ExportQueueLog.txt", "ExportQueue-mUpdateTimeRecord"
'    Return
End Sub

Private Sub mGetGuideCode()
    On Error GoTo ErrHand
    SQLQuery = "SELECT ustCode FROM ust WHERE ustName = 'Guide'"
    Set rst = cnn.Execute(SQLQuery)
    If rst.EOF Then
        igUstCode = 1
    Else
        igUstCode = rst!ustCode
    End If
    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mGetGuideCode"
End Sub

Private Sub mAdjustPriority()
    Dim ilMin As Integer
    On Error GoTo ErrHand
    SQLQuery = "SELECT Min(eqtPriority) FROM eqt_Export_Queue WHERE (eqtPriority > 0) AND (eqtStatus = 'R' or eqtStatus = 'P')"
    Set rst_Eqt = cnn.Execute(SQLQuery)
    If Not rst_Eqt.EOF Then
        If rst_Eqt(0).Value <> vbNull Then
            ilMin = rst_Eqt(0).Value
            If ilMin > 1 Then
                ilMin = ilMin - 1
                SQLQuery = "UPDATE eqt_Export_Queue SET "
                SQLQuery = SQLQuery & "eqtPriority = eqtPriority - " & ilMin
                SQLQuery = SQLQuery & "WHERE eqtPriority > 1 AND eqtStatus = 'R'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gHandleError "ExportQueueLog.txt", "ExportQueue-mAdjustPriority"
                    Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mAdjustPriority"
    Exit Sub
'ErrHand1:
'    gHandleError "ExportQueueLog.txt", "ExportQueue-mAdjustPriority"
'
End Sub

Private Sub mSetLastExportDate(ilEqtStatus As Integer)
    'ilEqtStatus: 0=Update; 1=Bypass as EQT removed
    Dim blUpdateDate As Boolean
    Dim llEhtCode As Long
    Dim slLDE As String
    Dim llStandardEhtCode As Long
    On Error GoTo ErrHand
    If ilEqtStatus = 1 Then
        Exit Sub
    End If
    'Check each partial if Exported
    SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & lgExportEhtCode
    Set rst_Eht = cnn.Execute(SQLQuery)
    If Not rst_Eht.EOF Then
        If rst_Eht!ehtStandardEhtCode = 0 Then
            blUpdateDate = True
            llEhtCode = lgExportEhtCode
        Else
            'Check if any more partial exist
            llStandardEhtCode = rst_Eht!ehtStandardEhtCode
            llEhtCode = llStandardEhtCode
            slLDE = Format$(rst_Eht!ehtLDE, sgSQLDateForm)
            SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtStandardEhtCode = " & llStandardEhtCode & " AND ehtCode <> " & lgExportEhtCode & " AND ehtLDE = '" & slLDE & "'"
            Set rst_Eht = cnn.Execute(SQLQuery)
            If rst_Eht.EOF Then
                blUpdateDate = False
            Else
                blUpdateDate = True
                Do While Not rst_Eht.EOF
                    'Note: Only one ehtCode exist for each unique partial, therefore no other match is required
                    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtEhtCode = " & rst_Eht!ehtCode & " AND (eqtStatus = 'C' or eqtStatus = 'E') "
                    Set rst_Eqt = cnn.Execute(SQLQuery)
                    If rst_Eqt.EOF Then
                        blUpdateDate = False
                        Exit Do
                    End If
                    rst_Eht.MoveNext
                Loop
            End If
        End If
        If blUpdateDate Then
            SQLQuery = "UPDATE eht_Export_Header SET "
            SQLQuery = SQLQuery & "ehtLDE = '" & Format$(sgExportEndDate, sgSQLDateForm) & "' "
            SQLQuery = SQLQuery & "WHERE ehtCode = " & llEhtCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
                Exit Sub
            End If
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
    Exit Sub
'ErrHand1:
'    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
End Sub

Private Sub mSetStatus(ilEqtStatus As Integer)
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    
    On Error GoTo ErrHand
    'If ilEqtStatus = 1 Then
    '    SQLQuery = "Delete From eqt_Export_Queue Where eqtCode = " & lgExportEqtCode
    'Else
        slDateTime = gNow()
        slNowDate = Format$(slDateTime, "m/d/yy")
        slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
        SQLQuery = "UPDATE eqt_Export_Queue SET "
        If (igExportReturn = 2) Or (ilEqtStatus = 1) Then
            sgTmfStatus = "E"
            SQLQuery = SQLQuery & "eqtStatus = 'E'" & ", "
        Else
            SQLQuery = SQLQuery & "eqtStatus = 'C'" & ", "
        End If
        SQLQuery = SQLQuery & "eqtResultFile = '" & sgExportResultName & "',"
        SQLQuery = SQLQuery & "eqtDateCompleted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "eqtTimeCompleted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
        SQLQuery = SQLQuery & "WHERE eqtCode = " & lgExportEqtCode
    'End If
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
        Exit Sub
    End If

    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
    Exit Sub
'ErrHand1:
'    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
End Sub

Private Sub mClearAnyProcessing()
    Dim slExportName As String
    Dim slStartDate As String
    
    SQLQuery = "SELECT * FROM eqt_Export_Queue WHERE eqtPriority > 0 AND eqtStatus = 'P'"
    Set rst_Eqt = cnn.Execute(SQLQuery)
    Do While Not rst_Eqt.EOF
        SQLQuery = "SELECT * FROM eht_Export_Header WHERE ehtCode = " & rst_Eqt!eqtEhtCode
        Set rst_Eht = cnn.Execute(SQLQuery)
        If Not rst_Eht.EOF Then
            slExportName = Trim$(rst_Eht!ehtExportName)
        Else
            slExportName = "Missing EhtCode: " & rst_Eqt!eqtEhtCode
        End If
        slStartDate = Format$(rst_Eqt!eqtStartDate, sgShowDateForm)
        gLogMsg "Following Export Terminated: " & slExportName & " Start Date " & slStartDate, "ExportQueueLog.Txt", False
        lgExportEqtCode = rst_Eqt!eqtCode
        igExportReturn = 2
        sgExportResultName = Trim$(rst_Eqt!eqtResultFile)
        mSetStatus 0
        rst_Eqt.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
    Exit Sub
ErrHand1:
    gHandleError "ExportQueueLog.txt", "ExportQueue-mSetLastExportDate"
End Sub

Private Sub mCheckGG()
    Dim c As Integer
    Dim slName As String
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim slStr As String
    Dim llDate As Long
    Dim llNow As Long
    
    Dim gg_rst As ADODB.Recordset
    If imLastHourGGChecked = Hour(Now) Then
        Exit Sub
    End If
    imLastHourGGChecked = Hour(Now)
    SQLQuery = "Select safName From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set gg_rst = cnn.Execute(SQLQuery)
    If Not gg_rst.EOF Then
        slName = Trim$(gg_rst!safName)
        ilField1 = Asc(slName)
        slStr = Mid$(slName, 2, 5)
        llDate = Val(slStr)
        llNow = gDateValue(Format$(Now, "m/d/yy"))
        ilField2 = Asc(Mid$(slName, 11, 1))
        If (ilField1 = 0) And (ilField2 = 1) Then
            If llDate <= llNow Then
                ilField2 = 0
            End If
        End If
        If (ilField1 = 0) And (ilField2 = 0) Then
            igGGFlag = 0
        End If
    End If
    gg_rst.Close
End Sub
Public Sub gAllowedExportsImportsInMenu(blIsOn As Boolean, ilVendor As Vendors)
    '8156
End Sub
