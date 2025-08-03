VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Affiliate Report Queue"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "AffRQMain.frx":0000
   LinkTopic       =   "Form1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.PictureBox plcSignon 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      Picture         =   "AffRQMain.frx":08CA
      ScaleHeight     =   6315
      ScaleWidth      =   10155
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10215
      Begin VB.Timer tmcSetTime 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   315
         Top             =   4635
      End
      Begin VB.Timer tmcRestartTask 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   390
         Top             =   5580
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   315
         Top             =   5145
      End
      Begin VB.CommandButton cmcMin 
         Caption         =   "Minimize"
         Height          =   330
         Left            =   3255
         TabIndex        =   2
         Top             =   5580
         Width           =   1380
      End
      Begin VB.CommandButton cmcStop 
         Caption         =   "Stop"
         Height          =   330
         Left            =   5715
         TabIndex        =   3
         Top             =   5580
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
         Height          =   3855
         Left            =   210
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   6800
         _Version        =   393216
         Rows            =   4
         Cols            =   11
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
         _Band(0).Cols   =   11
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
Private lmRqtCode As Long

Dim imLastHourGGChecked As Integer

Private imLastColSorted As Integer
Private imLastSort As Integer

Private rst_Rct As ADODB.Recordset
Private rst_rqt As ADODB.Recordset
Private rst_Ust As ADODB.Recordset

Private Const REPORTNAMEINDEX = 0
Private Const USERNAMEINDEX = 1
Private Const PRIORITYINDEX = 2
Private Const STATUEINDEX = 3
Private Const REQUESTEDINDEX = 4
Private Const STARTEDINDEX = 5
Private Const COMPLETEDINDEX = 6
Private Const DESCRIPTIONINDEX = 7
Private Const DELETEINDEX = 8
Private Const SORTINDEX = 9
Private Const RQTCODEINDEX = 10





Private Sub cmcMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmcStop_Click()
    imCancelled = True
End Sub

Private Sub grdStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdStatus.ToolTipText = ""
    If (grdStatus.MouseRow >= grdStatus.FixedRows) And (grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)) <> "" Then
        If grdStatus.MouseCol = REPORTNAMEINDEX Then
            grdStatus.ToolTipText = Trim$(grdStatus.TextMatrix(grdStatus.MouseRow, DESCRIPTIONINDEX))
        Else
            grdStatus.ToolTipText = grdStatus.TextMatrix(grdStatus.MouseRow, grdStatus.MouseCol)
        End If
    End If
End Sub

Private Sub MDIForm_Activate()
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
        mPopulate
        imFirstTime = False
    End If

End Sub

Private Sub MDIForm_Load()
    Dim ilPos As Integer
    
    sgCommand = Command$
    If App.PrevInstance Then
        MsgBox "Only one copy of Report Queue can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of Report Queue path: " & App.Path & " from " & Trim$(gGetComputerName()), "ReportQueueLog.Txt", False
        End
    End If
    imFirstTime = True
    lmRqtCode = -1
    igReportSource = 2
    sgUserName = "ReportQueue"
    imLastColSorted = -1
    imLastSort = -1
    igDemoMode = True   'Unused to avoid testing if web available
    gCenterStdAlone Me
    tmcStart.Enabled = True
End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRet As Integer
    
    If imClosed = True Then
        Exit Sub
    End If
    tmcRestartTask.Enabled = False
    tmcStart.Enabled = False
    tmcSetTime.Enabled = False
    ilRet = MsgBox("Stop the Report Queue", vbQuestion + vbYesNo, "Stop Service")
    If ilRet = vbNo Then
        Cancel = 1
        imCancelled = False
        tmcRestartTask.Enabled = True
        tmcSetTime.Enabled = True
        Exit Sub
    End If
    imClosed = True
    imCancelled = True
End Sub

Private Sub MDIForm_Resize()
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
    
    tmcStart.Enabled = False
    
    igGGFlag = 1
    igRptGGFlag = 1
    imLastHourGGChecked = -1
        
    'To avoid Web checking, leave igDemoMode as True
    'igDemoMode = False
    'If InStr(sgCommand, "Demo") Then
    '    igDemoMode = True
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
    bgReportQueue = False
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
    
    
    ilRet = mOpenPervasiveAPI()
    
    igShowMsgBox = False
    
    'tmcSetTime.Enabled = True
    mGetGuideCode
    ilRet = gInitGlobals()
    gPopAll
    mClearAnyProcessing
    'mPopulate
    Exit Sub
    
mReadFileErr:
    ilRet = Err.Number
    Resume Next
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffReportQueueLog", "Form-mStartUp"
    imCancelled = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    
    rst_Rct.Close
    rst_rqt.Close
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
    gUpdateTaskMonitor 0, "ARQ"
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    
    mStartUp
    gLogMsg "Report Queue path: " & App.Path & " from " & Trim$(gGetComputerName()), "ReportQueueLog.Txt", False
'    tmcTask.Interval = CInt(lmSleepTime)
'    tmcTask.Enabled = True
    sgTimeZone = Left$(gGetLocalTZName(), 1)
    tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
    tmcSetTime.Enabled = True
    mTaskLoop
End Sub



Private Sub mPopulate()
    Dim llRow As Long
    Dim slStr As String
    
    On Error GoTo ErrHand
    gSetMousePointer grdStatus, grdStatus, vbHourglass
    gGrid_Clear grdStatus, True
    llRow = grdStatus.FixedRows
    grdStatus.Redraw = False
    
    SQLQuery = "Select "
    SQLQuery = SQLQuery & "* "
    SQLQuery = SQLQuery & "From rqt "
    SQLQuery = SQLQuery & "Where "
    SQLQuery = SQLQuery & " rqtDateEntered >= '" & Format(gNow(), sgSQLDateForm) & "'"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    Do While Not rst_rqt.EOF
        If llRow >= grdStatus.Rows Then
            grdStatus.AddItem ""
        End If
        grdStatus.Row = llRow
    
        grdStatus.TextMatrix(llRow, REPORTNAMEINDEX) = Trim$(rst_rqt!rqtReportName)
        SQLQuery = "SELECT ustname, ustReportName, ustUserInitials FROM Ust Where ustCode = " & rst_rqt!rqtUstCode
        Set rst_Ust = gSQLSelectCall(SQLQuery)
        If Not rst_Ust.EOF Then
            'If Trim$(rst_Ust!ustUserInitials) <> "" Then
            '    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustUserInitials)
            'Else
                If Trim$(rst_Ust!ustReportName) <> "" Then
                    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustReportName)
                Else
                    grdStatus.TextMatrix(llRow, USERNAMEINDEX) = Trim$(rst_Ust!ustname)
                End If
            'End If
        End If
        If rst_rqt!rqtStatus = "P" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = "Running"
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
        ElseIf rst_rqt!rqtStatus = "C" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = ""
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
            grdStatus.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_rqt!rqtDateCompleted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeCompleted, sgShowTimeWSecForm)
        ElseIf rst_rqt!rqtStatus = "E" Then
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Error"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = ""
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
            If gDateValue(Format(rst_rqt!rqtDateStarted, sgShowDateForm)) <> gDateValue("1/1/1970") Then
                grdStatus.TextMatrix(llRow, STARTEDINDEX) = Format(rst_rqt!rqtDateStarted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeStarted, sgShowTimeWSecForm)
            End If
            If gDateValue(Format(rst_rqt!rqtDateCompleted, sgShowDateForm)) <> gDateValue("1/1/1970") Then
                grdStatus.TextMatrix(llRow, COMPLETEDINDEX) = Format(rst_rqt!rqtDateCompleted, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeCompleted, sgShowTimeWSecForm)
            End If
        Else
            grdStatus.TextMatrix(llRow, STATUEINDEX) = "Requested"
            grdStatus.TextMatrix(llRow, PRIORITYINDEX) = rst_rqt!rqtCode
            grdStatus.TextMatrix(llRow, REQUESTEDINDEX) = Format(rst_rqt!rqtDateEntered, sgShowDateForm) & " " & Format(rst_rqt!rqtTimeEntered, sgShowTimeWSecForm)
        End If
        grdStatus.TextMatrix(llRow, DESCRIPTIONINDEX) = rst_rqt!rqtDescription
        grdStatus.TextMatrix(llRow, RQTCODEINDEX) = rst_rqt!rqtCode
        llRow = llRow + 1
        rst_rqt.MoveNext
    Loop
    mReportSortCol PRIORITYINDEX
    mSetStatusGridColor
    gSetMousePointer grdStatus, grdStatus, vbDefault
    grdStatus.Redraw = True
    Exit Sub
ErrHand:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ReportQueueLog.txt", "ReportQueue-mPopulate"
    grdStatus.Redraw = True
    Resume Next
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
    Dim llRqtCode As Long
    Dim blMarketron As Boolean
    Dim blUnivision As Boolean
    Dim blCumulus As Boolean
    Dim blCSIWeb As Boolean
    Dim ilRqtStatus As Integer '0=Update, 1=Set as in Error and don't update LDE it was defind to mean Delete
    
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
            mCheckGG
            If (igGGFlag = 0) And (igRptGGFlag = 0) Then
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
                If grdStatus.TextMatrix(llRow, PRIORITYINDEX) <> "" Then
                    If grdStatus.TextMatrix(llRow, PRIORITYINDEX) <> "Custom" Then
                        If grdStatus.TextMatrix(llRow, PRIORITYINDEX) = "Processing" Then
                            blAnyProcessing = True
                            Exit For
                        Else
                            If Val(grdStatus.TextMatrix(llRow, PRIORITYINDEX)) < ilPriority Then
                                llPriorityRow = llRow
                                ilPriority = Val(grdStatus.TextMatrix(llRow, PRIORITYINDEX))
                            End If
                        End If
                    End If
                End If
            Next llRow
            If (Not blAnyProcessing) And (llPriorityRow <> -1) Then
                mPopArrays
                llRqtCode = Val(grdStatus.TextMatrix(llPriorityRow, RQTCODEINDEX))
                SQLQuery = "SELECT * FROM rqt WHERE rqtCode = " & llRqtCode
                Set rst_rqt = gSQLSelectCall(SQLQuery)

                grdStatus.TextMatrix(llPriorityRow, PRIORITYINDEX) = "Processing"
                DoEvents
                slDateTime = gNow()
                slNowDate = Format$(slDateTime, "m/d/yy")
                slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
                SQLQuery = "UPDATE rqt SET "
                SQLQuery = SQLQuery & "rqtStatus = 'P'" & ", "
                SQLQuery = SQLQuery & "rqtDateStarted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "rqtTimeStarted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
                SQLQuery = SQLQuery & "WHERE rqtCode = " & llRqtCode
                'cnn.Execute slSQL_AlertClear, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gSetMousePointer grdStatus, grdStatus, vbDefault
                    gHandleError "ReportQueueLog.txt", "ReportQueue-mTaskLoop"
                    grdStatus.Redraw = True
                    Exit Sub
                End If
                gUpdateTaskMonitor 1, "ARQ"
                igReportSource = 2
                igReportReturn = 0
                lgReportRqtCode = llRqtCode
                ilRet = mStartReportForm(rst_rqt!rqtReportName, llRqtCode)
                'mSetStatus ilRqtStatus
                gUpdateTaskMonitor 2, "ARQ"
                mAdjustPriority
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
    gHandleError "ReportQueueLog.txt", "ReportQueue-mTaskLoop"
    grdStatus.Redraw = True
    Resume Next
    Exit Sub
ErrHand1:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ReportQueueLog.txt", "ReportQueue-mTaskLoop"
    grdStatus.Redraw = True
    Return
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdStatus.ColWidth(RQTCODEINDEX) = 0
    grdStatus.ColWidth(DESCRIPTIONINDEX) = 0
    grdStatus.ColWidth(SORTINDEX) = 0
    grdStatus.ColWidth(USERNAMEINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(PRIORITYINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(STATUEINDEX) = grdStatus.Width * 0.1
    grdStatus.ColWidth(REQUESTEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(STARTEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(COMPLETEDINDEX) = grdStatus.Width * 0.15
    grdStatus.ColWidth(DELETEINDEX) = grdStatus.Width * 0.02
    grdStatus.ColWidth(REPORTNAMEINDEX) = grdStatus.Width - GRIDSCROLLWIDTH - 15
    For ilCol = REPORTNAMEINDEX To DELETEINDEX Step 1
        If ilCol <> REPORTNAMEINDEX Then
            grdStatus.ColWidth(REPORTNAMEINDEX) = grdStatus.ColWidth(REPORTNAMEINDEX) - grdStatus.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdStatus

End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdStatus.TextMatrix(0, REPORTNAMEINDEX) = "Report Name"
    grdStatus.TextMatrix(0, USERNAMEINDEX) = "User"
    grdStatus.TextMatrix(0, PRIORITYINDEX) = "Priority"
    grdStatus.TextMatrix(0, STATUEINDEX) = "Status"
    grdStatus.TextMatrix(0, REQUESTEDINDEX) = "Requested"
    grdStatus.TextMatrix(0, STARTEDINDEX) = "Started"
    grdStatus.TextMatrix(0, COMPLETEDINDEX) = "Completed"

End Sub

Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    'gGrid_Clear grdStatus, True
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        For llCol = REPORTNAMEINDEX To COMPLETEDINDEX Step 1
            grdStatus.Row = llRow
            grdStatus.Col = llCol
            If llCol = PRIORITYINDEX Then
                If grdStatus.TextMatrix(llRow, llCol) = "Requeted" Then
                    grdStatus.CellBackColor = vbWhite
                Else
                    grdStatus.CellBackColor = LIGHTYELLOW
                End If
            ElseIf llCol = DELETEINDEX Then
                If grdStatus.TextMatrix(llRow, llCol) = "Requeted" Then
                    grdStatus.CellBackColor = vbRed
                    grdStatus.CellForeColor = vbWhite
                    grdStatus.TextMatrix(llRow, DELETEINDEX) = "X"
                Else
                    grdStatus.CellBackColor = LIGHTYELLOW
                End If
            Else
                grdStatus.CellBackColor = LIGHTYELLOW
            End If
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
    gHandleError "ReportQueueLog.txt", "ReportQueue-mGetGuideCode"
End Sub

Private Sub mAdjustPriority()
    Dim ilMin As Integer
    On Error GoTo ErrHand
    SQLQuery = "SELECT Min(rqtPriority) FROM rqt WHERE (rqtPriority > 0) AND (rqtStatus = 'R' or rqtStatus = 'P')"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    If Not rst_rqt.EOF Then
        If rst_rqt(0).Value <> vbNull Then
            ilMin = rst_rqt(0).Value
            If ilMin > 1 Then
                ilMin = ilMin - 1
                SQLQuery = "UPDATE rqt SET "
                SQLQuery = SQLQuery & "rqtPriority = rqtPriority - " & ilMin
                SQLQuery = SQLQuery & "WHERE rqtPriority > 1 AND rqtStatus = 'R'"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Rplaced GoSo
                    'GoSub ErrHand1:
                    gHandleError "ReportQueueLog.txt", "ReportQueue-mAdjustPriority"
                    Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
ErrHand:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mAdjustPriority"
    Exit Sub
ErrHand1:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mAdjustPriority"
    
End Sub



Private Sub mSetStatus(ilRqtStatus As Integer)
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    
    On Error GoTo ErrHand
    If ilRqtStatus = 1 Then
        SQLQuery = "Delete From rqt Where rqtCode = " & lmRqtCode
    Else
        slDateTime = gNow()
        slNowDate = Format$(slDateTime, "m/d/yy")
        slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
        SQLQuery = "UPDATE rqt SET "
        If (igReportReturn = 2) Or (ilRqtStatus = 1) Then
            sgTmfStatus = "E"
            SQLQuery = SQLQuery & "rqtStatus = 'E'" & ", "
        Else
            SQLQuery = SQLQuery & "rqtStatus = 'C'" & ", "
        End If
        'SQLQuery = SQLQuery & "eqtResultFile = '" & sgExportResultName & "',"
        SQLQuery = SQLQuery & "rqtDateCompleted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "rqtTimeCompleted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
        SQLQuery = SQLQuery & "WHERE rqtCode = " & lmRqtCode
    End If
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "ReportQueueLog.txt", "ReportQueue-mSetStatus"
        Exit Sub
    End If

    Exit Sub
ErrHand:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mSetStatus"
    Exit Sub
ErrHand1:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mSetStatus"
End Sub

Private Sub mClearAnyProcessing()
    Dim slReportName As String
    Dim slRequested As String
    
    SQLQuery = "SELECT * FROM rqt WHERE rqtPriority > 0 AND rqtStatus = 'P'"
    Set rst_rqt = gSQLSelectCall(SQLQuery)
    Do While Not rst_rqt.EOF
        slReportName = Trim$(rst_rqt!rqtReportName)
        slRequested = Format$(rst_rqt!rqtDateEntered, sgShowDateForm)
        gLogMsg "Following Report Terminated: " & slReportName & " Requested Date " & slRequested, "ReportQueueLog.Txt", False
        lmRqtCode = rst_rqt!rqtCode
        igReportReturn = 2
        'sgExportResultName = Trim$(rst_Eqt!eqtResultFile)
        mSetStatus 0
        rst_rqt.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mClearAnyProcessing"
    Exit Sub
ErrHand1:
    gHandleError "ReportQueueLog.txt", "ReportQueue-mClearAnyProcessing"
End Sub

Private Sub mReportSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    
    For llRow = grdStatus.FixedRows To grdStatus.Rows - 1 Step 1
        slStr = Trim$(grdStatus.TextMatrix(llRow, REPORTNAMEINDEX))
        If slStr <> "" Then
            slStr = Trim$(grdStatus.TextMatrix(llRow, REQUESTEDINDEX))
            ilPos = InStr(1, slStr, " ", vbTextCompare)
            If ilPos > 0 Then
                slDate = Left(slStr, ilPos - 1)
                slTime = Mid(slStr, ilPos + 1)
            Else
                slDate = slStr
                slTime = "11:59:59pm"
            End If
            slDate = gDateValue(slDate)
            Do While Len(slDate) < 6
                slDate = "0" & slDate
            Loop
            slTime = Trim$(Str$(gTimeToLong(slTime, False)))
            Do While Len(slTime) < 6
                slTime = "0" & slTime
            Loop
            If ilCol = PRIORITYINDEX Then
                If grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing" Then
                    slSort = "A"
                ElseIf grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed" Then
                    slStr = "C" & slDate & slTime
                Else
                    slSort = (Trim$(grdStatus.TextMatrix(llRow, PRIORITYINDEX)))
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                    slSort = "B" & slSort
                End If
            ElseIf ilCol = REQUESTEDINDEX Then
                slSort = slDate & slTime
            ElseIf ilCol = STATUEINDEX Then
                If grdStatus.TextMatrix(llRow, STATUEINDEX) = "Processing" Then
                    slSort = "P"
                ElseIf grdStatus.TextMatrix(llRow, STATUEINDEX) = "Completed" Then
                    slStr = "C" & slDate & slTime
                Else
                    slSort = (Trim$(grdStatus.TextMatrix(llRow, PRIORITYINDEX)))
                    Do While Len(slSort) < 3
                        slSort = "0" & slSort
                    Loop
                    slSort = "R" & slSort
                End If
            Else
                slSort = UCase$(Trim$(grdStatus.TextMatrix(llRow, ilCol)))
                If slSort = "" Then
                    slSort = " "
                End If
            End If
            slStr = grdStatus.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStatus.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdStatus.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdStatus, REPORTNAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Function mStartReportForm(slInReportName As String, llRqtCode As Long) As Integer
    Dim ilRet As Integer
    Dim slReportName As String
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    
    On Error GoTo mStartReportFormErr
    igReportModelessStatus = -1
    slReportName = Trim$(slInReportName)
    sgReportListName = Trim$(slReportName)
    ilCount = 30
    Select Case Trim$(slReportName)
        Case "Station Information"
            'frmStationRpt.Show
        Case "Affiliate Agreements"
            'frmVehAffRpt.Show
        Case "Overdue C.P.s"
            frmDelqRpt.Show
        Case "Mailing Labels"
            'frmLabelRpt.Show
        Case "Advertiser Clearances"
            igReportModelessStatus = 0
            frmClearRpt.Show vbModeless
        Case "Pledges"
            'frmPledgeRpt.Show
        Case "Spot Clearance"
            'frmAiredRpt.Show
        Case "Affiliate Affidavit Posting Activity"
            'frmPostActivityRpt.Show
        Case "Web Log Activity"
            'frmLogActivityRpt.Show
        Case "Web Log Inactivity"
            'frmLogInactivityRpt.Show
        Case "Alert Status"
            'frmAlertRpt.Show
        Case "Affiliate Clearance Counts"
            'frmAffiliateRpt.Show
        Case "Program Clearance"
            'frmPgmClrRpt.Show
        Case "Pledged vs Aired Clearance"
            'frmPldgAirRpt.Show
        Case "Fed vs Aired Clearance"
            'frmPldgAirRpt.Show
        Case "Feed Verification"
            'frmVerifyRpt.Show
        Case "Export Journal"
            'frmJournalRpt.Show
        Case "Export Monitoring"
            'frmExpMonRpt.Show
        Case 18
            'frmMarkAssignRpt.Show
        Case "Non-Compliant(NCR)"
            'frmDelqRpt.Show
        Case "User Options"
            'frmUserOptionsRpt.Show
        Case "Site Options"
            'frmRptNoSel.Show
        Case "Affiliates Missing Weeks"
            'frmDelqRpt.Show
        Case "Regional Affiliate Copy Assignment", "Regional Affiliate Copy Tracing"
            'frmRgAssignRpt.Show
        Case "Groups"
            'frmGroupRpt.Show
        Case "Advertiser Fulfillment"
            'frmAdvFulFillRpt.Show
        Case "Contact Comments"
            'frmCommentRpt.Show
        Case "Web Import Log"
            'frmWebLogImportRpt.Show
        Case "Affiliate Log Delivery"           '2-21-12
            'frmLogDeliveryRpt.Show
        Case "Affiliate Spot Management"               '3-9-12
            'FrmSpotMgmtRpt.Show
        Case "Export History"             '6-7-12
            'frmExpHistoryRpt.Show
        Case "Station Sports Declaration", "Sports Clearance"          '10-9-12, 10-16-12
            'frmSportDeclareRpt.Show
        Case "Agreement Renewal Status"                             '11-8-12
            'frmRenewalRpt.Show
        Case "Advertiser Compliance"                                 '2-25-13
            'frmAdvComplyRpt.Show
        Case Else
            'frmReports.Show
    End Select
    If igReportModelessStatus = 0 Then
        Do While igReportModelessStatus = 0
            For ilLoop = 0 To 10 Step 1
                DoEvents
            Next ilLoop
            If ilCount >= 30 Then
                mPopulate
                ilCount = 0
            End If
            ilCount = ilCount + 1
            Sleep 1000
        Loop
        slDateTime = gNow()
        slNowDate = Format$(slDateTime, "m/d/yy")
        slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
        SQLQuery = "UPDATE rqt SET "
        If igReportReturn <> 2 Then
            SQLQuery = SQLQuery & "rqtStatus = 'C'" & ", "
        Else
            sgTmfStatus = "E"
            SQLQuery = SQLQuery & "rqtStatus = 'E'" & ", "
        End If
        SQLQuery = SQLQuery & "rqtDateCompleted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "rqtTimeCompleted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
        SQLQuery = SQLQuery & "WHERE rqtCode = " & llRqtCode
        'cnn.Execute slSQL_AlertClear, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand1:
            gSetMousePointer grdStatus, grdStatus, vbDefault
            gHandleError "ReportQueueLog.txt", "ReportQueue-mStartReportLoop"
            grdStatus.Redraw = True
            mStartReportForm = False
            Exit Function
        End If
    ElseIf igReportModelessStatus = -1 Then
        sgTmfStatus = "E"
        gLogMsg "Unable to find Report Name: " & Trim$(slReportName), "AffErrorLog.Txt", False
        slDateTime = gNow()
        slNowDate = Format$(slDateTime, "m/d/yy")
        slNowTime = Format$(slDateTime, "h:mm:ssAM/PM")
        SQLQuery = "UPDATE rqt SET "
        SQLQuery = SQLQuery & "rqtStatus = 'E'" & ", "
        SQLQuery = SQLQuery & "rqtDateCompleted = '" & Format$(slNowDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "rqtTimeCompleted = '" & Format$(slNowTime, sgSQLTimeForm) & "' "
        SQLQuery = SQLQuery & "WHERE rqtCode = " & llRqtCode
        'cnn.Execute slSQL_AlertClear, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand1:
            gSetMousePointer grdStatus, grdStatus, vbDefault
            gHandleError "ReportQueueLog.txt", "ReportQueue-mStartReportLoop"
            grdStatus.Redraw = True
            mStartReportForm = False
            Exit Function
        End If
    End If
    mStartReportForm = True
    Exit Function
mStartReportFormErr:
    mStartReportForm = False
    Resume Next
ErrHand1:
    gSetMousePointer grdStatus, grdStatus, vbDefault
    gHandleError "ReportQueueLog.txt", "ReportQueue-mStartReportLoop"
    grdStatus.Redraw = True
    Return
End Function

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
    Set gg_rst = gSQLSelectCall(SQLQuery)
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
        gSetRptGGFlag slName
    End If
    gg_rst.Close
End Sub
Public Sub gAllowedExportsImportsInMenu(blIsOn As Boolean, ilVendor As Vendors)
    '8156
End Sub
