VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MsComm32.ocx"
Begin VB.Form EngrServiceMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engineering Services"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "EngrServiceMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   10245
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
      Height          =   3270
      Left            =   0
      Picture         =   "EngrServiceMain.frx":08CA
      ScaleHeight     =   3210
      ScaleWidth      =   10155
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10215
      Begin VB.FileListBox lbcFile 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8745
         Pattern         =   "*.sch"
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2970
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.FileListBox lbcLogFile 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   8805
         MultiSelect     =   2  'Extended
         Pattern         =   "*.Log"
         TabIndex        =   27
         Top             =   570
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.DirListBox lbcLogPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8865
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   3240
      End
      Begin VB.DriveListBox cbcLogDrive 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8790
         TabIndex        =   25
         Top             =   -15
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.ListBox lbcCommercialSort 
         Height          =   255
         ItemData        =   "EngrServiceMain.frx":7D6A
         Left            =   4755
         List            =   "EngrServiceMain.frx":7D6C
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ListBox lbcSort 
         Height          =   255
         Left            =   7095
         Sorted          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Timer tmcRestartTask 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8430
         Top             =   135
      End
      Begin MSCommLib.MSComm spcItemID 
         Left            =   7515
         Top             =   2490
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.TextBox edcAutoPrior 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1395
         Width           =   885
      End
      Begin VB.TextBox edcSchdPrior 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8970
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   825
         Width           =   885
      End
      Begin VB.TextBox edcAutoFor 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1395
         Width           =   1500
      End
      Begin VB.TextBox edcSchdFor 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4035
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   825
         Width           =   1500
      End
      Begin VB.TextBox edcAutoPurge 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1395
         Width           =   1560
      End
      Begin VB.Timer tmcStart 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7995
         Top             =   180
      End
      Begin VB.CommandButton cmcMin 
         Caption         =   "Minimize"
         Height          =   330
         Left            =   3450
         TabIndex        =   2
         Top             =   2595
         Width           =   1380
      End
      Begin VB.TextBox edcMergeCheck 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1935
         Width           =   1560
      End
      Begin VB.TextBox edcAutoCreate 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1395
         Width           =   1560
      End
      Begin VB.TextBox edcSchdPurge 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   825
         Width           =   1560
      End
      Begin VB.TextBox edcSchdCreate 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   825
         Width           =   1560
      End
      Begin VB.CommandButton cmcStop 
         Caption         =   "Stop"
         Height          =   330
         Left            =   5235
         TabIndex        =   3
         Top             =   2595
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
      Begin VB.Label lacTestMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test System       Test System         Test System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   3345
         TabIndex        =   30
         Top             =   405
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label lacTestMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test System       Test System         Test System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   3345
         TabIndex        =   29
         Top             =   75
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Label lacAutoPrior 
         BackStyle       =   0  'Transparent
         Caption         =   "Prior to"
         Height          =   195
         Left            =   8325
         TabIndex        =   21
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lacSchdPrior 
         BackStyle       =   0  'Transparent
         Caption         =   "Prior to"
         Height          =   195
         Left            =   8325
         TabIndex        =   19
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lacAutoFor 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
         Height          =   195
         Left            =   3615
         TabIndex        =   17
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lacSchdFor 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lacAutoPurge 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge"
         Height          =   195
         Left            =   5925
         TabIndex        =   14
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label lacMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next Date, Time Task will be Run and for which Date(s)"
         Height          =   225
         Left            =   1095
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   7140
      End
      Begin VB.Label lacMergeCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Merge: Check at"
         Height          =   330
         Left            =   180
         TabIndex        =   10
         Top             =   1980
         Width           =   1500
      End
      Begin VB.Label lacAutoCreate 
         BackStyle       =   0  'Transparent
         Caption         =   "Automation: Create at"
         Height          =   330
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   1680
      End
      Begin VB.Label lacSchdPurge 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge"
         Height          =   195
         Left            =   5925
         TabIndex        =   6
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lacSchdCreate 
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule: Create at"
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   870
         Width           =   1500
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
   End
End
Attribute VB_Name = "EngrServiceMain"
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
Private imInCreateSchd As Integer
Private imInCreateAuto As Integer

Private hmSEE As Integer
Private hmSOE As Integer
Private hmCME As Integer
Private hmCTE As Integer
Private hmRLE As Integer
Private tmRLE As RLE

Private imCancelled As Integer
Private imClosed As Integer
Private lmSleepTime As Long
Private smSpotEventTypeName As String
Private imSpotETECode As Integer
Private tmSvMergeSOE As SOE
Private tmSvSchdSOE As SOE
Private tmSvAutoSOE As SOE
Private tmSvSchdSGE As SGE
Private tmSvAutoSGE As SGE
Private tmSvSchdPurgeSGE As SGE
Private tmSvAutoPurgeSGE As SGE

Private smSchdDates() As String
Private smAutoDates() As String

Private smT1Comment() As String
Private smT2Comment() As String

Private tmLoadUnchgdEvent() As LOADUNCHGDEVENT
Private tmSeeTimeSort() As SEETIMESORT

Private smFileNames As String
Private smRenameFile() As String

Private tmSHE As SHE
Private tmChgSHE() As SHE
Private tmMergeSHE() As SHE
Private tmCTE As CTE
Private tmDHE As DHE
Private tmARE As ARE

Private smSEEStamp As String
Private tmPrevSHE As SHE
Private tmPrevSEE() As SEE
Private tmNextSHE As SHE
Private tmNextSEE() As SEE
Private bmIncPurgeDate As Boolean

Private lm1970 As Long

Private smExportStr As String
Private hmExport As Integer

Private hmMerge As Integer

Private hmMsg As Integer

Const EVENTTYPEINDEX = 0
Const EVENTIDINDEX = 1
Const BUSNAMEINDEX = 2
Const BUSCTRLINDEX = 3
Const TIMEINDEX = 4
Const STARTTYPEINDEX = 5
Const FIXEDINDEX = 6
Const ENDTYPEINDEX = 7
Const DURATIONINDEX = 8
Const MATERIALINDEX = 9
Const AUDIONAMEINDEX = 10
Const AUDIOITEMIDINDEX = 11
Const AUDIOISCIINDEX = 12
Const AUDIOCTRLINDEX = 13
Const BACKUPNAMEINDEX = 18
Const BACKUPCTRLINDEX = 19
Const PROTNAMEINDEX = 14
Const PROTITEMIDINDEX = 15
Const PROTISCIINDEX = 16
Const PROTCTRLINDEX = 17
Const RELAY1INDEX = 20
Const RELAY2INDEX = 21
Const FOLLOWINDEX = 22
Const SILENCETIMEINDEX = 23
Const SILENCE1INDEX = 24
Const SILENCE2INDEX = 25
Const SILENCE3INDEX = 26
Const SILENCE4INDEX = 27
Const NETCUE1INDEX = 28
Const NETCUE2INDEX = 29
Const TITLE1INDEX = 30
Const TITLE2INDEX = 31
Const ABCFORMATINDEX = 32
Const ABCPGMCODEINDEX = 33
Const ABCXDSMODEINDEX = 34
Const ABCRECORDITEMINDEX = 35



Private Sub cmcMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmcStop_Click()
    imCancelled = True
'    tmcTask.Enabled = False
'    Unload EngrServiceMain
End Sub

Private Sub Form_Load()
    Dim ilPos As Integer
    
    sgCommand = Command$
    If App.PrevInstance Then
        MsgBox "Only one copy of EngrService can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of EngrService path: " & App.Path & " from " & Trim$(gGetComputerName()), "EngrService.Log", False
        End
    End If
    sgClientFields = "A"
    ilPos = InStr(1, sgCommand, "/Demo", 1)
    If ilPos > 0 Then
        sgClientFields = ""
    End If
    ilPos = InStr(1, sgCommand, "/WWO", 1)
    If ilPos > 0 Then
        sgClientFields = "W"
    End If
    sgStartIn = CurDir$
    igOperationMode = 1
    tgUIE.iCode = 1
    igBkgdProg = 10
    imInCreateSchd = False
    imInCreateAuto = False
    tmcStart.Enabled = True
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRet As Integer
    
    If imClosed = True Then
        Exit Sub
    End If
    tmcRestartTask.Enabled = False
    tmcStart.Enabled = False
    ilRet = MsgBox("Stop this Service", vbQuestion + vbYesNo, "Stop Service")
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
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub mStartUp()
    Dim ilRet As Integer
    Dim slTime As String
    Dim slBuffer As String
    Dim slTimeOut As String
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slDatabase As String
    Dim slLocations As String
    
    tmcStart.Enabled = False
    
    
    
    If InStr(1, sgStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
        slLocations = "Locations"
        slDatabase = "Database"
        lacTestMsg(0).Visible = False
        lacTestMsg(1).Visible = False
    Else
        igTestSystem = True
        slLocations = "TestLocations"
        slDatabase = "TestDatabase"
        lacTestMsg(0).Visible = True
        lacTestMsg(1).Visible = True
    End If

    lmSleepTime = 1000 ' 5 seconds '300000    '5 Minutes
    imCancelled = False
    imClosed = False
    bmIncPurgeDate = False
    sgDatabaseName = ""
    sgExportDirectory = ""
    sgImportDirectory = ""
    sgSQLDateForm = "yyyy-mm-dd"
    sgSQLTimeForm = "hh:mm:ss"
    igSQLSpec = 1               'Pervasive 2000
    sgShowDateForm = "m/d/yyyy"
    sgShowTimeWOSecForm = "hh:mm"
    sgShowTimeWSecForm = "hh:mm:ss"
    igWaitCount = 10
    igTimeOut = -1
    lgLastServiceDate = -1
    lgLastServiceTime = -1
    lm1970 = gDateValue("1/1/1970")
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Engineer.Ini"
    ilRet = 0
    On Error GoTo mReadFileErr
    slTime = FileDateTime(sgIniPathFileName)
    If ilRet <> 0 Then
        MsgBox "Engineer.Ini missing from " & sgStartupDirectory, vbCritical
        Unload EngrServiceMain
        Exit Sub
    End If
    igRunningFrom = 0
    ilPos = InStr(1, sgCommand, "/Client", 1)
    If ilPos > 0 Then
        igRunningFrom = 1
    End If
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

    If Not gLoadOption(slDatabase, "Name", sgDatabaseName) Then
        MsgBox "Engineer.Ini [" & slDatabase & "] 'Name' key is missing.", vbCritical
        Unload EngrServiceMain
        Exit Sub
    End If
'    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
'        MsgBox "Engineer.Ini [Locations] 'Export' key is missing.", vbCritical
'        Unload EngrServiceMain
'        Exit Sub
'    End If
'    'Import is optional
'    If gLoadOption("Locations", "Import", sgImportDirectory) Then
'        sgImportDirectory = gSetPathEndSlash(sgImportDirectory)
'    Else
'        sgImportDirectory = ""
'    End If
    
    
    'Commented out below because I can't see why you would need a backslash
    'on the end of a DSN name
    'sgDatabaseName = gSetPathEndSlash(sgDatabaseName)
'    sgExportDirectory = gSetPathEndSlash(sgExportDirectory)
'    sgImportDirectory = gSetPathEndSlash(sgImportDirectory)
    
    Call gLoadOption("SQLSpec", "Date", sgSQLDateForm)
    Call gLoadOption("SQLSpec", "Time", sgSQLTimeForm)
    If gLoadOption("SQLSpec", "System", slBuffer) Then
        If slBuffer = "P7" Then
            igSQLSpec = 0
        End If
    End If
    If gLoadOption(slLocations, "TimeOut", slTimeOut) Then
        igTimeOut = Val(slTimeOut)
    End If
    Call gLoadOption("Showform", "Date", sgShowDateForm)
    Call gLoadOption("Showform", "TimeWSec", sgShowTimeWSecForm)
    Call gLoadOption("Showform", "TimeWOSec", sgShowTimeWOSecForm)
    
    If igRunningFrom = 0 Then
        If Not gLoadOption(slLocations, "ServerDBPath", sgDBPath) Then
            MsgBox "Engineer.Ini " & slLocations & "] 'ServerDBPath' key is missing.", vbCritical
            Unload EngrServiceMain
            Exit Sub
        End If
    Else
        If Not gLoadOption(slLocations, "DBPath", sgDBPath) Then
            MsgBox "Engineer.Ini [" & slLocations & "] 'DBPath' key is missing.", vbCritical
            Unload EngrServiceMain
            Exit Sub
        End If
    End If
    sgDBPath = gSetPathEndSlash(sgDBPath)
    sgMsgDirectory = sgDBPath & "Messages\"
    
    If gLoadOption(slLocations, "Exe", sgExeDirectory) Then
        sgExeDirectory = gSetPathEndSlash(sgExeDirectory)
    Else
        sgExeDirectory = ""
    End If


    sgDSN = sgDatabaseName
    ' The sgDatabaseName may contain an ending backslash. Although this does not seem to have
    ' any effect, it does not seem like a good practice to let it stay like this here incase a later version of the RDO doesn't like it.
    If Mid(sgDSN, Len(sgDSN), 1) = "\" Then
        ' Yes it did end with a slash. Remove it.
        sgDSN = Left(sgDSN, Len(sgDSN) - 1)
    End If
    
    'Set cnn = New ADODB.Connection
    'cnn.Open "DSN=" & sgDSN
    'Set rst = New ADODB.Recordset
   '
   ' If igTimeOut >= 0 Then
   '     cnn.CommandTimeout = igTimeOut
   ' End If
   '
   ' hgDB = CBtrvMngrInit(0, "", "", sgDBPath, 0, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    ilRet = gStartPervasive()
    
    sgCurrSOEStamp = ""
    sgCurrSGEStamp = ""
    sgCurrSPEStamp = ""
    sgCurrITEStamp = ""
    
    ilRet = gGetServiceStatus_MIE_MessageInfo("EngrServiceMain", tgMie)
    
    gGetSiteOption
    gGetAuto
    mPopulate
    sgMergedLastDateRun = ""
    sgMergeLastTimeRun = ""
    sgMergedNextDateRun = ""
    sgMergeNextTimeRun = ""
    'Force set time
    tmSvMergeSOE.sMergeStopFlag = "~"
    tmSvSchdSOE.sSchAutoGenSeq = "~"
    tmSvAutoSOE.sSchAutoGenSeq = "~"
    tmSvSchdSOE.sSchAutoGenSeqTst = "~"
    tmSvAutoSOE.sSchAutoGenSeqTst = "~"
    tmSvSchdPurgeSGE.sPurgeAfterGen = "~"
    tmSvAutoPurgeSGE.sPurgeAfterGen = "~"
    mSetTimes
    
    Exit Sub
    
mReadFileErr:
    ilRet = Err.Number
    Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tgSpotCurrSEE
    Erase smSchdDates
    Erase smAutoDates
    Erase lgLibDheUsed
    Erase tmLoadUnchgdEvent
    Erase tmSeeTimeSort
    Erase tmChgSHE
    Erase tmPrevSEE
    Erase tmNextSEE
    Erase smRenameFile
    mEraseArrays
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    btrStopAppl
    Set EngrServiceMain = Nothing   'Remove data segment
    End
End Sub

Private Sub spcItemID_OnComm()
    gErrorMsgPort spcItemID
End Sub

Private Sub tmcRestartTask_Timer()
    tmcRestartTask.Enabled = False
    mTaskLoop
End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mStartUp
    If mIsServiceRunning() Then
        MsgBox "Only one copy of EngrService can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        gLogMsg "Second copy of EngrService path: " & App.Path & " from " & Trim$(gGetComputerName()), "EngrService.Log", False
        End
    End If
    gLogMsg "EngrService path: " & App.Path & " from " & Trim$(gGetComputerName()), "EngrService.Log", False
'    tmcTask.Interval = CInt(lmSleepTime)
'    tmcTask.Enabled = True
    mTaskLoop
End Sub

Private Sub mSetTimes()
    Dim slNowDate As String
    Dim slDate As String
    Dim slNowTime As String
    Dim slTime As String
    Dim ilSGE As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilDayIndex As Integer
    Dim ilLeadDays As Integer
    Dim ilSetMerge As Integer
    Dim ilSetSchd As Integer
    Dim ilSetAuto As Integer
    Dim ilSetPurge As Integer
    Dim slDateTime As String
    Dim slSetDates As String
    
    gGetSiteOption
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "ddddd")
    slNowTime = Format$(slDateTime, "ttttt")
    'Merge
    ilSetMerge = mSetMerge()
    ilSetSchd = mSetSchd()
    ilSetAuto = mSetAuto()
    ilSetPurge = mSetPurge()
    If ilSetMerge Then
        mSetMergeTimes
        LSet tmSvMergeSOE = tgSOE
    End If
    'Schedule Time
    If ilSetSchd Then
        If (Not igTestSystem) And ((tgSOE.sSchAutoGenSeq = "I") Or (tgSOE.sSchAutoGenSeq = "S")) Then
            For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
                If (tgCurrSGE(ilSGE).sType = "S") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) Then
                        sgSchdNextDateRun = slNowDate
                        sgSchdNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    Else
                        sgSchdNextDateRun = DateAdd("d", 1, slNowDate)
                        sgSchdNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    End If
                    edcSchdCreate.Text = sgSchdNextDateRun & " " & sgSchdNextTimeRun
                    LSet tmSvSchdSGE = tgCurrSGE(ilSGE)
                    Exit For
                End If
            Next ilSGE
         ElseIf (igTestSystem) And ((tgSOE.sSchAutoGenSeqTst = "I") Or (tgSOE.sSchAutoGenSeqTst = "S")) Then
            For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
                If (tgCurrSGE(ilSGE).sType = "S") And (tgCurrSGE(ilSGE).sSubType = "T") Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) Then
                        sgSchdNextDateRun = slNowDate
                        sgSchdNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    Else
                        sgSchdNextDateRun = DateAdd("d", 1, slNowDate)
                        sgSchdNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    End If
                    edcSchdCreate.Text = sgSchdNextDateRun & " " & sgSchdNextTimeRun
                    LSet tmSvSchdSGE = tgCurrSGE(ilSGE)
                    Exit For
                End If
            Next ilSGE
       Else
            sgSchdNextDateRun = "After Automation"
            sgSchdNextTimeRun = "After Automation"
            edcSchdCreate.Text = "After Automation"
        End If
        LSet tmSvSchdSOE = tgSOE
    End If
    If ilSetAuto Then
        If (Not igTestSystem) And ((tgSOE.sSchAutoGenSeq = "I") Or (tgSOE.sSchAutoGenSeq = "A")) Then
            For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
                If (tgCurrSGE(ilSGE).sType = "A") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) Then
                        sgAutoNextDateRun = slNowDate
                        sgAutoNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    Else
                        sgAutoNextDateRun = DateAdd("d", 1, slNowDate)
                        sgAutoNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    End If
                    edcAutoCreate.Text = sgAutoNextDateRun & " " & sgAutoNextTimeRun
                    LSet tmSvAutoSGE = tgCurrSGE(ilSGE)
                    Exit For
                End If
            Next ilSGE
        ElseIf (igTestSystem) And ((tgSOE.sSchAutoGenSeqTst = "I") Or (tgSOE.sSchAutoGenSeqTst = "A")) Then
            For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
                If (tgCurrSGE(ilSGE).sType = "A") And (tgCurrSGE(ilSGE).sSubType = "T") Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) Then
                        sgAutoNextDateRun = slNowDate
                        sgAutoNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    Else
                        sgAutoNextDateRun = DateAdd("d", 1, slNowDate)
                        sgAutoNextTimeRun = Format$(tgCurrSGE(ilSGE).sGenTime, "hh:mm:00")
                    End If
                    edcAutoCreate.Text = sgAutoNextDateRun & " " & sgAutoNextTimeRun
                    LSet tmSvAutoSGE = tgCurrSGE(ilSGE)
                    Exit For
                End If
            Next ilSGE
        Else
            sgAutoNextDateRun = "After Schedule"
            sgAutoNextTimeRun = "After Schedule"
            edcAutoCreate.Text = "After Schedule"
        End If
        LSet tmSvAutoSOE = tgSOE
    End If
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (tgCurrSGE(ilSGE).sType = "S") And (tgCurrSGE(ilSGE).sSubType = "P") Then
            slStr = ""
            slSetDates = ""
            For ilDay = 0 To 6 Step 1
                If InStr(1, sgSchdNextDateRun, "After", vbTextCompare) <= 0 Then
                    slDate = sgSchdNextDateRun
                Else
                    slDate = sgAutoNextDateRun
                End If
                Select Case ilDay
                    Case 0
                        ilLeadDays = tgCurrSGE(ilSGE).iGenMo
                        ilDayIndex = vbMonday
                    Case 1
                        ilLeadDays = tgCurrSGE(ilSGE).iGenTu
                        ilDayIndex = vbTuesday
                    Case 2
                        ilLeadDays = tgCurrSGE(ilSGE).iGenWe
                        ilDayIndex = vbWednesday
                    Case 3
                        ilLeadDays = tgCurrSGE(ilSGE).iGenTh
                        ilDayIndex = vbThursday
                    Case 4
                        ilLeadDays = tgCurrSGE(ilSGE).iGenFr
                        ilDayIndex = vbFriday
                    Case 5
                        ilLeadDays = tgCurrSGE(ilSGE).iGenSa
                        ilDayIndex = vbSaturday
                    Case 6
                        ilLeadDays = tgCurrSGE(ilSGE).iGenSu
                        ilDayIndex = vbSunday
                End Select
                slDate = DateAdd("d", ilLeadDays, slDate)
                If Weekday(slDate, vbSunday) = ilDayIndex Then
                    If slStr = "" Then
                        slStr = Format(slDate, "m/d")
                        slSetDates = Format(slDate, "m/d/yy")
                    Else
                        slStr = slStr & ", " & Format$(slDate, "m/d")
                        slSetDates = slSetDates & ", " & Format(slDate, "m/d/yy")
                    End If
                End If
            Next ilDay
            sgSchdForDates = slSetDates
            edcSchdFor.Text = slStr
            Exit For
        End If
    Next ilSGE

    If ilSetPurge Then
        sgPurgeDate = Format$(gNow(), "ddddd")
        sgPurgeDate = DateAdd("d", -tgSOE.iDaysRetainAsAir, sgPurgeDate)
        If bmIncPurgeDate Then
            sgPurgeDate = DateAdd("d", 1, sgPurgeDate)
        End If
        For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
            If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
                If tgCurrSGE(ilSGE).sPurgeAfterGen = "N" Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) Then
                        sgSchdPurgeNextDateRun = slNowDate
                        sgSchdPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    Else
                        sgSchdPurgeNextDateRun = DateAdd("d", 1, slNowDate)
                        sgSchdPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    End If
                    edcSchdPurge.Text = sgSchdPurgeNextDateRun & " " & sgSchdPurgeNextTimeRun
                    edcSchdPrior.Text = sgPurgeDate
                    'edcAutoPurge.Visible = False
                    'lacAutoPurge.Visible = False
                ElseIf tgCurrSGE(ilSGE).sPurgeAfterGen = "Y" Then
                    sgSchdPurgeNextDateRun = "After Schedule"
                    sgSchdPurgeNextTimeRun = "After Schedule"
                    edcSchdPurge.Text = "After Schedule"
                    edcSchdPrior.Text = sgPurgeDate
                    edcAutoPurge.Visible = False
                    lacAutoPurge.Visible = False
                    edcAutoPrior.Visible = False
                    lacAutoPrior.Visible = False
                Else
                    sgSchdPurgeNextDateRun = ""
                    sgSchdPurgeNextTimeRun = ""
                    edcSchdPurge.Text = ""
                    edcSchdPurge.Visible = False
                    lacSchdPurge.Visible = False
                    edcSchdPrior.Visible = False
                    lacSchdPrior.Visible = False
                End If
                LSet tmSvSchdPurgeSGE = tgCurrSGE(ilSGE)
                Exit For
            ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And (tgCurrSGE(ilSGE).sSubType = "T") Then
                If tgCurrSGE(ilSGE).sPurgeAfterGen = "N" Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) Then
                        sgSchdPurgeNextDateRun = slNowDate
                        sgSchdPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    Else
                        sgSchdPurgeNextDateRun = DateAdd("d", 1, slNowDate)
                        sgSchdPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    End If
                    edcSchdPurge.Text = sgSchdPurgeNextDateRun & " " & sgSchdPurgeNextTimeRun
                    edcSchdPrior.Text = sgPurgeDate
                    'edcAutoPurge.Visible = False
                    'lacAutoPurge.Visible = False
                ElseIf tgCurrSGE(ilSGE).sPurgeAfterGen = "Y" Then
                    sgSchdPurgeNextDateRun = "After Schedule"
                    sgSchdPurgeNextTimeRun = "After Schedule"
                    edcSchdPurge.Text = "After Schedule"
                    edcSchdPrior.Text = sgPurgeDate
                    edcAutoPurge.Visible = False
                    lacAutoPurge.Visible = False
                    edcAutoPrior.Visible = False
                    lacAutoPrior.Visible = False
                Else
                    sgSchdPurgeNextDateRun = ""
                    sgSchdPurgeNextTimeRun = ""
                    edcSchdPurge.Text = ""
                    edcSchdPurge.Visible = False
                    lacSchdPurge.Visible = False
                    edcSchdPrior.Visible = False
                    lacSchdPrior.Visible = False
                End If
                LSet tmSvSchdPurgeSGE = tgCurrSGE(ilSGE)
                Exit For
            End If
        Next ilSGE
        For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
            If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
                If tgCurrSGE(ilSGE).sPurgeAfterGen = "N" Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) Then
                        sgAutoPurgeNextDateRun = slNowDate
                        sgAutoPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    Else
                        sgAutoPurgeNextDateRun = DateAdd("d", 1, slNowDate)
                        sgAutoPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    End If
                    edcAutoPurge.Text = sgAutoPurgeNextDateRun & " " & sgAutoPurgeNextTimeRun
                    edcAutoPrior.Text = sgPurgeDate
                    'edcSchdPurge.Visible = False
                    'lacSchdPurge.Visible = False
                ElseIf tgCurrSGE(ilSGE).sPurgeAfterGen = "Y" Then
                    sgAutoPurgeNextDateRun = "After Automation"
                    sgAutoPurgeNextTimeRun = "After Automation"
                    edcAutoPurge.Text = "After Automation"
                    edcAutoPrior.Text = sgPurgeDate
                    edcSchdPurge.Visible = False
                    lacSchdPurge.Visible = False
                    edcSchdPrior.Visible = False
                    lacSchdPrior.Visible = False
                Else
                    sgAutoPurgeNextDateRun = ""
                    sgAutoPurgeNextTimeRun = ""
                    edcAutoPurge.Text = ""
                    edcAutoPurge.Visible = False
                    lacAutoPurge.Visible = False
                    edcAutoPrior.Visible = False
                    lacAutoPrior.Visible = False
                End If
                LSet tmSvAutoPurgeSGE = tgCurrSGE(ilSGE)
                Exit For
            ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And (tgCurrSGE(ilSGE).sSubType = "T") Then
                If tgCurrSGE(ilSGE).sPurgeAfterGen = "N" Then
                    If gTimeToLong(slNowTime, False) < gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) Then
                        sgAutoPurgeNextDateRun = slNowDate
                        sgAutoPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    Else
                        sgAutoPurgeNextDateRun = DateAdd("d", 1, slNowDate)
                        sgAutoPurgeNextTimeRun = Format$(tgCurrSGE(ilSGE).sPurgeTime, "hh:mm:00")
                    End If
                    edcAutoPurge.Text = sgAutoPurgeNextDateRun & " " & sgAutoPurgeNextTimeRun
                    edcAutoPrior.Text = sgPurgeDate
                    'edcSchdPurge.Visible = False
                    'lacSchdPurge.Visible = False
                ElseIf tgCurrSGE(ilSGE).sPurgeAfterGen = "Y" Then
                    sgAutoPurgeNextDateRun = "After Automation"
                    sgAutoPurgeNextTimeRun = "After Automation"
                    edcAutoPurge.Text = "After Automation"
                    edcAutoPrior.Text = sgPurgeDate
                    edcSchdPurge.Visible = False
                    lacSchdPurge.Visible = False
                    edcSchdPrior.Visible = False
                    lacSchdPrior.Visible = False
                Else
                    sgAutoPurgeNextDateRun = ""
                    sgAutoPurgeNextTimeRun = ""
                    edcAutoPurge.Text = ""
                    edcAutoPurge.Visible = False
                    lacAutoPurge.Visible = False
                    edcAutoPrior.Visible = False
                    lacAutoPrior.Visible = False
                End If
                LSet tmSvAutoPurgeSGE = tgCurrSGE(ilSGE)
                Exit For
            End If
        Next ilSGE
    End If
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (tgCurrSGE(ilSGE).sType = "A") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
            slStr = ""
            slSetDates = ""
            For ilDay = 0 To 6 Step 1
                If InStr(1, sgAutoNextDateRun, "After", vbTextCompare) <= 0 Then
                    slDate = sgAutoNextDateRun
                Else
                    slDate = sgSchdNextDateRun
                End If
                Select Case ilDay
                    Case 0
                        ilLeadDays = tgCurrSGE(ilSGE).iGenMo
                        ilDayIndex = vbMonday
                    Case 1
                        ilLeadDays = tgCurrSGE(ilSGE).iGenTu
                        ilDayIndex = vbTuesday
                    Case 2
                        ilLeadDays = tgCurrSGE(ilSGE).iGenWe
                        ilDayIndex = vbWednesday
                    Case 3
                        ilLeadDays = tgCurrSGE(ilSGE).iGenTh
                        ilDayIndex = vbThursday
                    Case 4
                        ilLeadDays = tgCurrSGE(ilSGE).iGenFr
                        ilDayIndex = vbFriday
                    Case 5
                        ilLeadDays = tgCurrSGE(ilSGE).iGenSa
                        ilDayIndex = vbSaturday
                    Case 6
                        ilLeadDays = tgCurrSGE(ilSGE).iGenSu
                        ilDayIndex = vbSunday
                End Select
                slDate = DateAdd("d", ilLeadDays, slDate)
                If Weekday(slDate, vbSunday) = ilDayIndex Then
                    If slStr = "" Then
                        slStr = Format(slDate, "m/d")
                        slSetDates = Format(slDate, "m/d/yy")
                    Else
                        slStr = slStr & ", " & Format$(slDate, "m/d")
                        slSetDates = slSetDates & ", " & Format(slDate, "m/d/yy")
                    End If
                End If
            Next ilDay
            sgAutoForDates = slSetDates
            edcAutoFor.Text = slStr
            Exit For
        End If
    Next ilSGE
    
End Sub

Private Sub mCreateSchd(slAirDate As String)
    Dim ilRet As Integer
    Dim llRow As Long
    Dim slComment As String
    Dim ilLoop As Integer
    Dim llNewAgedDHECode As Long
    

    If imInCreateSchd Then
        gLogMsg "mCreateSchd called a second time while in nCreateSchd.  The second call is ignored " & slAirDate, "EngrService.Log", False
        Exit Sub
    End If
    imInCreateSchd = True
    
    hmRLE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRLE, "", sgDBPath & "RLE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    tmRLE.lCode = 0
    tmRLE.sFileName = "SHE"
    tmRLE.lRecCode = DateValue(slAirDate)
    tmRLE.iUieCode = tgUIE.iCode
    tmRLE.sEnteredDate = Format$(gNow(), sgShowDateForm)
    tmRLE.sEnteredTime = Format$(gNow(), sgShowTimeWSecForm)
    If Not gPutInsert_RLE_Record_Locks(tmRLE, "mCreateSchd", hmRLE) Then
        btrDestroy hmRLE
        gLogMsg "Schedule being Created by another process for " & slAirDate, "EngrService.Log", False
        imInCreateSchd = False
        Exit Sub
    End If
    btrDestroy hmRLE
    
    mPopulate

    ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
    If ilRet Then
        imInCreateSchd = False
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateSchd")
        gLogMsg "Schedule Previously Created for " & slAirDate, "EngrService.Log", False
        Exit Sub
    End If
    
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmSOE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSOE, "", sgDBPath & "SOE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCME = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCME, "", sgDBPath & "CME.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    tmSHE.lCode = 0
    ilRet = gGetEventsFromLibraries(slAirDate)
    gCreateHeader slAirDate, tmSHE
    ilRet = gPutInsert_SHE_ScheduleHeader(0, tmSHE, "Schedule Definition-mSave: SHE")
    If Not ilRet Then
        imInCreateSchd = False
        btrDestroy hmSEE
        btrDestroy hmSOE
        btrDestroy hmCME
        btrDestroy hmCTE
        gLogMsg "Unable to Create Schedule Header for " & slAirDate, "EngrService.Log", False
        Exit Sub
    End If
    For llRow = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
        'If tgCurrSEE(llRow).l1CteCode > 0 Then
        '    ilRet = gGetRec_CTE_CommtsTitle(tgCurrSEE(llRow).l1CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
        '    slComment = Trim$(tmCTE.sComment)
        '    gSetCTE slComment, "T1", tmCTE
        '    ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Schedule Definition-mSave: Insert CTE", hmCTE)
        '    If ilRet Then
        '        tgCurrSEE(llRow).l1CteCode = tmCTE.lCode
        '    Else
        '        tgCurrSEE(llRow).l1CteCode = 0
        '    End If
        'Else
        '    tgCurrSEE(llRow).l1CteCode = 0
        'End If
        tgCurrSEE(llRow).lCode = 0
        tgCurrSEE(llRow).lSheCode = tmSHE.lCode
        tgCurrSEE(llRow).sAction = "N"
        tgCurrSEE(llRow).sSentStatus = "N"
        tgCurrSEE(llRow).sSentDate = Format$("12/31/2069", sgShowDateForm)
        ilRet = gPutInsert_SEE_ScheduleEvents(tgCurrSEE(llRow), "Schedule Definition-mSave: SEE", hmSEE, hmSOE)
        ilRet = gCreateCMEForSchd(tmSHE, tgCurrSEE(llRow), imSpotETECode, hmCME)
        gSetUsedFlags tgCurrSEE(llRow), hmCTE
    Next llRow
    For ilLoop = 0 To UBound(lgLibDheUsed) - 1 Step 1
        tmDHE.lCode = lgLibDheUsed(ilLoop)
        ilRet = gPutUpdate_DHE_DayHeaderInfo(2, tmDHE, "Schedule-mSave: Update DHE", llNewAgedDHECode)
    Next ilLoop
    ReDim lgLibDheUsed(0 To 0) As Long
    ilRet = mCheckEventConflicts(slAirDate, 0)
    If ilRet Then
        tmSHE.sConflictExist = "Y"
    Else
        tmSHE.sConflictExist = "N"
    End If
    ilRet = gPutUpdate_SHE_ScheduleHeader(5, tmSHE, "Schedule Definition-mSave: Update SHE", 0)
    btrDestroy hmSEE
    btrDestroy hmSOE
    btrDestroy hmCME
    btrDestroy hmCTE
    ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateSchd")
    gLogMsg "Successfully Created Schedule for " & slAirDate, "EngrService.Log", False
    imInCreateSchd = False
End Sub

Private Sub mCreateAuto(slAirDate As String)
    Dim ilRet As Integer
    Dim slComment As String
    Dim llRow As Long
    Dim ilEteCode As Integer
    Dim ilTestEteCode As Integer
    Dim llSEECode As Long
    Dim slEventCategory As String
    Dim slEventAutoCode As String
    Dim slTestEventCategory As String
    Dim slTestEventAutoCode As String
    Dim slExportFileName As String
    Dim slMsgFileName As String
    Dim ilLength As Integer
    Dim slDate As String
    Dim slTestDate As String
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilBDE As Integer
    Dim ilSend As Integer
    Dim llTest As Long
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim llAirDate As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slPrevDate As String
    Dim slNextDate As String
    Dim ilIndex As Integer
    Dim llLoop As Long
    Dim llAvailLength As Long
    Dim ilDelete As Integer
    Dim llOldSHECode As Long
    Dim ilLoadError As Integer

    If imInCreateAuto Then
        gLogMsg "mCreateAuto called a second time while in mCreateAuto.  The second call is ignored " & slAirDate, "EngrService.Log", False
        Exit Sub
    End If
    imInCreateAuto = True

    ilLoadError = False
    hmRLE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmRLE, "", sgDBPath & "RLE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    tmRLE.lCode = 0
    tmRLE.sFileName = "AUT"
    tmRLE.lRecCode = DateValue(slAirDate)
    tmRLE.iUieCode = tgUIE.iCode
    tmRLE.sEnteredDate = Format$(gNow(), sgShowDateForm)
    tmRLE.sEnteredTime = Format$(gNow(), sgShowTimeWSecForm)
    If Not gPutInsert_RLE_Record_Locks(tmRLE, "mCreateAuto", hmRLE) Then
        btrDestroy hmRLE
        gLogMsg "Auto Load being Created by another process for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    btrDestroy hmRLE

    mPopulate

    gGetAuto
    ilLength = gExportStrLength()
    If tgNoCharAFE.iDate = 8 Then
        slDate = Format$(slAirDate, "yyyymmdd")
    ElseIf tgNoCharAFE.iDate = 6 Then
        slDate = Format$(slAirDate, "yymmdd")
    End If
    ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
    If Not ilRet Then
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
        gLogMsg "Unable to find Schedule for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    If tmSHE.sLoadedAutoStatus = "L" Then
        tmSHE.sCreateLoad = "N"
        ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
        gLogMsg "Schedule previously Created for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    If Not gOpenAutoMsgFile(slAirDate, slMsgFileName, hmMsg) Then
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
        gLogMsg "Unable to Create Load Message file: " & slMsgFileName & " for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    If Not gOpenAutoExportFile(tmSHE, slAirDate, slExportFileName, hmExport) Then
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
        If igOperationMode = 1 Then
            Print #hmMsg, "Unable to Create Load file: " & slExportFileName & " see " & "EngrServiceError.Txt" & " for error message"
        End If
        Close #hmMsg
        gLogMsg "Unable to Create Load file: " & slExportFileName & " for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    ilRet = gGetRecs_SEE_ScheduleEvents(sgCurrSEEStamp, tmSHE.lCode, "EngrSchd-Get Events", tgCurrSEE())
    If Not ilRet Then
        ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
        If igOperationMode = 1 Then
            Print #hmMsg, "Unable to Access Schedule Event File, see " & "EngrServiceError.Txt" & " for error message"
        End If
        Close #hmMsg
        Close #hmExport
        gLogMsg "Unable to Access Schedule Event File for " & slAirDate, "EngrService.Log", False
        imInCreateAuto = False
        Exit Sub
    End If
    'Mark any overbook spot as don't send
    'For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
    llLoop = 0
    Do While llLoop <= UBound(tgCurrSEE) - 1
        slEventCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        'If Avail, then check if overbooked
        If slEventCategory = "A" Then
            llAvailLength = tgCurrSEE(llLoop).lDuration
            For llTest = 0 To UBound(tgCurrSEE) - 1 Step 1
                If (tgCurrSEE(llLoop).iBdeCode = tgCurrSEE(llTest).iBdeCode) And (tgCurrSEE(llLoop).lTime = tgCurrSEE(llTest).lTime) And (tgCurrSEE(llTest).iEteCode = imSpotETECode) Then
                    llAvailLength = llAvailLength - tgCurrSEE(llTest).lDuration
                End If
            Next llTest
            If llAvailLength < 0 Then
                ilDelete = False
                llAvailLength = tgCurrSEE(llLoop).lDuration
                llTest = 0  'llLoop + 1
                Do
                    If (llTest <> llLoop) And (tgCurrSEE(llLoop).iBdeCode = tgCurrSEE(llTest).iBdeCode) And (tgCurrSEE(llLoop).lTime = tgCurrSEE(llTest).lTime) And (tgCurrSEE(llTest).iEteCode = imSpotETECode) Then
                        If ilDelete Then
                            For llRow = llTest To UBound(tgCurrSEE) - 2 Step 1
                                LSet tgCurrSEE(llRow) = tgCurrSEE(llRow + 1)
                            Next llRow
                            ReDim Preserve tgCurrSEE(0 To UBound(tgCurrSEE) - 1) As SEE
                        Else
                            If llAvailLength - tgCurrSEE(llTest).lDuration < 0 Then
                                ilDelete = True
                                For llRow = llTest To UBound(tgCurrSEE) - 2 Step 1
                                    LSet tgCurrSEE(llRow) = tgCurrSEE(llRow + 1)
                                Next llRow
                                ReDim Preserve tgCurrSEE(0 To UBound(tgCurrSEE) - 1) As SEE
                            Else
                                llAvailLength = llAvailLength - tgCurrSEE(llTest).lDuration
                                llTest = llTest + 1
                            End If
                        End If
                    Else
                        llTest = llTest + 1
                    End If
                Loop While llTest <= UBound(tgCurrSEE) - 1
            End If
        End If
    'Next llLoop
        llLoop = llLoop + 1
    Loop
    llAirDate = DateValue(slAirDate)
    If tmSHE.sLoadedAutoStatus = "L" Then
        slPrevDate = DateAdd("d", -1, slAirDate)
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slPrevDate, "EngrSchedule-Get Previous Schedule by Date", tmPrevSHE)
        If ilRet Then
            ilRet = gGetRecs_SEE_ScheduleEvents(smSEEStamp, tmPrevSHE.lCode, "EngrSchd-Get Events", tmPrevSEE())
            If Not ilRet Then
                ReDim tmPrevSEE(0 To 0) As SEE
            End If
        Else
            ReDim tmPrevSEE(0 To 0) As SEE
        End If
        slNextDate = DateAdd("d", 1, slAirDate)
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slNextDate, "EngrSchedule-Get Previous Schedule by Date", tmNextSHE)
        If ilRet Then
            ilRet = gGetRecs_SEE_ScheduleEvents(smSEEStamp, tmNextSHE.lCode, "EngrSchd-Get Events", tmNextSEE())
            If Not ilRet Then
                ReDim tmNextSEE(0 To 0) As SEE
            End If
        Else
            ReDim tmNextSEE(0 To 0) As SEE
        End If
    Else
        ReDim tmPrevSEE(0 To 0) As SEE
        ReDim tmNextSEE(0 To 0) As SEE
    End If
    'Resort times by Event Actual time then Bus
    gAutoSortTime tgCurrSEE()
    gAutoSortTime tmPrevSEE()
    gAutoSortTime tmNextSEE()
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = DateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    ReDim tmLoadUnchgdEvent(0 To 0) As LOADUNCHGDEVENT
    If tmSHE.sLoadedAutoStatus = "L" Then
        'Determine first and Last Event for each bus
        For llRow = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
            'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            '    If tgCurrSEE(llRow).iEteCode = tgCurrETE(ilETE).iCode Then
            '        slEventCategory = tgCurrETE(ilETE).sCategory
            '        Exit For
            '    End If
            'Next ilETE
            'If (slEventCategory = "P") Or (slEventCategory = "S") Then
            If gAutoExportRow(tgCurrSEE(llRow).iEteCode, slEventCategory, slEventAutoCode) Then
                If tgCurrSEE(llRow).sSentStatus <> "S" Then
                    'Check If today and enough time
                    ilSend = True
                    If llAirDate = llNowDate Then
                        If llNowTime > tgCurrSEE(llRow).lTime Then
                            ilSend = False
                        End If
                    End If
                    If ilSend Then
                        ilIndex = -1
                        For ilLoop = 0 To UBound(tmLoadUnchgdEvent) - 1 Step 1
                            If tgCurrSEE(llRow).iBdeCode = tmLoadUnchgdEvent(ilLoop).iBdeCode Then
                                ilIndex = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                        If ilIndex = -1 Then
                            ReDim Preserve tmLoadUnchgdEvent(0 To UBound(tmLoadUnchgdEvent) + 1)
                            ilIndex = UBound(tmLoadUnchgdEvent) - 1
                            tmLoadUnchgdEvent(ilIndex).iBdeCode = tgCurrSEE(llRow).iBdeCode
                            tmLoadUnchgdEvent(ilIndex).lFirstSEECode = -1
                            tmLoadUnchgdEvent(ilIndex).lLastSEECode = -1
                            tmLoadUnchgdEvent(ilIndex).sSendStatus = "N"
                            tmLoadUnchgdEvent(ilIndex).iLastMsgGen = False
                            ilFound = False
                            If (tgCurrSEE(llRow).sSentStatus = "N") And (tgCurrSEE(llRow).sAction = "N") Then
                                tmLoadUnchgdEvent(ilIndex).lFirstSEECode = tgCurrSEE(llRow).lCode
                                ilFound = True
                            End If
                            For llTest = llRow - 1 To LBound(tgCurrSEE) Step -1
                                'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                                '    If tgCurrSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                                '        slTestEventCategory = tgCurrETE(ilETE).sCategory
                                '        Exit For
                                '    End If
                                'Next ilETE
                                'If (slTestEventCategory = "P") Or (slTestEventCategory = "S") Then
                                If gAutoExportRow(tgCurrSEE(llTest).iEteCode, slTestEventCategory, slTestEventAutoCode) Then
                                    If tgCurrSEE(llRow).iBdeCode = tgCurrSEE(llTest).iBdeCode Then
                                        If tgCurrSEE(llTest).sSentStatus = "S" Then
                                            tmLoadUnchgdEvent(ilIndex).lFirstSEECode = tgCurrSEE(llTest).lCode
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next llTest
                            If Not ilFound Then
                                'Check if exist in previous day
                                For llTest = UBound(tmPrevSEE) - 1 To LBound(tmPrevSEE) Step -1
                                    'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                                    '    If tmPrevSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                                    '        slTestEventCategory = tgCurrETE(ilETE).sCategory
                                    '        Exit For
                                    '    End If
                                    'Next ilETE
                                    'If (slTestEventCategory = "P") Or (slTestEventCategory = "S") Then
                                    If gAutoExportRow(tmPrevSEE(llTest).iEteCode, slTestEventCategory, slTestEventAutoCode) Then
                                        If tgCurrSEE(llRow).iBdeCode = tmPrevSEE(llTest).iBdeCode Then
                                            If tmPrevSEE(llTest).sSentStatus = "S" Then
                                                If llTest <> 0 Then
                                                    tmLoadUnchgdEvent(ilIndex).lFirstSEECode = -llTest - 2
                                                Else
                                                    tmLoadUnchgdEvent(ilIndex).lFirstSEECode = -2
                                                End If
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next llTest
                                If Not ilFound Then
                                    slStr = ""
                                    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                        If tgCurrSEE(llRow).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                            slStr = Trim$(tgCurrBDE(ilBDE).sName)
                                            Exit For
                                        End If
                                    Next ilBDE
                                    Print #hmMsg, "Unable to Find Unchanged Starting Event on Bus " & slStr
                                    'Start with event that can't find previous event for
                                    tmLoadUnchgdEvent(ilIndex).lFirstSEECode = tgCurrSEE(llRow).lCode
                                    ilLoadError = True
                                End If
                            End If
                        End If
                        ilFound = False
                        If (tgCurrSEE(llRow).sSentStatus = "N") And (tgCurrSEE(llRow).sAction = "N") Then
                            tmLoadUnchgdEvent(ilIndex).lLastSEECode = tgCurrSEE(llRow).lCode
                            ilFound = True
                        Else
                            If tmLoadUnchgdEvent(ilIndex).iLastMsgGen = False Then
                                tmLoadUnchgdEvent(ilIndex).lLastSEECode = -1
                            End If
                        End If
                        'Look for next item on same bus, if it requires to be sent, then send this item
                        For llTest = llRow + 1 To UBound(tgCurrSEE) - 1 Step 1
                            'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            '    If tgCurrSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                            '        slTestEventCategory = tgCurrETE(ilETE).sCategory
                            '        Exit For
                            '    End If
                            'Next ilETE
                            'If (slTestEventCategory = "P") Or (slTestEventCategory = "S") Then
                            If gAutoExportRow(tgCurrSEE(llTest).iEteCode, slTestEventCategory, slTestEventAutoCode) Then
                                If tgCurrSEE(llRow).iBdeCode = tgCurrSEE(llTest).iBdeCode Then
                                    If tgCurrSEE(llTest).sSentStatus = "S" Then
                                        tmLoadUnchgdEvent(ilIndex).lLastSEECode = tgCurrSEE(llTest).lCode
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next llTest
                        If (Not ilFound) And (tmLoadUnchgdEvent(ilIndex).iLastMsgGen = False) Then
                            'Check next day
                            For llTest = LBound(tmNextSEE) To UBound(tmNextSEE) - 1 Step 1
                                'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                                '    If tmNextSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                                '        slTestEventCategory = tgCurrETE(ilETE).sCategory
                                '        Exit For
                                '    End If
                                'Next ilETE
                                'If (slTestEventCategory = "P") Or (slTestEventCategory = "S") Then
                                If gAutoExportRow(tmNextSEE(llTest).iEteCode, slTestEventCategory, slTestEventAutoCode) Then
                                    If tgCurrSEE(llRow).iBdeCode = tmNextSEE(llTest).iBdeCode Then
                                        If tmNextSEE(llTest).sSentStatus = "S" Then
                                            If llTest <> 0 Then
                                                tmLoadUnchgdEvent(ilIndex).lLastSEECode = -llTest - 2
                                            Else
                                                tmLoadUnchgdEvent(ilIndex).lLastSEECode = -2
                                            End If
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next llTest
                            If Not ilFound Then
                                slStr = ""
                                For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                    If tgCurrSEE(llRow).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                        slStr = Trim$(tgCurrBDE(ilBDE).sName)
                                        Exit For
                                    End If
                                Next ilBDE
                                Print #hmMsg, "Unable to Find Unchanged Ending Event on Bus " & slStr
                                tmLoadUnchgdEvent(ilIndex).iLastMsgGen = True
                                ilLoadError = True
                                'End with last event that unable to find event after
                                'tmLoadUnchgdEvent(ilIndex).lLastSEECode = tgCurrSEE(llTest).lCode
                                For llTest = UBound(tgCurrSEE) To llRow Step -1
                                    If tgCurrSEE(llRow).iBdeCode = tgCurrSEE(llTest).iBdeCode Then
                                       tmLoadUnchgdEvent(ilIndex).lLastSEECode = tgCurrSEE(llTest).lCode
                                    End If
                                Next llTest
                            End If
                        End If
                    End If
                End If
            End If
        Next llRow
    End If
    For llRow = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
        ilEteCode = tgCurrSEE(llRow).iEteCode
        'slEventCategory = ""
        'For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        '    If tgCurrSEE(llRow).iEteCode = tgCurrETE(ilETE).iCode Then
        '        slEventCategory = tgCurrETE(ilETE).sCategory
        '        slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
        '        Exit For
        '    End If
        'Next ilETE
        'If (slEventCategory = "P") Or (slEventCategory = "S") Then
        If gAutoExportRow(ilEteCode, slEventCategory, slEventAutoCode) Then
            If (tmSHE.sLoadedAutoStatus = "L") Then
                'ilFound = -1
                ilSend = False
                For ilLoop = 0 To UBound(tmLoadUnchgdEvent) - 1 Step 1
                    If tgCurrSEE(llRow).iBdeCode = tmLoadUnchgdEvent(ilLoop).iBdeCode Then
                        If tmLoadUnchgdEvent(ilLoop).sSendStatus = "N" Then
                            If tmLoadUnchgdEvent(ilLoop).lFirstSEECode >= 0 Then
                                If tmLoadUnchgdEvent(ilLoop).lFirstSEECode = tgCurrSEE(llRow).lCode Then
                                    tmLoadUnchgdEvent(ilLoop).sSendStatus = "S"
                                    ilSend = True
                                Else
                                    ilSend = False
                                End If
                            ElseIf tmLoadUnchgdEvent(ilLoop).lFirstSEECode < -1 Then
                                llTest = -tmLoadUnchgdEvent(ilLoop).lFirstSEECode - 2
                                ilTestEteCode = tmPrevSEE(llTest).iEteCode
                                slTestEventCategory = ""
                                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                                    If tmPrevSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                                        slTestEventCategory = tgCurrETE(ilETE).sCategory
                                        slTestEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
                                        Exit For
                                    End If
                                Next ilETE
                                If tgNoCharAFE.iDate = 8 Then
                                    slTestDate = Format$(slPrevDate, "yyyymmdd")
                                ElseIf tgNoCharAFE.iDate = 6 Then
                                    slTestDate = Format$(slPrevDate, "yymmdd")
                                End If
                                gAutoSendSEE hmExport, slTestEventCategory, slTestEventAutoCode, slTestDate, ilTestEteCode, ilLength, tmPrevSEE(llTest)
                                tmLoadUnchgdEvent(ilLoop).sSendStatus = "S"
                                ilSend = True
                            End If
                        ElseIf tmLoadUnchgdEvent(ilLoop).sSendStatus = "S" Then
                            ilSend = True
                            If tmLoadUnchgdEvent(ilLoop).lLastSEECode = tgCurrSEE(llRow).lCode Then
                                tmLoadUnchgdEvent(ilLoop).sSendStatus = "F"
                            End If
                        Else
                            ilSend = False
                        End If
                        Exit For
                    End If
                Next ilLoop
                'If ilFound = -1 Then
                '    ilSend = False
                'End If
            Else
                ilSend = True
            End If
        Else
            ilSend = False
        End If
        If ilSend Then
            'Check If today and enough time
            If DateValue(slAirDate) = DateValue(slNowDate) Then
                If llNowTime > tgCurrSEE(llRow).lTime Then
                    ilSend = False
                End If
            End If
        End If
        If ilSend Then
            If tgCurrSEE(llRow).sAction <> "D" Then
                gAutoSendSEE hmExport, slEventCategory, slEventAutoCode, slDate, ilEteCode, ilLength, tgCurrSEE(llRow)
            End If
            'Update SEE
            llSEECode = tgCurrSEE(llRow).lCode
            If llSEECode > 0 Then
                ilRet = gPutUpdate_SEE_SentFlag(llSEECode, "EngrSchd- Update SEE Sent Flag")
            End If
        End If
    Next llRow
    If (tmSHE.sLoadedAutoStatus = "L") Then
        For ilLoop = 0 To UBound(tmLoadUnchgdEvent) - 1 Step 1
            If (tmLoadUnchgdEvent(ilLoop).sSendStatus <> "F") And (tmLoadUnchgdEvent(ilLoop).lLastSEECode < -1) Then
                llTest = -tmLoadUnchgdEvent(ilLoop).lLastSEECode - 2
                ilTestEteCode = tmNextSEE(llTest).iEteCode
                slTestEventCategory = ""
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If tmNextSEE(llTest).iEteCode = tgCurrETE(ilETE).iCode Then
                        slTestEventCategory = tgCurrETE(ilETE).sCategory
                        slTestEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
                        Exit For
                    End If
                Next ilETE
                If tgNoCharAFE.iDate = 8 Then
                    slTestDate = Format$(slNextDate, "yyyymmdd")
                ElseIf tgNoCharAFE.iDate = 6 Then
                    slTestDate = Format$(slNextDate, "yymmdd")
                End If
                gAutoSendSEE hmExport, slTestEventCategory, slTestEventAutoCode, slTestDate, ilTestEteCode, ilLength, tmNextSEE(llTest)
            End If
        Next ilLoop
    End If
    Close hmMsg
    Close hmExport
    gRenameExportFile
    'Update header file
    If ilLoadError Then
        tmSHE.sLoadStatus = "E"
    Else
        tmSHE.sLoadStatus = "N"
    End If
    ilRet = gPutUpdate_SHE_ScheduleHeader(7, tmSHE, "Schedule Definition-mCreateAuto: Update SHE", 0)
    ilRet = gPutUpdate_SHE_SentFlags(tmSHE.lCode, "EngrSchd- Update SHE Sent Flags")
    If (tmSHE.sLoadedAutoStatus <> "L") Then
        tmSHE.sLoadedAutoStatus = "L"
        tmSHE.iChgSeqNo = 0
    Else
        tmSHE.sLoadedAutoStatus = "L"
        tmSHE.iChgSeqNo = tmSHE.iChgSeqNo + 1
    End If
    tmSHE.sLoadedAutoDate = Format$(gNow(), sgShowDateForm)
    tmSHE.sCreateLoad = "N"
    ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
    ilRet = gPutDelete_RLE_Record_Locks(tmRLE.lCode, "mCreateAuto")
    gLogMsg "Successfully Created Automation File for " & slAirDate, "EngrService.Log", False
    mArchiveLoad
    imInCreateAuto = False
End Sub

Private Sub mPopDNE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_DNE_DayName("C", "L", sgCurrLibDNEStamp, "EngrSchd-mPopulate Library Names", tgCurrLibDNE())
End Sub

Private Sub mPopDSE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_DSE_DaySubName("C", sgCurrDSEStamp, "EngrSchd-mPopDSE Day Subname", tgCurrDSE())
End Sub


Private Sub mPopBDE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrSchd-mPopBDE Bus Definition", tgCurrBDE())
End Sub

Private Sub mPopCCE_Audio()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "A", sgCurrAudioCCEStamp, "EngrSchd-mPopCCE_Audio Control Character", tgCurrAudioCCE())
End Sub

Private Sub mPopCCE_Bus()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_CCE_ControlChar("C", "B", sgCurrBusCCEStamp, "EngrSchd-mPopCCE_Bus Control Character", tgCurrBusCCE())
End Sub

Private Sub mPopTTE_StartType()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "S", sgCurrStartTTEStamp, "EngrSchd-mPopTTE_StartType Start Type", tgCurrStartTTE())
End Sub

Private Sub mPopTTE_EndType()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_TTE_TimeType("C", "E", sgCurrEndTTEStamp, "EngrSchd-mPopTTE_EndType End Type", tgCurrEndTTE())
End Sub

Private Sub mPopASE()
    Dim ilRet As Integer

    mPopANE
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrSchd-mPopASE Audio Source", tgCurrASE())
End Sub

Private Sub mPopSCE()

    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_SCE_SilenceChar("C", sgCurrSCEStamp, "EngrSchd-mPopSCE Silence Character", tgCurrSCE())
End Sub

Private Sub mPopNNE()

    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_NNE_NetcueName("C", sgCurrNNEStamp, "EngrSchd-mPopNNE Netcue", tgCurrNNE())
End Sub

Private Sub mPopCTE()

    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", sgCurrCTEStamp, "EngrSchd-mPopCTE Title 2", tgCurrCTE())
End Sub

Private Sub mPopANE()

    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrSchd-mPopANE Audio Audio Names", tgCurrANE())
End Sub

Private Sub mPopARE()

    Dim ilRet As Integer

    ilRet = gGetRecs_ARE_AdvertiserRefer(sgCurrAREStamp, "EngrSchd-mPopARE Advertiser Names", tgCurrARE())
End Sub

Private Sub mPopETE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrSchd-mPopETE Event Properties", tgCurrEPE())
End Sub

Private Sub mPopMTE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_MTE_MaterialType("C", sgCurrMTEStamp, "EngrSchd-mPopMTE Material Type", tgCurrMTE())
End Sub

Private Sub mPopRNE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrSchd-mPopRNE Relay", tgCurrRNE())
End Sub

Private Sub mPopFNE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_FNE_FollowName("C", sgCurrFNEStamp, "EngrSchd-mPopFNE Follow", tgCurrFNE())
End Sub

Private Sub mPopulate()
    Dim ilRet As Integer
    Dim ilETE As Integer
    
    mPopANE
    mPopASE
    mPopBDE
    mPopCCE_Audio
    mPopCCE_Bus
    mPopCTE
    mPopDNE
    mPopDSE
    mPopETE
    mPopFNE
    mPopMTE
    mPopNNE
    mPopRNE
    mPopSCE
    mPopTTE_EndType
    mPopTTE_StartType
    mPopARE
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
    imSpotETECode = 0
    smSpotEventTypeName = "Spot"
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            imSpotETECode = tgCurrETE(ilETE).iCode
            smSpotEventTypeName = Trim$(tgCurrETE(ilETE).sName)
            Exit For
        End If
    Next ilETE
End Sub

Private Sub mMakeExportStr(ilStartCol As Integer, ilNoChar As Integer, llCol As Long, ilUCase As Integer, ilEteCode As Integer, slInStr As String)
    Dim slStr As String
    If (ilStartCol > 0) And (gExportCol(ilEteCode, llCol)) Then
        slStr = Trim$(slInStr)
        Do While Len(slStr) < ilNoChar
            slStr = slStr & " "
        Loop
        If ilUCase Then
            slStr = UCase$(slStr)
        End If
        Mid(smExportStr, ilStartCol, ilNoChar) = slStr
    End If
End Sub



Private Function mOpenAutoExportFile(slAirDate As String, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilPosE As Integer
    Dim slName As String
    Dim slPath As String
    Dim slDateTime As String
    Dim slChar As String
    Dim slSeqNo As String

    On Error GoTo mOpenAutoExportFileErr:
    'slNowDate = Format$(gNow(), sgShowDateForm)
    slName = ""
    slPath = ""
    For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
        If ((tgCurrAPE(ilLoop).sType = "CE") And (igRunningFrom = 1)) Or ((tgCurrAPE(ilLoop).sType = "SE") And (igRunningFrom = 0)) Then
            If ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                If (tmSHE.sLoadedAutoStatus = "L") Then
                    slName = Trim$(tgCurrAPE(ilLoop).sChgFileName) & "." & Trim$(tgCurrAPE(ilLoop).sChgFileExt)
                Else
                    slName = Trim$(tgCurrAPE(ilLoop).sNewFileName) & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
                End If
                ilPos = InStr(1, slName, "Date", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                        'slDate = Format$(slNowDate, tgCurrAPE(ilLoop).sDateFormat)
                        slDate = Format$(slAirDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                    Else
                        'slDate = Format$(slNowDate, "yymmdd")
                        slDate = Format$(slAirDate, "yymmdd")
                    End If
                    slName = Left$(slName, ilPos - 1) & slDate & Mid(slName, ilPos + 4)
                End If
                ilPos = InStr(1, slName, "Time", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilLoop).sTimeFormat) <> "" Then
                        slTime = Format$(slNowDate, Trim$(tgCurrAPE(ilLoop).sTimeFormat))
                    Else
                        slTime = Format$(slNowDate, "hhmmss")
                    End If
                    slName = Left$(slName, ilPos - 1) & slTime & Mid(slName, ilPos + 4)
                End If
                'Check for Sequence number
                If (tmSHE.sLoadedAutoStatus = "L") Then
                    ilPos = InStr(1, slName, "S", vbTextCompare)
                    If ilPos > 0 Then
                        ilPosE = ilPos + 1
                        Do While ilPosE <= Len(slName)
                            slChar = Mid$(slName, ilPosE, 1)
                            If StrComp(slChar, "S", vbTextCompare) <> 0 Then
                                Exit Do
                            End If
                            ilPosE = ilPosE + 1
                        Loop
                        slSeqNo = Trim$(Str$(tmSHE.iChgSeqNo + 1))
                        Do While Len(slSeqNo) < ilPosE - ilPos
                            slSeqNo = "0" & slSeqNo
                        Loop
                        Mid$(slName, ilPos, ilPosE - ilPos) = slSeqNo
                    End If
                End If
            End If
            'slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            If (Not igTestSystem) And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            ElseIf (igTestSystem) And (tgCurrAPE(ilLoop).sSubType = "T") Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            End If
            If slPath <> "" Then
                If right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            'Exit For
        End If
    Next ilLoop
    If slName = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Load File Name missing for Client from Automation Equipment Definition", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Load File Name missing for Client from Automation Equipment Definition", "EngrErrors.Txt", False
            MsgBox "Load File Name missing for Client from Automation Equipment Definition", vbCritical
        End If
        mOpenAutoExportFile = False
        Exit Function
    End If
    If slPath = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Load Path missing for Client from Automation Equipment Definition", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Load Path missing for Client from Automation Equipment Definition", "EngrErrors.Txt", False
            MsgBox "Load Path missing for Client from Automation Equipment Definition", vbCritical
        End If
        mOpenAutoExportFile = False
        Exit Function
    End If
    
    ilRet = 0
    slToFile = slPath & slName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    ilRet = 0
    sgLoadFileName = slToFile
    ilPos = InStr(1, slToFile, ".", vbBinaryCompare)
    If ilPos > 0 Then
        Mid(slToFile, ilPos, 1) = "_"
        slToFile = slToFile & ".txt"
    Else
        mOpenAutoExportFile = False
        Exit Function
    End If
    sgTmpLoadFileName = slToFile
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    ilRet = 0
    On Error GoTo mOpenAutoExportFileErr:
    hmExport = FreeFile
    Open slToFile For Output As hmExport
    If ilRet <> 0 Then
        Close hmExport
        hmExport = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error# " & Err.Number, "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error# " & Err.Number, "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & " error# " & Err.Number, vbCritical
        End If
        mOpenAutoExportFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hmExport, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hmExport, ""
    slMsgFileName = slToFile
    mOpenAutoExportFile = True
    Exit Function
mOpenAutoExportFileErr:
    ilRet = 1
    Resume Next
End Function

Private Function mCheckEventConflicts(slAirDate As String, ilAfterSchdOrMerge As Integer) As Integer
'   ilAfterSchdOrMerge: 0= After Schedule; 1=After Merge
    Dim llRow1 As Long
    Dim llRow2 As Long
    Dim ilHour1 As Integer
    Dim ilHour2 As Integer
    Dim slHours1 As String
    Dim slHours2 As String
    Dim ilDay1 As Integer
    Dim ilDay2 As Integer
    Dim slDays1 As String
    Dim slDays2 As String
    Dim slStr As String
    Dim ilBus1 As Integer
    Dim ilBus2 As Integer
    Dim llTime1 As Long
    Dim llTime2 As Long
    Dim llDur1 As Long
    Dim llDur2 As Long
    Dim llStartTime1 As Long
    Dim llEndTime1 As Long
    Dim llStartTime2 As Long
    Dim llEndTime2 As Long
    Dim ilPriAudio1 As Integer
    Dim ilProtAudio1 As Integer
    Dim ilBkupAudio1 As Integer
    Dim ilPriAudio2 As Integer
    Dim ilProtAudio2 As Integer
    Dim ilBkupAudio2 As Integer
    Dim slPriItemID1 As String
    Dim slPriItemID2 As String
    Dim slProtItemID1 As String
    Dim slProtItemID2 As String
    Dim slBkupItemID1 As String
    Dim slBkupItemID2 As String
    Dim ilASE As Integer
    Dim ilETE As Integer
    Dim slEventCategory As String
    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim llLoop1 As Long
    Dim llLoop2 As Long
    Dim llPostTime As Long
    Dim llPreTime As Long
    Dim ilATE As Integer
    Dim ilCheckBus As Integer
    Dim ilCheckAudio As Integer
    Dim llEventID1 As Long
    Dim llEventID2 As Long
    
    mCheckEventConflicts = False
    'Conflict test not required when schedule created as each time library/template is added a conflict test is done
    'Removed conflict checking for schedule
    'Remove Conflict checking after merge
    'These test can be removed because library and templates are checked for conflicts.
    'If avail will not be defined to use the same audio on different buses
    '10/9/09:  The conflict checking has been removed from Library and Templates, therefore, add it back into schedule creation
    If (ilAfterSchdOrMerge = 0) Or (ilAfterSchdOrMerge = 1) Then
    '    Exit Function
    End If
    ilRet = mOpenConflictMsgFile(slAirDate, ilAfterSchdOrMerge, slMsgFileName)
    If Not ilRet Then
        Exit Function
    End If
    Print #hmMsg, "** Conflict Test: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, "For: " & slAirDate
    
    'Find max Pre and Post adjustment time to help minimize compare time
    llPostTime = 0
    llPreTime = 0
    For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
        If tgCurrATE(ilATE).lPreBufferTime > llPreTime Then
            llPreTime = tgCurrATE(ilATE).lPreBufferTime
        End If
        If tgCurrATE(ilATE).lPostBufferTime > llPostTime Then
            llPostTime = tgCurrATE(ilATE).lPostBufferTime
        End If
    Next ilATE

    lbcSort.Clear
    For llRow1 = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
        llTime1 = tgCurrSEE(llRow1).lTime
        slStr = Trim$(Str$(llTime1))
        Do While Len(slStr) < 8
            slStr = "0" & slStr
        Loop
        lbcSort.AddItem slStr
        lbcSort.ItemData(lbcSort.NewIndex) = llRow1
    Next llRow1
    
'    For llRow1 = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
    For llLoop1 = 0 To lbcSort.ListCount - 1 Step 1
        llRow1 = lbcSort.ItemData(llLoop1)
        slEventCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrSEE(llRow1).iEteCode = tgCurrETE(ilETE).iCode Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If (slEventCategory = "P") Or ((slEventCategory = "A") And (ilAfterSchdOrMerge = 0)) Or ((slEventCategory = "S") And (ilAfterSchdOrMerge = 1)) Then
            If (tgCurrSEE(llRow1).sAction <> "D") And (tgCurrSEE(llRow1).sAction <> "R") Then
                llEventID1 = tgCurrSEE(llRow1).lEventID
                If (slEventCategory = "S") And (ilAfterSchdOrMerge = 1) Then
                    llTime1 = tgCurrSEE(llRow1).lSpotTime
                Else
                    llTime1 = tgCurrSEE(llRow1).lTime
                End If
                llDur1 = tgCurrSEE(llRow1).lDuration
                llStartTime1 = llTime1
                llEndTime1 = llStartTime1 + llDur1
                If llEndTime1 < llStartTime1 Then
                    llEndTime1 = llStartTime1
                End If
                If llEndTime1 > 864000 Then
                    llEndTime1 = 864000
                End If
                ilPriAudio1 = -1
                For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                    If tgCurrSEE(llRow1).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                        ilPriAudio1 = tgCurrASE(ilASE).iPriAneCode
                        Exit For
                    End If
                Next ilASE
                ilProtAudio1 = tgCurrSEE(llRow1).iProtAneCode
                ilBkupAudio1 = tgCurrSEE(llRow1).iBkupAneCode
                slPriItemID1 = Trim$(tgCurrSEE(llRow1).sAudioItemID)
                slProtItemID1 = Trim$(tgCurrSEE(llRow1).sProtItemID)
                slBkupItemID1 = Trim$(tgCurrSEE(llRow1).sAudioItemID)
                ilBus1 = tgCurrSEE(llRow1).iBdeCode
                If (ilPriAudio1 <> 0) And (ilPriAudio1 = ilProtAudio1) Then
                    mPrintEventMsg "Primary and Protection defined with same Audio Name", ilPriAudio1, llEventID1, ilBus1, llTime1
                    mCheckEventConflicts = True
                End If
                If (ilPriAudio1 <> 0) And (ilPriAudio1 = ilBkupAudio1) Then
                    mPrintEventMsg "Primary and Backup defined with same Audio Name", ilPriAudio1, llEventID1, ilBus1, llTime1
                    mCheckEventConflicts = True
                End If
                If (ilBkupAudio1 <> 0) And (ilBkupAudio1 = ilProtAudio1) Then
                    mPrintEventMsg "Backup and Protection defined with same Audio Name", ilBkupAudio1, llEventID1, ilBus1, llTime1
                    mCheckEventConflicts = True
                End If
    '            For llRow2 = llRow1 + 1 To UBound(tgCurrSEE) - 1 Step 1
                'For llLoop2 = 0 To lbcSort.ListCount - 1 Step 1
                For llLoop2 = llLoop1 + 1 To lbcSort.ListCount - 1 Step 1
                    llRow2 = lbcSort.ItemData(llLoop2)
                    slEventCategory = ""
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If tgCurrSEE(llRow2).iEteCode = tgCurrETE(ilETE).iCode Then
                            slEventCategory = tgCurrETE(ilETE).sCategory
                            Exit For
                        End If
                    Next ilETE
                    llEventID2 = tgCurrSEE(llRow2).lEventID
                    If (slEventCategory = "S") And (ilAfterSchdOrMerge = 1) Then
                        llTime2 = tgCurrSEE(llRow2).lSpotTime
                    Else
                        llTime2 = tgCurrSEE(llRow2).lTime
                    End If
                    llDur2 = tgCurrSEE(llRow2).lDuration
                    llStartTime2 = llTime2
                    llEndTime2 = llStartTime2 + llDur2
                    If llEndTime2 < llStartTime2 Then
                        llEndTime2 = llStartTime2
                    End If
                    If llEndTime2 > 864000 Then
                        llEndTime2 = 864000
                    End If
                    'Compare can stop once the start time of the llLoop2 item is beyond compare time
                    If llEndTime1 + llPostTime + llPreTime + 3000 < llStartTime2 Then
                        Exit For
                    End If
                    If (slEventCategory = "P") Or ((slEventCategory = "A") And (ilAfterSchdOrMerge = 0)) Or ((slEventCategory = "S") And (ilAfterSchdOrMerge = 1)) Then
                        If (tgCurrSEE(llRow2).sAction <> "D") And (tgCurrSEE(llRow2).sAction <> "R") Then
                            ilPriAudio2 = -1
                            For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                If tgCurrSEE(llRow2).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                    ilPriAudio2 = tgCurrASE(ilASE).iPriAneCode
                                    Exit For
                                End If
                            Next ilASE
                            ilProtAudio2 = tgCurrSEE(llRow2).iProtAneCode
                            ilBkupAudio2 = tgCurrSEE(llRow2).iBkupAneCode
                            slPriItemID2 = Trim$(tgCurrSEE(llRow2).sAudioItemID)
                            slProtItemID2 = Trim$(tgCurrSEE(llRow2).sProtItemID)
                            slBkupItemID2 = Trim$(tgCurrSEE(llRow2).sAudioItemID)
                            ilBus2 = tgCurrSEE(llRow2).iBdeCode
                            ilCheckBus = True
                            If (tgCurrSEE(llRow1).lDheCode <> tgCurrSEE(llRow2).lDheCode) Then
                                If ((tgCurrSEE(llRow1).sIgnoreConflicts = "B") Or (tgCurrSEE(llRow1).sIgnoreConflicts = "I")) Then
                                    ilCheckBus = False
                                End If
                                If ((tgCurrSEE(llRow2).sIgnoreConflicts = "B") Or (tgCurrSEE(llRow2).sIgnoreConflicts = "I")) Then
                                    ilCheckBus = False
                                End If
                            End If
                            If tgSOE.sMatchBNotT = "N" Then
                                ilCheckBus = False
                            End If
                            If (ilBus1 = ilBus2) And (ilCheckBus) Then
                                If (llEndTime2 > llStartTime1) And (llStartTime2 < llEndTime1) Or (llStartTime1 = llStartTime2) Then
                                    'Conflict
                                    mPrintBusConflictMsg "Bus Conflict", ilBus1, llEventID1, llEventID2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                            End If
                            ilCheckAudio = True
                            If (tgCurrSEE(llRow1).lDheCode <> tgCurrSEE(llRow2).lDheCode) Then
                                If ((tgCurrSEE(llRow1).sIgnoreConflicts = "A") Or (tgCurrSEE(llRow1).sIgnoreConflicts = "I")) Then
                                    ilCheckAudio = False
                                End If
                                If ((tgCurrSEE(llRow2).sIgnoreConflicts = "A") Or (tgCurrSEE(llRow2).sIgnoreConflicts = "I")) Then
                                    ilCheckAudio = False
                                End If
                            End If
                            If ilCheckAudio Then
                                If mAudioConflicts(ilPriAudio1, ilPriAudio2, slPriItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Primary and Primary Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilPriAudio1, ilProtAudio2, slPriItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Primary and Protection Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilPriAudio1, ilBkupAudio2, slPriItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Primary and Backup Audio Conflict", ilPriAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilProtAudio1, ilPriAudio2, slProtItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Protection and Primary Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilProtAudio1, ilProtAudio2, slProtItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Protection and Protection Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilProtAudio1, ilBkupAudio2, slProtItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Protection and Backup Audio Conflict", ilProtAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilPriAudio2, slBkupItemID1, slPriItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Backup and Primary Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilProtAudio2, slBkupItemID1, slProtItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Backup and Protection Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                                If mAudioConflicts(ilBkupAudio1, ilBkupAudio2, slBkupItemID1, slBkupItemID2, llStartTime1, llEndTime1, llStartTime2, llEndTime2, ilBus1, ilBus2) Then
                                    'Conflict
                                    mPrintAudioConflictMsg "Backup and Backup Audio Conflict", ilBkupAudio1, llEventID1, llEventID2, ilBus1, ilBus2, llTime1, llTime2
                                    mCheckEventConflicts = True
                                End If
                            End If
                        End If
                    End If
                Next llLoop2
    '            Next llRow2
            End If
        End If
'    Next llRow1
    Next llLoop1
    Close hmMsg

End Function

Private Function mAudioConflicts(ilAudio1 As Integer, ilAudio2 As Integer, slItemID1 As String, slItemID2 As String, llStartTime1 As Long, llEndTime1 As Long, llStartTime2 As Long, llEndTime2 As Long, ilBus1 As Integer, ilBus2 As Integer) As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim llAdjStartTime1 As Long
    Dim llAdjEndTime1 As Long
    Dim llAdjStartTime2 As Long
    Dim llAdjEndTime2 As Long
    
    mAudioConflicts = False
    If ilAudio1 <= 0 Then
        Exit Function
    End If
    If ilAudio1 = ilAudio2 Then
        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
            If tgCurrANE(ilANE).iCode = ilAudio1 Then
                If tgCurrANE(ilANE).sCheckConflicts <> "N" Then
                    If (llStartTime1 <> llStartTime2) Or (llEndTime1 <> llEndTime2) Then
                        If tgSOE.sMatchANotT <> "N" Then
                            For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                    llPreTime = tgCurrATE(ilATE).lPreBufferTime
                                    llPostTime = tgCurrATE(ilATE).lPostBufferTime
                                    llAdjStartTime1 = llStartTime1 - llPreTime
                                    If llAdjStartTime1 < 0 Then
                                        llAdjStartTime1 = 0
                                    End If
                                    llAdjEndTime1 = llEndTime1 + llPostTime
                                    If llAdjEndTime1 > 864000 Then
                                        llAdjEndTime1 = 864000
                                    End If
                                    llAdjStartTime2 = llStartTime2 - llPreTime
                                    If llAdjStartTime2 < 0 Then
                                        llAdjStartTime2 = 0
                                    End If
                                    llAdjEndTime2 = llEndTime2 + llPostTime
                                    If llAdjEndTime2 > 864000 Then
                                        llAdjEndTime2 = 864000
                                    End If
                                    If (llAdjEndTime2 > llAdjStartTime1) And (llAdjStartTime2 < llAdjEndTime1) Then
                                        mAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            Next ilATE
                        End If
                    Else
                        If ilBus1 <> ilBus2 Then
                            If tgSOE.sMatchATNotB <> "N" Then
                                mAudioConflicts = True
                                Exit Function
                            End If
                        Else
                            If tgSOE.sMatchATBNotI <> "N" Then
                                If (Trim$(slItemID1) <> "") And (Trim$(slItemID2) <> "") Then
                                    If StrComp(slItemID1, slItemID2, vbTextCompare) <> 0 Then
                                        mAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilANE
    End If
    
End Function

Private Function mOpenConflictMsgFile(slAirDate As String, ilAfterSchdOrMerge As Integer, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo mOpenConflictMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        If ilAfterSchdOrMerge = 0 Then
            slToFile = sgMsgDirectory & "ConflictAfterSchedule_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        Else
            slToFile = sgMsgDirectory & "ConflictAfterMerge_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        End If
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenConflictMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        mOpenConflictMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMsgFileName = slToFile
    mOpenConflictMsgFile = True
    Exit Function
mOpenConflictMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Sub mPrintAudioConflictMsg(slMsg As String, ilAudio1 As Integer, llEventID1 As Long, llEventID2 As Long, ilBus1 As Integer, ilBus2 As Integer, llTime1 As Long, llTime2 As Long)
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slBus1 As String
    Dim slBus2 As String
    Dim ilBDE As Integer
    Dim slAudio As String
    Dim ilANE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slTime2 = gLongToStrLengthInTenth(llTime2, True)
    slAudio = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tgCurrANE(ilANE).iCode = ilAudio1 Then
            slAudio = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    slBus2 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus2 = tgCurrBDE(ilBDE).iCode Then
            slBus2 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID1 > 0 Then
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & "(Bus " & slBus1 & ")" & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2 & "(Bus " & slBus2 & ")"
        Else
            Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & "(Bus " & slBus1 & ")" & " and at " & slTime2 & "(Bus " & slBus2 & ")"
        End If
    Else
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")" & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2 & "(Bus " & slBus2 & ")"
        Else
            Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")" & " and " & slTime2 & "(Bus " & slBus2 & ")"
        End If
    End If
End Sub
Private Sub mPrintBusConflictMsg(slMsg As String, ilBus1 As Integer, llEventID1 As Long, llEventID2 As Long, llTime1 As Long, llTime2 As Long)
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slBus1 As String
    Dim slBus2 As String
    Dim ilBDE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slTime2 = gLongToStrLengthInTenth(llTime2, True)
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID1 > 0 Then
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " for event ID " & Trim$(Str$(llEventID1)) & " at " & slTime1 & " and at " & slTime2
        End If
    Else
        If llEventID2 > 0 Then
            Print #hmMsg, slMsg & " on " & slBus1 & " for events at " & slTime1 & " and Event ID " & Trim$(Str$(llEventID2)) & " at " & slTime2
        Else
            Print #hmMsg, slMsg & " on " & slBus1 & " for events at " & slTime1 & " and " & slTime2
        End If
    End If
End Sub


Private Sub mPrintEventMsg(slMsg As String, ilAudio1 As Integer, llEventID As Long, ilBus1 As Integer, llTime1 As Long)
    Dim slTime1 As String
    Dim slBus1 As String
    Dim ilBDE As Integer
    Dim slAudio As String
    Dim ilANE As Integer
    
    slTime1 = gLongToStrLengthInTenth(llTime1, True)
    slAudio = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tgCurrANE(ilANE).iCode = ilAudio1 Then
            slAudio = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    slBus1 = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If ilBus1 = tgCurrBDE(ilBDE).iCode Then
            slBus1 = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    If llEventID > 0 Then
        Print #hmMsg, slMsg & " on " & slAudio & " for event ID " & Trim$(Str$(llEventID)) & " at " & slTime1 & "(Bus " & slBus1 & ")"
    Else
        Print #hmMsg, slMsg & " on " & slAudio & " for events at " & slTime1 & "(Bus " & slBus1 & ")"
    End If
End Sub



Private Function mOpenMergeFile(slAirDate As String, slMergeFilePri As String, slMergeFileBkup As String) As Integer
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slName As String
    Dim slPathPri As String
    Dim slPathBkup As String
    Dim slPathPriTest As String
    Dim slPathBkupTest As String
    Dim slDateTime As String

    On Error GoTo mOpenMergeFileErr:
    slName = ""
    slPathPri = ""
    slPathBkup = ""
    slPathPriTest = ""
    slPathBkupTest = ""
    
    slName = Trim$(tgSOE.sMergeFileFormat) & "." & Trim$(tgSOE.sMergeFileExt)
    ilPos = InStr(1, slName, "Date", vbTextCompare)
    If ilPos > 0 Then
        If Trim$(tgSOE.sMergeDateFormat) <> "" Then
            slDate = Format$(slAirDate, Trim$(tgSOE.sMergeDateFormat))
        Else
            slDate = Format$(slAirDate, "yymmdd")
        End If
        slName = Left$(slName, ilPos - 1) & slDate & Mid(slName, ilPos + 4)
    End If
    For ilLoop = 0 To UBound(tgCurrSPE) - 1 Step 1
        If ((tgCurrSPE(ilLoop).sType = "SP") And (igRunningFrom = 0)) Or ((tgCurrSPE(ilLoop).sType = "CP") And (igRunningFrom = 1)) Then
            If (Not igTestSystem) And ((tgCurrSPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrSPE(ilLoop).sSubType) = "")) Then
                slPathPri = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathPri <> "" Then
                    If right(slPathPri, 1) <> "\" Then
                        slPathPri = slPathPri & "\"
                    End If
                End If
                'Exit For
            End If
            If (igTestSystem) And (tgCurrSPE(ilLoop).sSubType = "T") Then
                slPathPri = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathPri <> "" Then
                    If right(slPathPri, 1) <> "\" Then
                        slPathPri = slPathPri & "\"
                    End If
                End If
                'Exit For
            ElseIf (Not igTestSystem) And (tgCurrSPE(ilLoop).sSubType = "T") Then
                slPathPriTest = Trim$(tgCurrSPE(ilLoop).sPath)
                If slPathPriTest <> "" Then
                    If right(slPathPriTest, 1) <> "\" Then
                        slPathPriTest = slPathPriTest & "\"
                    End If
                End If
            End If
        End If
    Next ilLoop
'    For ilLoop = 0 To UBound(tgCurrSPE) - 1 Step 1
'        If ((tgCurrSPE(ilLoop).sType = "SB") And (igRunningFrom = 0)) Or ((tgCurrSPE(ilLoop).sType = "CB") And (igRunningFrom = 1)) Then
'            slPathBkup = Trim$(tgCurrSPE(ilLoop).sPath)
'            If slPathBkup <> "" Then
'                If Right(slPathBkup, 1) <> "\" Then
'                    slPathBkup = slPathBkup & "\"
'                End If
'            End If
'            Exit For
'        End If
'    Next ilLoop
    If slName = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Merge File Name missing for Client from Site Option", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Merge File Name missing for Client from Site Option", "EngrErrors.Txt", False
            MsgBox "Merge File Name missing for Client from Site Option", vbCritical
        End If
        mOpenMergeFile = False
        Exit Function
    End If
    If slPathPri = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Merge Path missing for Client from Site Option", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Merge Path missing for Client from Site Option", "EngrErrors.Txt", False
            MsgBox "Merge File Name missing for Client from Site Option", vbCritical
        End If
        mOpenMergeFile = False
        Exit Function
    End If
    
    ilRet = 0
    slToFile = slPathPri & slName
    slDateTime = FileDateTime(slToFile)
    If ilRet <> 0 Then
        If slPathBkup <> "" Then
            ilRet = 0
            slToFile = slPathBkup & slName
            slDateTime = FileDateTime(slToFile)
            If ilRet <> 0 Then
                'MsgBox "Merge File missing from " & slPathPri & slName & " and from " & slToFile, vbOKOnly
                mOpenMergeFile = False
                Exit Function
            End If
        Else
            'MsgBox "Merge File missing from " & slPathPri & slName, vbOKOnly
            mOpenMergeFile = False
            Exit Function
        End If
    End If
    ilRet = 0
    On Error GoTo mOpenMergeFileErr:
    hmMerge = FreeFile
    Open slToFile For Input Access Read As hmMerge
    If ilRet <> 0 Then
        Close hmMerge
        hmMerge = -1
        If igOperationMode = 1 Then
            gLogMsg "Open Merge File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open Merge File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Merge File Name missing for Client from Site Option", vbCritical
        End If
        mOpenMergeFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMergeFilePri = slPathPri & slName
    If slPathBkup <> "" Then
        slMergeFileBkup = slPathBkup & slName
    Else
        slMergeFileBkup = ""
    End If
    If (Not igTestSystem) And (tgSOE.sMergeStopFlagTst = "N") And (slPathPriTest <> "") Then
        On Error GoTo mOpenMergeFileErr:
        ilRet = 0
        slToFile = slPathPriTest & slName
        slDateTime = FileDateTime(slToFile)
        If ilRet = 0 Then
            Kill slPathPriTest & slName
        End If
        FileCopy slPathPri & slName, slPathPriTest & slName
    End If
    On Error GoTo 0
    mOpenMergeFile = True
    Exit Function
mOpenMergeFileErr:
    ilRet = 1
    Resume Next
End Function

Private Function mOpenMergeMsgFile(slAirDate As String, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo mOpenMergeMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "MergeSpots_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenMergeMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        mOpenMergeMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hmMsg, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hmMsg, ""
    slMsgFileName = slToFile
    mOpenMergeMsgFile = True
    Exit Function
mOpenMergeMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Function mOpenLoadMsgFile(slAirDate As String, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo mOpenLoadMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "Load_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenLoadMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        mOpenLoadMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMsgFileName = slToFile
    mOpenLoadMsgFile = True
    Exit Function
mOpenLoadMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Function mMerge(slAirDate As String) As Integer
'    Dim ilRet As Integer
'    Dim ilEof As Integer
'    Dim slLine As String
'    Dim slDate As String
'    Dim llAirDate As Long
'    Dim slTime As String
'    Dim llTime As Long
'    Dim slTitle As String
'    Dim slLen As String
'    Dim slBus As String
'    Dim slCopy As String
'    Dim llLoop As Long
'    Dim ilETE As Integer
'    Dim ilBDE As Integer
'    Dim ilBus As Integer
'    Dim llRow As Long
'    Dim llUpper As Long
'    Dim ilFound As Integer
'    Dim llPrevAvailLoop As Long
'    Dim slDateTime As String
'    Dim slNowDate As String
'    Dim slNowTime As String
'    Dim llNowDate As Long
'    Dim llNowTime As Long
'    Dim ilRemove As Integer
'    Dim ilFindMatch As Integer
'    Dim llAvailLength As Long
'    Dim llCheck As Long
'
'    mMerge = True
'    llAirDate = DateValue(slAirDate)
'    slDateTime = gNow()
'    slNowDate = Format(slDateTime, "ddddd")
'    slNowTime = Format(slDateTime, "ttttt")
'    llNowDate = DateValue(slNowDate)
'    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
'    If llAirDate = llNowDate Then
'        Print #hmMsg, "Commercial Merge Spots Prior to " & gLongToTime(llNowTime) & " on " & slAirDate & " not checked"
'    End If
'    'Remove Spots
'    llLoop = LBound(tgCurrSEE)
'    ReDim tgSpotCurrSEE(0 To 0) As SEE
'    Do While llLoop < UBound(tgCurrSEE)
'        ilFound = False
'        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'            If tgCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                ilFound = True
'                If tgCurrETE(ilETE).sCategory = "S" Then
'                    ilRemove = True
'                    If llAirDate = llNowDate Then
'                        If llNowTime > tgCurrSEE(llLoop).lTime Then
'                            ilRemove = False
'                        End If
'                    End If
'                    If ilRemove Then
'                        LSet tgSpotCurrSEE(UBound(tgSpotCurrSEE)) = tgCurrSEE(llLoop)
'                        ReDim Preserve tgSpotCurrSEE(0 To UBound(tgSpotCurrSEE) + 1) As SEE
'                        For llRow = llLoop + 1 To UBound(tgCurrSEE) - 1 Step 1
'                            LSet tgCurrSEE(llRow - 1) = tgCurrSEE(llRow)
'                        Next llRow
'                        ReDim Preserve tgCurrSEE(0 To UBound(tgCurrSEE) - 1) As SEE
'                    Else
'                        llLoop = llLoop + 1
'                    End If
'                Else
'                    llLoop = llLoop + 1
'                End If
'            End If
'        Next ilETE
'        If Not ilFound Then
'            llLoop = llLoop + 1
'        End If
'    Loop
'    Do
'        'Get Lines
'        ilRet = 0
'        On Error GoTo mMergeErr:
'        Line Input #hmMerge, slLine
'        On Error GoTo 0
'        If ilRet <> 0 Then
'            Exit Do
'        End If
'        If Trim$(slLine) <> "" Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            End If
'        End If
'        DoEvents
'        If Trim$(slLine) <> "" Then
'            slDate = Mid$(slLine, 3, 2) & "/" & Mid$(slLine, 5, 2) & "/" & Mid$(slLine, 1, 2)
'            If DateValue(slDate) <> llAirDate Then
'                mMerge = False
'                Print #hmMsg, "Commercial Merge Spot Date " & slDate & " does not Match Schedule Date " & slAirDate
'                Exit Function
'            End If
'            slTime = Mid$(slLine, 11, 2) & ":" & Mid$(slLine, 13, 2) & ":" & Mid$(slLine, 15, 2)
'            llTime = 10 * gLengthToLong(slTime)
'            slBus = Trim$(Mid$(slLine, 18, 5))
'            slCopy = Mid$(slLine, 24, 5)
'            slTitle = Trim$(Mid$(slLine, 30, 15))
'            slLen = "00:" & Mid$(slLine, 46, 2) & ":" & Mid$(slLine, 48, 2)
'            ilFound = False
'            llPrevAvailLoop = -1
'            ilFindMatch = True
'            If llAirDate = llNowDate Then
'                If llNowTime > llTime Then
'                    ilFindMatch = False
'                End If
'            End If
'            If ilFindMatch Then
'                For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
'                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'                        If tgCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                            If tgCurrETE(ilETE).sCategory = "A" Then
'                                For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
'                                    If (tgCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode) Then
'                                        ilBus = StrComp(Trim$(tgCurrBDE(ilBDE).sName), slBus, vbTextCompare)
'                                        If (ilBus = 0) Then
'                                            If (tgCurrSEE(llLoop).lTime = llTime) Then  'Or ((tgCurrSEE(llLoop).lTime > llTime) And (llPrevAvailLoop <> -1)) Then
'                                                ilFound = True
'                                                'Create event
'                                                llUpper = UBound(tgCurrSEE)
'                                                mInitSEE tgCurrSEE(llUpper)
'                                                If (tgCurrSEE(llLoop).lTime = llTime) Then
'                                                    LSet tgCurrSEE(llUpper) = tgCurrSEE(llLoop)
'                                                    llPrevAvailLoop = llLoop
'                                                Else
'                                                    LSet tgCurrSEE(llUpper) = tgCurrSEE(llPrevAvailLoop)
'                                                End If
'                                                tgCurrSEE(llUpper).lCode = 0
'                                                tgCurrSEE(llUpper).iEteCode = imSpotETECode
'                                                tgCurrSEE(llUpper).lDuration = 10 * gLengthToLong(slLen)
'                                                If tgCurrSEE(llUpper).iAudioAseCode > 0 Then
'                                                    tgCurrSEE(llUpper).sAudioItemID = slCopy
'                                                End If
'                                                If tgCurrSEE(llUpper).iProtAneCode > 0 Then
'                                                    tgCurrSEE(llUpper).sProtItemID = slCopy
'                                                End If
'                                                tgCurrSEE(llUpper).lSpotTime = llTime
'                                                tmARE.lCode = 0
'                                                tmARE.sName = slTitle
'                                                tmARE.sUnusued = ""
'                                                'Check that avail is not overbooked
'                                                llAvailLength = tgCurrSEE(llLoop).lDuration
'                                                For llCheck = llLoop + 1 To llUpper - 1 Step 1
'                                                    If (tgCurrSEE(llLoop).iBdeCode = tgCurrSEE(llCheck).iBdeCode) And (tgCurrSEE(llCheck).iEteCode = imSpotETECode) And (tgCurrSEE(llCheck).lTime = llTime) Then
'                                                        llAvailLength = llAvailLength - tgCurrSEE(llCheck).lDuration
'                                                    End If
'                                                Next llCheck
'                                                llAvailLength = llAvailLength - tgCurrSEE(llUpper).lDuration
'                                                If llAvailLength >= 0 Then
'                                                    ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchd-Merge Insert Advertiser Name")
'                                                    If ilRet Then
'                                                        tgCurrSEE(llUpper).lAreCode = tmARE.lCode
'                                                        mSpotMatch tgCurrSEE(llUpper)
'                                                        ReDim Preserve tgCurrSEE(0 To llUpper + 1) As SEE
'                                                    Else
'                                                        mMerge = False
'                                                        Print #hmMsg, "Unable to Add Advertiser/Product " & slDate & " " & slTime & " " & slTitle
'                                                        mInitSEE tgCurrSEE(llUpper)
'                                                    End If
'                                                Else
'                                                    mMerge = False
'                                                    Print #hmMsg, "Commercial Merge Spot Overbooked Avail " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                                                End If
'                                                Exit For
'                                            ElseIf tgCurrSEE(llLoop).lTime < llTime Then
'                                                If llPrevAvailLoop <> -1 Then
'                                                    If tmCurrSEE(llLoop).lTime > tmCurrSEE(llPrevAvailLoop).lTime Then
'                                                        llPrevAvailLoop = llLoop
'                                                    End If
'                                                Else
'                                                    llPrevAvailLoop = llLoop
'                                                End If
'                                            End If
'                                        End If
'                                    End If
'                                Next ilBDE
'                                If ilFound Then
'                                    Exit For
'                                End If
'                            End If
'                        End If
'                    Next ilETE
'                    If ilFound Then
'                        Exit For
'                    End If
'                Next llLoop
'                If (Not ilFound) And (llPrevAvailLoop >= 0) Then
'                    ilFound = True
'                    llUpper = UBound(tgCurrSEE)
'                    mInitSEE tgCurrSEE(llUpper)
'                    LSet tgCurrSEE(llUpper) = tgCurrSEE(llPrevAvailLoop)
'                    tgCurrSEE(llUpper).lCode = 0
'                    tgCurrSEE(llUpper).iEteCode = imSpotETECode
'                    tgCurrSEE(llUpper).lDuration = 10 * gLengthToLong(slLen)
'                    If tgCurrSEE(llUpper).iAudioAseCode > 0 Then
'                        tgCurrSEE(llUpper).sAudioItemID = slCopy
'                    End If
'                    If tgCurrSEE(llUpper).iProtAneCode > 0 Then
'                        tgCurrSEE(llUpper).sProtItemID = slCopy
'                    End If
'                    tgCurrSEE(llUpper).lSpotTime = llTime
'                    tmARE.lCode = 0
'                    tmARE.sName = slTitle
'                    tmARE.sUnusued = ""
'                    'Check that avail is not overbooked
'                    llAvailLength = tgCurrSEE(llPrevAvailLoop).lDuration
'                    For llCheck = llPrevAvailLoop + 1 To llUpper - 1 Step 1
'                        If (tgCurrSEE(llPrevAvailLoop).iBdeCode = tgCurrSEE(llCheck).iBdeCode) And (tgCurrSEE(llCheck).iEteCode = imSpotETECode) And (tgCurrSEE(llCheck).lTime = llTime) Then
'                            llAvailLength = llAvailLength - tgCurrSEE(llCheck).lDuration
'                        End If
'                    Next llCheck
'                    llAvailLength = llAvailLength - tgCurrSEE(llUpper).lDuration
'                    If llAvailLength >= 0 Then
'                        ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchd-Merge Insert Advertiser Name")
'                        If ilRet Then
'                            tgCurrSEE(llUpper).lAreCode = tmARE.lCode
'                            mSpotMatch tgCurrSEE(llUpper)
'                            ReDim Preserve tgCurrSEE(0 To llUpper + 1) As SEE
'                        Else
'                            mMerge = False
'                            Print #hmMsg, "Unable to Add Advertiser/Product " & slDate & " " & slTime & " " & slTitle
'                            mInitSEE tgCurrSEE(llUpper)
'                        End If
'                    Else
'                        mMerge = False
'                        Print #hmMsg, "Commercial Merge Spot Overbooked Avail " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                    End If
'                End If
'                If Not ilFound Then
'                    mMerge = False
'                    Print #hmMsg, "Commercial Merge Spot Avail Not Found " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
'                End If
'            End If
'        End If
'    Loop Until ilEof
'    Exit Function
'mMergeErr:
'    ilRet = Err.Number
'    Resume Next
End Function



Private Sub mTestItemID()
'    Dim llLoop As Long
'    Dim ilETE As Integer
'    Dim ilITE As Integer
'    Dim tlPriITE As ITE
'    Dim tlSecITE As ITE
'    Dim slCart As String
'    Dim slQuery As String
'    Dim slPriQuery As String
'    Dim slResult As String
'    Dim slTitle As String
'    Dim ilASE As Integer
'    Dim slTestItemID As String
'    Dim ilATE As Integer
'    Dim ilANE As Integer
'    Dim ilRet As Integer
'    Dim ilTestPort As Integer
'
'    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
'        If tgCurrITE(ilITE).sType = "P" Then
'            LSet tlPriITE = tgCurrITE(ilITE)
'            Exit For
'        End If
'    Next ilITE
'    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
'        If tgCurrITE(ilITE).sType = "S" Then
'            LSet tlSecITE = tgCurrITE(ilITE)
'            Exit For
'        End If
'    Next ilITE
'    ilTestPort = True
'    For llLoop = 0 To UBound(tgCurrSEE) - 1 Step 1
'        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
'            If tgCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
'                If tgCurrETE(ilETE).sCategory = "S" Then
'                    slCart = Trim$(tgCurrSEE(llLoop).sAudioItemID)
'                    slTitle = ""
'                    If slCart <> "" Then
'                        slTestItemID = ""
'                        For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                            If tgCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
'                                For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
'                                    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
'                                        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
'                                            If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
'                                                slTestItemID = tgCurrATE(ilATE).sTestItemID
'                                                Exit For
'                                            End If
'                                        Next ilATE
'                                        If slTestItemID <> "" Then
'                                            Exit For
'                                        End If
'                                    End If
'                                Next ilANE
'                                If slTestItemID <> "" Then
'                                    Exit For
'                                End If
'                            End If
'                        Next ilASE
'                        If (slTestItemID = "Y") And (ilTestPort) Then
'                            ilRet = gGetRec_ARE_AdvertiserRefer(tgCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
'                            If ilRet Then
'                                slTitle = Trim$(tmARE.sName)
'                            End If
'                        End If
'                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) Then
'                            gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
'                            ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slPriQuery, slResult)
'                            If ilRet Then
'                                slResult = Mid$(slResult, Len(slPriQuery) + 1)
'                                If StrComp(Trim$(slTitle), slResult, vbTextCompare) = 0 Then
'                                    tgCurrSEE(llLoop).sAudioItemIDChk = "O"
'                                Else
'                                    tgCurrSEE(llLoop).sAudioItemIDChk = "F"
'                                End If
'                            Else
'                                If StrComp(slResult, "Failed", vbTextCompare) = 0 Then
'                                    ilTestPort = False
'                                End If
'                                tgCurrSEE(llLoop).sAudioItemIDChk = "N"
'                            End If
'                        Else
'                            If (slTestItemID = "Y") And (slTitle = "") Then
'                                tgCurrSEE(llLoop).sAudioItemIDChk = "N"
'                            End If
'                        End If
'                    End If
'                    slCart = Trim$(tgCurrSEE(llLoop).sProtItemID)
'                    If slCart <> "" Then
'                        slTestItemID = ""
'                        For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
'                            If tgCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
'                                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
'                                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
'                                        slTestItemID = tgCurrATE(ilATE).sTestItemID
'                                        Exit For
'                                    End If
'                                Next ilATE
'                                If slTestItemID <> "" Then
'                                    Exit For
'                                End If
'                            End If
'                        Next ilANE
'                        If (slTestItemID = "Y") And (slTitle = "") And (ilTestPort) Then
'                            ilRet = gGetRec_ARE_AdvertiserRefer(tgCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
'                            If ilRet Then
'                                slTitle = Trim$(tmARE.sName)
'                            End If
'                        End If
'                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) Then
'                            gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
'                            ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slPriQuery, slResult)
'                            If ilRet Then
'                                slResult = Mid$(slResult, Len(slPriQuery) + 1)
'                                If StrComp(Trim$(slTitle), slResult, vbTextCompare) = 0 Then
'                                    tgCurrSEE(llLoop).sProtItemIDChk = "O"
'                                Else
'                                    tgCurrSEE(llLoop).sProtItemIDChk = "F"
'                                End If
'                            Else
'                                If StrComp(slResult, "Failed", vbTextCompare) = 0 Then
'                                    ilTestPort = False
'                                End If
'                                tgCurrSEE(llLoop).sProtItemIDChk = "N"
'                            End If
'                        Else
'                            If (slTestItemID = "Y") And (slTitle = "") Then
'                                tgCurrSEE(llLoop).sProtItemIDChk = "N"
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        Next ilETE
'    Next llLoop
End Sub

Private Sub mCreateMerge()
    Dim ilMaxLeadDays As Integer
    Dim ilSGE As Integer
    Dim ilDayS As Integer
    Dim slNowDate As String
    Dim slAirDate As String
    Dim slDate As String
    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim slMergePriFile As String
    Dim slMergePriFileWOExt As String
    Dim slMergeBkupFile As String
    Dim ilSEECompare As Integer
    Dim llOldSEECode As Long
    Dim llRow As Long
    Dim ilPos As Integer
    Dim ilSHE As Integer
    Dim ilSEEChg As Integer
    Dim llOldSHECode As Long
    Dim ilSave As Integer
    Dim ilMergeError As Integer
    
'    'Determine date range
'    ilMaxLeadDays = -1
'    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
'        If tgCurrSGE(ilSGE).sType = "S" Then
'            If tgCurrSGE(ilSGE).iGenMo > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenMo
'            End If
'            If tgCurrSGE(ilSGE).iGenTu > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenTu
'            End If
'            If tgCurrSGE(ilSGE).iGenWe > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenWe
'            End If
'            If tgCurrSGE(ilSGE).iGenTh > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenTh
'            End If
'            If tgCurrSGE(ilSGE).iGenFr > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenFr
'            End If
'            If tgCurrSGE(ilSGE).iGenSa > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenSa
'            End If
'            If tgCurrSGE(ilSGE).iGenSu > ilMaxLeadDays Then
'                ilMaxLeadDays = tgCurrSGE(ilSGE).iGenSu
'            End If
'            Exit For
'        End If
'    Next ilSGE

    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmSOE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSOE, "", sgDBPath & "SOE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    mPopulate

    slNowDate = Format$(gNow(), "ddddd")
    ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slNowDate, "mCreateMerge, Get SHE", tmMergeSHE())
    'Get Site and check on Block
    'Loop on dates
    For ilSHE = 0 To UBound(tmMergeSHE) - 1 Step 1
        slAirDate = tmMergeSHE(ilSHE).sAirDate
        ilRet = mOpenMergeFile(slAirDate, slMergePriFile, slMergeBkupFile)
        If ilRet Then
            ilPos = InStr(1, slMergePriFile, ".", vbTextCompare)
            If ilPos > 0 Then
                slMergePriFileWOExt = Left$(slMergePriFile, ilPos - 1)
            Else
                slMergePriFileWOExt = slMergePriFile
            End If
            ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
            If ilRet Then
                ilRet = gGetRecs_SEE_ScheduleEvents(sgCurrSEEStamp, tmSHE.lCode, "EngrSchd-Get Events", tgCurrSEE())
                If ilRet Then
                    ilRet = mOpenMergeMsgFile(slAirDate, slMsgFileName)
                    'ilRet = mMerge(slAirDate)
                    gLogMsg "Create Merge for " & slAirDate, "EngrService.Log", False
                    ilRet = gMerge(0, slAirDate, hmMerge, hmMsg, tgCurrSEE(), smT1Comment(), smT2Comment(), lbcCommercialSort, ilMergeError)
                    Close #hmMerge
                    If ilRet Then
                        If ilMergeError Then
                            tmSHE.sSpotMergeStatus = "E"
                        Else
                            tmSHE.sSpotMergeStatus = "M"
                        End If
                        ilRet = gPutUpdate_SHE_ScheduleHeader(6, tmSHE, "Schedule Definition-mSave: Update SHE", 0)
                        ilRet = mCheckEventConflicts(slAirDate, 1)
                        If ilRet Then
                            tmSHE.sConflictExist = "Y"
                        Else
                            tmSHE.sConflictExist = "N"
                        End If
                        ''mTestItemID
                        '3/7/06- Remove automatic checking of ItemID
                        'gItemIDCheck spcItemID, tgCurrSEE()
                        'Update SHE and SEE
                        On Error Resume Next
                        'Kill slMergePriFile
                        Kill slMergePriFileWOExt & ".old"
                        On Error Resume Next
                        Name slMergePriFile As slMergePriFileWOExt & ".old"
                         
                        tmSHE.iVersion = tmSHE.iVersion + 1
                        ilRet = gPutUpdate_SHE_ScheduleHeader(3, tmSHE, "Schedule Definition-mSave: Update SHE", llOldSHECode)
                        ilSEEChg = False
                        For llRow = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
                            llOldSEECode = tgCurrSEE(llRow).lCode
                            If tgCurrSEE(llRow).lCode > 0 Then
                                If (tgCurrSEE(llRow).iEteCode <> imSpotETECode) Then
                                    ilSave = False
                                Else
                                    ilSEECompare = mCompareSEE(llOldSEECode)
                                    If ilSEECompare Then
                                        ilSave = False
                                    Else
                                        ilSave = True
                                    End If
                                End If
                            Else
                                ilSave = True
                            End If
                            If ilSave Then
                                tgCurrSEE(llRow).l1CteCode = 0
                                tgCurrSEE(llRow).lCode = 0
                                tgCurrSEE(llRow).lSheCode = tmSHE.lCode
                                If (tgCurrSEE(llRow).iEteCode = imSpotETECode) Then
                                    If Not ilSEECompare Then
                                        tgCurrSEE(llRow).sAction = "C"
                                        tgCurrSEE(llRow).sSentStatus = "N"
                                        tgCurrSEE(llRow).sSentDate = Format$("12/31/2069", sgShowDateForm)
                                        ilSEEChg = True
                                    Else
                                        'tgCurrSEE(llRow).sAction = "U"
                                    End If
                                Else
                                    'tgCurrSEE(llRow).sAction = "U"
                                End If
                                ilRet = gPutInsert_SEE_ScheduleEvents(tgCurrSEE(llRow), "Schedule Definition-mSave: SEE", hmSEE, hmSOE)
                                If llOldSEECode <> 0 Then
                                    ilRet = gPutReplace_SEE_SHECode(llOldSEECode, llOldSHECode, "Schedule Replace-mSave: SEE")
                                    ilRet = gUpdateAIE(1, tmSHE.iVersion, "SEE", llOldSEECode, tgCurrSEE(llRow).lCode, tmSHE.lOrigSheCode, "Schedule Definition- mSave: Insert SEE:AIE")
                                    gSetUsedFlags tgCurrSEE(llRow), hmCTE
                                End If
                            End If
                        Next llRow
                        If (tmSHE.sLoadedAutoStatus = "L") Then
                            If tmSHE.sCreateLoad <> "Y" Then
                                tmSHE.sCreateLoad = "Y"
                                ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateMerge- Update SHE", llOldSHECode)
                            End If
                        End If
                        '5/16/13: Handle case where spots removed
                        'Remove extra spots if merged
                        For llRow = 0 To UBound(tgSpotCurrSEE) - 1 Step 1
                            If tgSpotCurrSEE(llRow).lCode > 0 Then
                                If (tgSpotCurrSEE(llRow).iEteCode = imSpotETECode) Then
                                    ilRet = gPutDelete_CME_Conflict_Master("S", tmSHE.lCode, tgSpotCurrSEE(llRow).lCode, 0, "Schedule Definition- mSave: Delete SEE in CME", hmCME)
                                    ilRet = gPutDelete_SEE_Schedule_Events(tgSpotCurrSEE(llRow).lCode, "Schedule Definition-mSave: SEE")
                                End If
                            End If
                        Next llRow
                        ReDim tgSpotCurrSEE(0 To 0) As SEE
                    End If
                Else
                    Close #hmMerge
                End If
            Else
                Close #hmMerge
            End If
        End If
    Next ilSHE
    btrDestroy hmSEE
    btrDestroy hmSOE
    btrDestroy hmCTE
End Sub

Private Function mCompareSEE(llCode As Long) As Integer
    Dim ilSEENew As Integer
    Dim ilSEEOld As Integer
    Dim ilEBE As Integer
    Dim slStr As String
    Dim ilBDE As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    If llCode > 0 Then
        For ilSEENew = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1 Step 1
            If llCode = tgCurrSEE(ilSEENew).lCode Then
                For ilSEEOld = LBound(tgSpotCurrSEE) To UBound(tgSpotCurrSEE) - 1 Step 1
                    If llCode = tgSpotCurrSEE(ilSEEOld).lCode Then
                        
                        'Compare fields
                        'Buses
                        
                        If tgCurrSEE(ilSEENew).iBdeCode <> tgSpotCurrSEE(ilSEEOld).iBdeCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iBusCceCode <> tgSpotCurrSEE(ilSEEOld).iBusCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iEteCode <> tgSpotCurrSEE(ilSEEOld).iEteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).lTime <> tgSpotCurrSEE(ilSEEOld).lTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If (tgCurrSEE(ilSEENew).iEteCode = imSpotETECode) Then
                            If tgCurrSEE(ilSEENew).lSpotTime <> tgSpotCurrSEE(ilSEEOld).lSpotTime Then
                                mCompareSEE = False
                                Exit Function
                            End If
                        End If
                        If tgCurrSEE(ilSEENew).iStartTteCode <> tgSpotCurrSEE(ilSEEOld).iStartTteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).sFixedTime <> tgSpotCurrSEE(ilSEEOld).sFixedTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iEndTteCode <> tgSpotCurrSEE(ilSEEOld).iEndTteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).lDuration <> tgSpotCurrSEE(ilSEEOld).lDuration Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iMteCode <> tgSpotCurrSEE(ilSEEOld).iMteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iAudioAseCode <> tgSpotCurrSEE(ilSEEOld).iAudioAseCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sAudioItemID, tgSpotCurrSEE(ilSEEOld).sAudioItemID, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).sAudioItemIDChk <> tgSpotCurrSEE(ilSEEOld).sAudioItemIDChk Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sAudioISCI, tgSpotCurrSEE(ilSEEOld).sAudioISCI, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iAudioCceCode <> tgSpotCurrSEE(ilSEEOld).iAudioCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iBkupAneCode <> tgSpotCurrSEE(ilSEEOld).iBkupAneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iBkupCceCode <> tgSpotCurrSEE(ilSEEOld).iBkupCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iProtAneCode <> tgSpotCurrSEE(ilSEEOld).iProtAneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sProtItemID, tgSpotCurrSEE(ilSEEOld).sProtItemID, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).sProtItemIDChk <> tgSpotCurrSEE(ilSEEOld).sProtItemIDChk Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sProtISCI, tgSpotCurrSEE(ilSEEOld).sProtISCI, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iProtCceCode <> tgSpotCurrSEE(ilSEEOld).iProtCceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i1RneCode <> tgSpotCurrSEE(ilSEEOld).i1RneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i2RneCode <> tgSpotCurrSEE(ilSEEOld).i2RneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iFneCode <> tgSpotCurrSEE(ilSEEOld).iFneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).lSilenceTime <> tgSpotCurrSEE(ilSEEOld).lSilenceTime Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i1SceCode <> tgSpotCurrSEE(ilSEEOld).i1SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i2SceCode <> tgSpotCurrSEE(ilSEEOld).i2SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i3SceCode <> tgSpotCurrSEE(ilSEEOld).i3SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).i4SceCode <> tgSpotCurrSEE(ilSEEOld).i4SceCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iStartNneCode <> tgSpotCurrSEE(ilSEEOld).iStartNneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).iEndNneCode <> tgSpotCurrSEE(ilSEEOld).iEndNneCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If tgCurrSEE(ilSEENew).l2CteCode <> tgSpotCurrSEE(ilSEEOld).l2CteCode Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sABCFormat, tgSpotCurrSEE(ilSEEOld).sABCFormat, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sABCPgmCode, tgSpotCurrSEE(ilSEEOld).sABCPgmCode, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sABCXDSMode, tgSpotCurrSEE(ilSEEOld).sABCXDSMode, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        If StrComp(tgCurrSEE(ilSEENew).sABCRecordItem, tgSpotCurrSEE(ilSEEOld).sABCRecordItem, vbTextCompare) <> 0 Then
                            mCompareSEE = False
                            Exit Function
                        End If
                        '5/16/13: Handle case where spots removed
                        If tgSpotCurrSEE(ilSEEOld).lCode < 0 Then
                            ilRet = ilRet
                        End If
                        tgSpotCurrSEE(ilSEEOld).lCode = -tgSpotCurrSEE(ilSEEOld).lCode
                        
                        mCompareSEE = True
                        Exit Function
                    End If
                Next ilSEEOld
                '5/16/13: Change to false
                '         This code should not occur
                'mCompareSEE = True
                mCompareSEE = False
                Exit Function
            End If
        Next ilSEENew
    Else
        mCompareSEE = False
    End If
    
    
    
End Function

Private Function mOpenAutoMsgFile(slAirDate As String, slMsgFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo mOpenAutoMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "AutoExport_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo mOpenAutoMsgFileErr:
    hmMsg = FreeFile
    Open slToFile For Output As hmMsg
    If ilRet <> 0 Then
        Close hmMsg
        hmMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        mOpenAutoMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMsgFileName = slToFile
    mOpenAutoMsgFile = True
    Exit Function
mOpenAutoMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Private Sub mEraseArrays()
    Erase tgCurrACE
    Erase tgCurrADE
    Erase tgCurrAEE
    Erase tgCurrAFE
    Erase tgCurrAPE
    Erase tgCurrANE
    Erase tgBothANE
    Erase tgUsedANE
    Erase tgCurrASE
    Erase tgBothASE
    Erase tgCurrATE
    Erase tgBothATE
    Erase tgUsedATE
    Erase tgCurrBDE
    Erase tgBothBDE
    Erase tgUsedBDE
    Erase tgCurrBGE
    Erase tgCurrBSE
    Erase tgCurrCCE
    Erase tgCurrAudioCCE
    Erase tgUsedAudioCCE
    Erase tgCurrBusCCE
    Erase tgUsedBusCCE
    Erase tgCurrCTE
    Erase tgCurrDEE
    Erase tgCurrDHE
    Erase tgCurrLibDHE
    Erase tgBothLibDHE
    Erase tgCurrDNE
    Erase tgCurrLibDNE
    Erase tgCurrTempDHE
    Erase tgCurrTempDNE
    Erase tgCurrDSE
    Erase tgCurrEBE
    Erase tgCurrEPE
    Erase tgCurrETE
    Erase tgUsedETE
    Erase tgCurrFNE
    Erase tgUsedFNE
    Erase tgCurrITE
    Erase tgCurrMTE
    Erase tgUsedMTE
    Erase tgCurrNNE
    Erase tgUsedNNE
    Erase tgCurrRNE
    Erase tgUsedRNE
    Erase tgCurrSCE
    Erase tgUsedSCE
    Erase tgCurrSGE
    Erase tgCurrSOE
    Erase tgCurrSPE
    Erase tgCurrTTE
    Erase tgCurrStartTTE
    Erase tgUsedStartTTE
    Erase tgCurrEndTTE
    Erase tgUsedEndTTE
    Erase tgCurrTNE
    Erase tgCurrUIE
    Erase tgCurrUTE
    
    Erase lgLibDheUsed
    
    Erase tgJobTaskNames
    Erase tgListTaskNames
    Erase tgExtraTaskNames
    Erase tgAlertTaskNames
    Erase tgNoticeTaskNames
    
    Erase tgDDFFileNames
    
    Erase tgReportNames
    
    Erase tgFilterValues
    Erase tgFilterFields
    
    Erase tgSchdReplaceValues
    Erase tgReplaceFields
    Erase tgLibReplaceValues
    
    Erase tgItemIDChk
End Sub

Private Sub mTaskLoop()
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slMergeStatus As String
    Dim slSHEDate As String
    Dim ilLoadAsAirLog As Integer
    Dim ilTaskCount As Integer
    
    ilTaskCount = -1
    Do
        Sleep lmSleepTime
        If imCancelled Then
            Unload EngrServiceMain
            Exit Sub
        End If
        For ilLoop = 0 To 100 Step 1
            DoEvents
        Next ilLoop
        If (ilTaskCount = -1) Or (ilTaskCount = 60) Then
        
            gCheckIfDisconnected
            
            mUpdateServiceTime
            
            'Determine if any task needs to be performed
            slDateTime = gNow()
            slNowDate = Format$(slDateTime, "ddddd")
            slNowTime = Format$(slDateTime, "ttttt")
            ilLoadAsAirLog = False
            ilRet = gGetMergeStatus("EngrService: mTaskLoop- Get Merge Status", slMergeStatus)
            If (slMergeStatus <> "Y") And (ilRet = True) Then
                If DateValue(sgMergedNextDateRun) = DateValue(slNowDate) Then
                    If gTimeToLong(sgMergeNextTimeRun, False) <= gTimeToLong(slNowTime, False) Then
                        lacMsg.Caption = "Looking for Commercial Merge Files"
                        lacMsg.Visible = True
                        mCreateMerge
                        tmSvMergeSOE.sMergeStopFlag = "~"
                        sgMergeLastTimeRun = ""
                        lacMsg.Visible = False
                    End If
                End If
            End If
            For ilLoop = 0 To 100 Step 1
                DoEvents
            Next ilLoop
            'Check to see if any special Automation Loads need generation
            ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByLoadStatusAndDate("Y", slNowDate, "EngrService: mTaskLoop- Check for Changed Schedules", tmChgSHE())
            If ilRet Then
                For ilLoop = 0 To UBound(tmChgSHE) - 1 Step 1
                    slSHEDate = tmChgSHE(ilLoop).sAirDate
                    lacMsg.Caption = "Creating Load file for " & slSHEDate
                    lacMsg.Visible = True
                    gLogMsg "Creating special Load file for " & slSHEDate & " and Loop count = " & ilLoop, "EngrService.Log", False
                    mCreateAuto slSHEDate
                    imInCreateAuto = False
                    lacMsg.Visible = False
                Next ilLoop
                If UBound(tmChgSHE) > 0 Then
                    tmSvSchdSOE.sSchAutoGenSeq = "~"
                    tmSvSchdSOE.sSchAutoGenSeqTst = "~"
                End If
                ilLoadAsAirLog = True
            End If
            For ilLoop = 0 To 100 Step 1
                DoEvents
            Next ilLoop
            If (StrComp(sgSchdNextDateRun, "After Automation", vbTextCompare) <> 0) And (sgSchdNextDateRun <> "") Then
                If DateValue(sgSchdNextDateRun) = DateValue(slNowDate) Then
                    If gTimeToLong(sgSchdNextTimeRun, False) <= gTimeToLong(slNowTime, False) Then
                        'Start Schedule
                        gParseCDFields sgSchdForDates, False, smSchdDates()
                        If Trim$(smSchdDates(LBound(smSchdDates))) <> "" Then
                            For ilLoop = LBound(smSchdDates) To UBound(smSchdDates) Step 1
                                lacMsg.Caption = "Creating Schedule for " & smSchdDates(ilLoop)
                                lacMsg.Visible = True
                                gLogMsg "Creating Schedule for " & smSchdDates(ilLoop) & " on " & Trim$(gGetComputerName()), "EngrService.Log", False
                                mCreateSchd smSchdDates(ilLoop)
                                imInCreateSchd = False
                                lacMsg.Visible = False
                            Next ilLoop
                        End If
                        tmSvSchdSOE.sSchAutoGenSeq = "~"
                        tmSvSchdSOE.sSchAutoGenSeqTst = "~"
                        For ilLoop = 0 To 100 Step 1
                            DoEvents
                        Next ilLoop
                        If StrComp(sgAutoNextDateRun, "After Schedule", vbTextCompare) = 0 Then
                            'Start Automation
                            gParseCDFields sgAutoForDates, False, smAutoDates()
                            If Trim$(smAutoDates(LBound(smAutoDates))) <> "" Then
                                For ilLoop = LBound(smAutoDates) To UBound(smAutoDates) Step 1
                                    lacMsg.Caption = "Creating Load file for " & smAutoDates(ilLoop)
                                    lacMsg.Visible = True
                                    gLogMsg "Creating Load file for " & smAutoDates(ilLoop) & " After Schedule Created on " & Trim$(gGetComputerName()), "EngrService.Log", False
                                    mCreateAuto smAutoDates(ilLoop)
                                    imInCreateAuto = False
                                    lacMsg.Visible = False
                                Next ilLoop
                                ilLoadAsAirLog = True
                            End If
                            tmSvAutoSOE.sSchAutoGenSeq = "~"
                            tmSvAutoSOE.sSchAutoGenSeqTst = "~"
                        End If
                        For ilLoop = 0 To 100 Step 1
                            DoEvents
                        Next ilLoop
                        If (StrComp(sgSchdPurgeNextDateRun, "After Schedule", vbTextCompare) = 0) Or ((StrComp(sgAutoNextDateRun, "After Schedule", vbTextCompare) = 0) And (StrComp(sgSchdPurgeNextDateRun, "After Automation", vbTextCompare) = 0)) Then
                            'Start Purge
                            lacMsg.Caption = "Purging Old Information Prior to " & sgPurgeDate
                            lacMsg.Visible = True
                            If Trim$(sgPurgeDate) <> "" Then
                                ilRet = gSchdAndAsAiredDelete(sgPurgeDate, "Delete Schedule and As Air Prior to " & sgPurgeDate)
                                ilRet = gLibraryDelete(sgPurgeDate, "Delete Library Prior to " & sgPurgeDate)
                                ilRet = gTemplateSchdDelete(sgPurgeDate, "Delete Template Schedule Prior to " & sgPurgeDate)
                                ilRet = gCommentDelete("Delete Comment")
                                bmIncPurgeDate = True
                            End If
                            tmSvSchdPurgeSGE.sPurgeAfterGen = "~"
                            tmSvAutoPurgeSGE.sPurgeAfterGen = "~"
                            lacMsg.Visible = False
                        End If
                    End If
                End If
            End If
            For ilLoop = 0 To 100 Step 1
                DoEvents
            Next ilLoop
            If (StrComp(sgAutoNextDateRun, "After Schedule", vbTextCompare) <> 0) And (sgAutoNextDateRun <> "") Then
                If DateValue(sgAutoNextDateRun) = DateValue(slNowDate) Then
                    If gTimeToLong(sgAutoNextTimeRun, False) <= gTimeToLong(slNowTime, False) Then
                        'Start Automation
                        gParseCDFields sgAutoForDates, False, smAutoDates()
                        If Trim$(smAutoDates(LBound(smAutoDates))) <> "" Then
                            For ilLoop = LBound(smAutoDates) To UBound(smAutoDates) Step 1
                                lacMsg.Caption = "Creating Load file for " & smAutoDates(ilLoop)
                                lacMsg.Visible = True
                                gLogMsg "Creating Load file for " & smAutoDates(ilLoop) & " on " & Trim$(gGetComputerName()), "EngrService.Log", False
                                mCreateAuto smAutoDates(ilLoop)
                                imInCreateAuto = False
                                lacMsg.Visible = False
                            Next ilLoop
                            ilLoadAsAirLog = True
                        End If
                        tmSvAutoSOE.sSchAutoGenSeq = "~"
                        tmSvAutoSOE.sSchAutoGenSeqTst = "~"
                        For ilLoop = 0 To 100 Step 1
                            DoEvents
                        Next ilLoop
                        If StrComp(sgSchdNextDateRun, "After Automation", vbTextCompare) = 0 Then
                            'Start Start
                            gParseCDFields sgSchdForDates, False, smSchdDates()
                            If Trim$(smSchdDates(LBound(smSchdDates))) <> "" Then
                                For ilLoop = LBound(smSchdDates) To UBound(smSchdDates) Step 1
                                    lacMsg.Caption = "Creating Schedule for " & smSchdDates(ilLoop)
                                    lacMsg.Visible = True
                                    gLogMsg "Creating Schedule for " & smSchdDates(ilLoop) & " After Load file Created on " & Trim$(gGetComputerName()), "EngrService.Log", False
                                    mCreateSchd smSchdDates(ilLoop)
                                    imInCreateSchd = False
                                    lacMsg.Visible = False
                                Next ilLoop
                            End If
                            tmSvSchdSOE.sSchAutoGenSeq = "~"
                            tmSvSchdSOE.sSchAutoGenSeqTst = "~"
                        End If
                        For ilLoop = 0 To 100 Step 1
                            DoEvents
                        Next ilLoop
                        If (StrComp(sgSchdPurgeNextDateRun, "After Automation", vbTextCompare) = 0) Or ((StrComp(sgSchdNextDateRun, "After Automation", vbTextCompare) = 0) And (StrComp(sgSchdPurgeNextDateRun, "After Schedule", vbTextCompare) = 0)) Then
                            'Start Purge
                            lacMsg.Caption = "Purging Old Information Prior to " & sgPurgeDate
                            lacMsg.Visible = True
                            If Trim$(sgPurgeDate) <> "" Then
                                ilRet = gSchdAndAsAiredDelete(sgPurgeDate, "Delete Schedule and As Air Prior to " & sgPurgeDate)
                                ilRet = gLibraryDelete(sgPurgeDate, "Delete Library Prior to " & sgPurgeDate)
                                ilRet = gTemplateSchdDelete(sgPurgeDate, "Delete Template Schedule Prior to " & sgPurgeDate)
                                ilRet = gCommentDelete("Delete Comment")
                                bmIncPurgeDate = True
                            End If
                            tmSvSchdPurgeSGE.sPurgeAfterGen = "~"
                            tmSvAutoPurgeSGE.sPurgeAfterGen = "~"
                            lacMsg.Visible = False
                        End If
                    End If
                End If
            End If
            For ilLoop = 0 To 100 Step 1
                DoEvents
            Next ilLoop
            If (StrComp(sgSchdPurgeNextDateRun, "After Schedule", vbTextCompare) <> 0) And (StrComp(sgSchdPurgeNextDateRun, "After Automation", vbTextCompare) <> 0) And (sgSchdPurgeNextDateRun <> "") Then
                If DateValue(sgSchdPurgeNextDateRun) = DateValue(slNowDate) Then
                    If gTimeToLong(sgSchdPurgeNextTimeRun, False) <= gTimeToLong(slNowTime, False) Then
                        'Start Purge
                        lacMsg.Caption = "Purging Old Information Prior to " & sgPurgeDate
                        lacMsg.Visible = True
                        If sgPurgeDate <> "" Then
                            ilRet = gSchdAndAsAiredDelete(sgPurgeDate, "Delete Schedule and As Air Prior to " & sgPurgeDate)
                            ilRet = gLibraryDelete(sgPurgeDate, "Delete Library Prior to " & sgPurgeDate)
                            ilRet = gTemplateSchdDelete(sgPurgeDate, "Delete Template Schedule Prior to " & sgPurgeDate)
                            ilRet = gCommentDelete("Delete Comment")
                            bmIncPurgeDate = True
                        End If
                        tmSvSchdPurgeSGE.sPurgeAfterGen = "~"
                        tmSvAutoPurgeSGE.sPurgeAfterGen = "~"
                        lacMsg.Visible = False
                    End If
                End If
            End If
            For ilLoop = 0 To 100 Step 1
                DoEvents
            Next ilLoop
            If ilLoadAsAirLog Then
                mLoadAsAirLog
            End If
            lacMsg.Visible = False
            mSetTimes
            ilTaskCount = 1
        Else
            ilTaskCount = ilTaskCount + 1
        End If
   Loop
End Sub

Private Function mSetMerge() As Integer
    mSetMerge = False
    If tgSOE.sMergeStopFlag <> tmSvMergeSOE.sMergeStopFlag Then
        mSetMerge = True
        Exit Function
    End If
    If gTimeToLong(tgSOE.sMergeStartTime, False) <> gTimeToLong(tmSvMergeSOE.sMergeStartTime, False) Then
        mSetMerge = True
        Exit Function
    End If
    If gTimeToLong(tgSOE.sMergeEndTime, False) <> gTimeToLong(tmSvMergeSOE.sMergeEndTime, False) Then
        mSetMerge = True
        Exit Function
    End If
    If tgSOE.iMergeChkInterval <> tmSvMergeSOE.iMergeChkInterval Then
        mSetMerge = True
        Exit Function
    End If
End Function

Private Function mSetSchd()
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    Dim ilSGE As Integer
    
    mSetSchd = False
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "ddddd")
    slNowTime = Format$(slDateTime, "ttttt")
    If Not igTestSystem Then
        If tgSOE.sSchAutoGenSeq <> tmSvSchdSOE.sSchAutoGenSeq Then
            mSetSchd = True
            Exit Function
        End If
    ElseIf igTestSystem Then
        If tgSOE.sSchAutoGenSeqTst <> tmSvSchdSOE.sSchAutoGenSeqTst Then
            mSetSchd = True
            Exit Function
        End If
    End If
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
            If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) <> gTimeToLong(tmSvSchdSGE.sGenTime, False) Then
                If sgSchdNextDateRun = slNowDate Then
                    If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(tmSvSchdSGE.sGenTime, False) Then
                        mSetSchd = True
                        Exit Function
                    Else
                        If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(slNowTime, False) + lmSleepTime / 1000 Then
                            mSetSchd = True
                            Exit Function
                        End If
                    End If
                Else
                    mSetSchd = True
                End If
            End If
            Exit For
        ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And (tgCurrSGE(ilSGE).sSubType = "T") Then
            If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) <> gTimeToLong(tmSvSchdSGE.sGenTime, False) Then
                If sgSchdNextDateRun = slNowDate Then
                    If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(tmSvSchdSGE.sGenTime, False) Then
                        mSetSchd = True
                        Exit Function
                    Else
                        If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(slNowTime, False) + lmSleepTime / 1000 Then
                            mSetSchd = True
                            Exit Function
                        End If
                    End If
                Else
                    mSetSchd = True
                End If
            End If
            Exit For
        End If
    Next ilSGE
End Function

Private Function mSetAuto() As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    Dim ilSGE As Integer
    
    mSetAuto = False
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "ddddd")
    slNowTime = Format$(slDateTime, "ttttt")
    If Not igTestSystem Then
        If tgSOE.sSchAutoGenSeq <> tmSvAutoSOE.sSchAutoGenSeq Then
            mSetAuto = True
            Exit Function
        End If
    End If
    If igTestSystem Then
        If tgSOE.sSchAutoGenSeqTst <> tmSvAutoSOE.sSchAutoGenSeqTst Then
            mSetAuto = True
            Exit Function
        End If
    End If
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
            If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) <> gTimeToLong(tmSvAutoSGE.sGenTime, False) Then
                If sgAutoNextDateRun = slNowDate Then
                    If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(tmSvAutoSGE.sGenTime, False) Then
                        mSetAuto = True
                        Exit Function
                    Else
                        If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(slNowTime, False) + lmSleepTime / 1000 Then
                            mSetAuto = True
                            Exit Function
                        End If
                    End If
                Else
                    mSetAuto = True
                End If
            End If
            Exit For
        ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And (tgCurrSGE(ilSGE).sSubType = "T") Then
            If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) <> gTimeToLong(tmSvAutoSGE.sGenTime, False) Then
                If sgAutoNextDateRun = slNowDate Then
                    If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(tmSvAutoSGE.sGenTime, False) Then
                        mSetAuto = True
                        Exit Function
                    Else
                        If gTimeToLong(tgCurrSGE(ilSGE).sGenTime, False) > gTimeToLong(slNowTime, False) + lmSleepTime / 1000 Then
                            mSetAuto = True
                            Exit Function
                        End If
                    End If
                Else
                    mSetAuto = True
                End If
            End If
            Exit For
        End If
    Next ilSGE
End Function

Private Function mSetPurge() As Integer
    Dim ilSGE As Integer
    mSetPurge = False
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
            If tgCurrSGE(ilSGE).sPurgeAfterGen <> tmSvSchdPurgeSGE.sPurgeAfterGen Then
                mSetPurge = True
                Exit Function
            End If
            If gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) <> gTimeToLong(tmSvSchdPurgeSGE.sPurgeTime, False) Then
                mSetPurge = True
                Exit Function
            End If
        ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "S") And (tgCurrSGE(ilSGE).sSubType = "T") Then
            If tgCurrSGE(ilSGE).sPurgeAfterGen <> tmSvSchdPurgeSGE.sPurgeAfterGen Then
                mSetPurge = True
                Exit Function
            End If
            If gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) <> gTimeToLong(tmSvSchdPurgeSGE.sPurgeTime, False) Then
                mSetPurge = True
                Exit Function
            End If
        End If
    Next ilSGE
    For ilSGE = 0 To UBound(tgCurrSGE) - 1 Step 1
        If (Not igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And ((tgCurrSGE(ilSGE).sSubType = "P") Or (Trim$(tgCurrSGE(ilSGE).sSubType) = "")) Then
            If tgCurrSGE(ilSGE).sPurgeAfterGen <> tmSvAutoPurgeSGE.sPurgeAfterGen Then
                mSetPurge = True
                Exit Function
            End If
            If gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) <> gTimeToLong(tmSvAutoPurgeSGE.sPurgeTime, False) Then
                mSetPurge = True
                Exit Function
            End If
        ElseIf (igTestSystem) And (tgCurrSGE(ilSGE).sType = "A") And (tgCurrSGE(ilSGE).sSubType = "T") Then
            If tgCurrSGE(ilSGE).sPurgeAfterGen <> tmSvAutoPurgeSGE.sPurgeAfterGen Then
                mSetPurge = True
                Exit Function
            End If
            If gTimeToLong(tgCurrSGE(ilSGE).sPurgeTime, False) <> gTimeToLong(tmSvAutoPurgeSGE.sPurgeTime, False) Then
                mSetPurge = True
                Exit Function
            End If
        End If
    Next ilSGE
End Function

Private Sub mSetMergeTimes()
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDateTime As String
    Dim slTime As String
    
    slDateTime = gNow()
    slNowDate = Format$(slDateTime, "ddddd")
    slNowTime = Format$(slDateTime, "ttttt")
    If gTimeToLong(slNowTime, False) < gTimeToLong(tgSOE.sMergeStartTime, False) Then
        sgMergedNextDateRun = slNowDate
        sgMergeNextTimeRun = Format$(tgSOE.sMergeStartTime, "hh:mm:00")
    Else
        If gTimeToLong(slNowTime, True) < gTimeToLong(tgSOE.sMergeEndTime, True) Then
            If sgMergeLastTimeRun = "" Then
                sgMergedLastDateRun = slNowDate
                sgMergeLastTimeRun = slNowTime
            End If
            slTime = Format$(gLongToTime(gTimeToLong(sgMergeLastTimeRun, False) + 60 * tgSOE.iMergeChkInterval), "hh:mm:00")
            sgMergedNextDateRun = slNowDate
            sgMergeNextTimeRun = Format$(slTime, "hh:mm:00")
        Else
            sgMergedNextDateRun = DateAdd("d", 1, slNowDate)
            sgMergeNextTimeRun = Format$(tgSOE.sMergeStartTime, "hh:mm:00")
        End If
    End If
    If tgSOE.sMergeStopFlag = "Y" Then
        edcMergeCheck.Text = "Stopped"
    Else
        edcMergeCheck.Text = sgMergedNextDateRun & " " & sgMergeNextTimeRun
    End If
End Sub


Private Sub mSendSEE(slEventCategory As String, slEventAutoCode As String, slDate As String, ilEteCode As Integer, ilLength As Integer, tlSEE As SEE)
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilBDE As Integer
    Dim ilCCE As Integer
    Dim ilTTE As Integer
    Dim ilMTE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilRNE As Integer
    Dim ilFNE As Integer
    Dim ilSCE As Integer
    Dim ilNNE As Integer
    Dim slComment As String
    Dim ilRet As Integer
    Dim llEndTime As Long
    Dim slEndType As String
    
    '9/12/11: Bypass Spots wil Live copy, hard code test for L with copy cart name (Jim)
    If slEventCategory = "S" Then
        If Left(tlSEE.sAudioItemID, 1) = "L" Then
            Exit Sub
        End If
    End If
    slComment = ""
    If slEventCategory = "P" Then
        If tlSEE.l1CteCode > 0 Then
            ilRet = gGetRec_CTE_CommtsTitle(tlSEE.l1CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
            If ilRet Then
                slComment = Trim$(tmCTE.sComment)
            End If
        End If
    ElseIf slEventCategory = "S" Then
        If tlSEE.lAreCode > 0 Then
            ilRet = gGetRec_ARE_AdvertiserRefer(tlSEE.lAreCode, "EngrSchd-mMoveSEERecToCtrls: Advertiser", tmARE)
            If ilRet Then
                slComment = Trim$(tmARE.sName)
            End If
        End If
    End If
    smExportStr = String(ilLength, " ")
    slStr = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If tlSEE.iBdeCode = tgCurrBDE(ilBDE).iCode Then
            slStr = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    mMakeExportStr tgStartColAFE.iBus, tgNoCharAFE.iBus, BUSNAMEINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        If tlSEE.iBusCceCode = tgCurrBusCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iBusControl, tgNoCharAFE.iBusControl, BUSCTRLINDEX, True, ilEteCode, slStr
    If slEventCategory = "P" Then
        slStr = gLongToStrTimeInTenth(tlSEE.lTime)
    Else
        slStr = gLongToStrTimeInTenth(tlSEE.lSpotTime)
    End If
    mMakeExportStr tgStartColAFE.iTime, tgNoCharAFE.iTime, TIMEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
        If tlSEE.iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
            slStr = Trim$(tgCurrStartTTE(ilTTE).sName)
            Exit For
        End If
    Next ilTTE
    mMakeExportStr tgStartColAFE.iStartType, tgNoCharAFE.iStartType, STARTTYPEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
        If tlSEE.iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
            slStr = Trim$(tgCurrEndTTE(ilTTE).sName)
            Exit For
        End If
    Next ilTTE
    mMakeExportStr tgStartColAFE.iEndType, tgNoCharAFE.iEndType, ENDTYPEINDEX, False, ilEteCode, slStr
    slEndType = slStr
    ''12/11/12: Show Duration of zero as 00:00:00.0
    ''If (tlSEE.lDuration > 0) Then
    '2/22/13: Don't show duration if zero and End Type = MAN or EXT
    'If (tlSEE.lDuration >= 0) Then
    If (tlSEE.lDuration > 0) Or ((tlSEE.lDuration = 0) And (Trim$(slEndType) <> "MAN") And (Trim$(slEndType) <> "EXT")) Then
        slStr = gLongToStrLengthInTenth(tlSEE.lDuration, True)
    Else
        slStr = ""
    End If
    mMakeExportStr tgStartColAFE.iDuration, tgNoCharAFE.iDuration, DURATIONINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
        If tlSEE.iMteCode = tgCurrMTE(ilMTE).iCode Then
            slStr = Trim$(tgCurrMTE(ilMTE).sName)
            Exit For
        End If
    Next ilMTE
    mMakeExportStr tgStartColAFE.iMaterialType, tgNoCharAFE.iMaterialType, MATERIALINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        If tlSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
            For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    slStr = Trim$(tgCurrANE(ilANE).sName)
                End If
            Next ilANE
            Exit For
        End If
    Next ilASE
    mMakeExportStr tgStartColAFE.iAudioName, tgNoCharAFE.iAudioName, AUDIONAMEINDEX, True, ilEteCode, slStr
    slStr = Trim$(tlSEE.sAudioItemID)
    mMakeExportStr tgStartColAFE.iAudioItemID, tgNoCharAFE.iAudioItemID, AUDIOITEMIDINDEX, False, ilEteCode, slStr
    slStr = Trim$(tlSEE.sAudioISCI)
    mMakeExportStr tgStartColAFE.iAudioISCI, tgNoCharAFE.iAudioISCI, AUDIOISCIINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iAudioControl, tgNoCharAFE.iAudioControl, AUDIOCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tlSEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
            slStr = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    mMakeExportStr tgStartColAFE.iBkupAudioName, tgNoCharAFE.iBkupAudioName, BACKUPNAMEINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iBkupAudioControl, tgNoCharAFE.iBkupAudioControl, BACKUPCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tlSEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
            slStr = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    mMakeExportStr tgStartColAFE.iProtAudioName, tgNoCharAFE.iProtAudioName, PROTNAMEINDEX, True, ilEteCode, slStr
    slStr = Trim$(tlSEE.sProtItemID)
    mMakeExportStr tgStartColAFE.iProtItemID, tgNoCharAFE.iProtItemID, PROTITEMIDINDEX, False, ilEteCode, slStr
    slStr = Trim$(tlSEE.sProtISCI)
    mMakeExportStr tgStartColAFE.iProtISCI, tgNoCharAFE.iProtISCI, PROTISCIINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iProtAudioControl, tgNoCharAFE.iProtAudioControl, PROTCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        If tlSEE.i1RneCode = tgCurrRNE(ilRNE).iCode Then
            slStr = Trim$(tgCurrRNE(ilRNE).sName)
            Exit For
        End If
    Next ilRNE
    mMakeExportStr tgStartColAFE.iRelay1, tgNoCharAFE.iRelay1, RELAY1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        If tlSEE.i2RneCode = tgCurrRNE(ilRNE).iCode Then
            slStr = Trim$(tgCurrRNE(ilRNE).sName)
            Exit For
        End If
    Next ilRNE
    mMakeExportStr tgStartColAFE.iRelay2, tgNoCharAFE.iRelay2, RELAY2INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
        If tlSEE.iFneCode = tgCurrFNE(ilFNE).iCode Then
            slStr = Trim$(tgCurrFNE(ilFNE).sName)
            Exit For
        End If
    Next ilFNE
    mMakeExportStr tgStartColAFE.iFollow, tgNoCharAFE.iFollow, FOLLOWINDEX, False, ilEteCode, slStr
    If tlSEE.lSilenceTime > 0 Then
        slStr = gLongToLength(tlSEE.lSilenceTime, False)    'gLongToStrLengthInTenth(tlSEE.lSilenceTime, False)
    Else
        slStr = ""
    End If
    mMakeExportStr tgStartColAFE.iSilenceTime, tgNoCharAFE.iSilenceTime, SILENCETIMEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i1SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence1, tgNoCharAFE.iSilence1, SILENCE1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i2SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence2, tgNoCharAFE.iSilence2, SILENCE2INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i3SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence3, tgNoCharAFE.iSilence3, SILENCE3INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i4SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence4, tgNoCharAFE.iSilence4, SILENCE4INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tlSEE.iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            slStr = Trim$(tgCurrNNE(ilNNE).sName)
            Exit For
        End If
    Next ilNNE
    mMakeExportStr tgStartColAFE.iStartNetcue, tgNoCharAFE.iStartNetcue, NETCUE1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tlSEE.iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            slStr = Trim$(tgCurrNNE(ilNNE).sName)
            Exit For
        End If
    Next ilNNE
    mMakeExportStr tgStartColAFE.iStopNetcue, tgNoCharAFE.iStopNetcue, NETCUE2INDEX, False, ilEteCode, slStr
    If (slEventCategory = "P") Then
        mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, TITLE1INDEX, False, ilEteCode, slComment
    Else
        mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, TITLE1INDEX, False, ilEteCode, slComment
    End If
    slStr = ""
    If tlSEE.l2CteCode > 0 Then
        ilRet = gGetRec_CTE_CommtsTitle(tlSEE.l2CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
        If ilRet Then
            slStr = Trim$(tmCTE.sComment)
        End If
    End If
    mMakeExportStr tgStartColAFE.iTitle2, tgNoCharAFE.iTitle2, TITLE2INDEX, False, ilEteCode, slStr
    If sgClientFields = "A" Then
        slStr = Trim$(tlSEE.sABCFormat)
        mMakeExportStr tgStartColAFE.iABCFormat, tgNoCharAFE.iABCFormat, ABCFORMATINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCPgmCode)
        mMakeExportStr tgStartColAFE.iABCPgmCode, tgNoCharAFE.iABCPgmCode, ABCPGMCODEINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCXDSMode)
        mMakeExportStr tgStartColAFE.iABCXDSMode, tgNoCharAFE.iABCXDSMode, ABCXDSMODEINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCRecordItem)
        mMakeExportStr tgStartColAFE.iABCRecordItem, tgNoCharAFE.iABCRecordItem, ABCRECORDITEMINDEX, False, ilEteCode, slStr
    End If
    'Event Type
    'If mColOk(ilEteCode, EVENTTYPEINDEX) Then
        If tgStartColAFE.iEventType > 0 Then
            slStr = slEventAutoCode
            Do While Len(slStr) < tgNoCharAFE.iEventType
                slStr = slStr & " "
            Loop
            Mid(smExportStr, tgStartColAFE.iEventType, tgNoCharAFE.iEventType) = slStr
        End If
    'End If
    'Fixed
    If gExportCol(ilEteCode, FIXEDINDEX) Then
        slStr = Trim$(tlSEE.sFixedTime)
        If slStr = "Y" Then
            If tgStartColAFE.iFixedTime > 0 Then
                slStr = Trim$(tgAEE.sFixedTimeChar)
                Do While Len(slStr) < tgNoCharAFE.iFixedTime
                    slStr = slStr & " "
                Loop
                Mid(smExportStr, tgStartColAFE.iFixedTime, tgNoCharAFE.iFixedTime) = slStr
            End If
        End If
    End If
    'Date
    If tgStartColAFE.iDate > 0 Then
        slStr = slDate
        Do While Len(slStr) < tgNoCharAFE.iDate
            slStr = slStr & " "
        Loop
        Mid(smExportStr, tgStartColAFE.iDate, tgNoCharAFE.iDate) = slStr
    End If
    'End Time
    If gExportCol(ilEteCode, DURATIONINDEX) Then
        If tgStartColAFE.iEndTime > 0 Then
            '2/22/13: Don't show Out Time if duration is zero and End Type = MAN or EXT
            If (tlSEE.lDuration > 0) Or ((tlSEE.lDuration = 0) And (Trim$(slEndType) <> "MAN") And (Trim$(slEndType) <> "EXT")) Then
                If slEventCategory = "P" Then
                    llEndTime = tlSEE.lTime + tlSEE.lDuration
                Else
                    llEndTime = tlSEE.lSpotTime + tlSEE.lDuration
                End If
                If llEndTime > 864000 Then
                    llEndTime = llEndTime - 864000
                End If
                slStr = gLongToStrLengthInTenth(llEndTime, True)
            Else
                slStr = ""
            End If
            Do While Len(slStr) < tgNoCharAFE.iEndTime
                slStr = slStr & " "
            Loop
            Mid(smExportStr, tgStartColAFE.iEndTime, tgNoCharAFE.iEndTime) = slStr
        End If
    End If
    'Event ID
    If tgStartColAFE.iEventID > 0 Then
        slStr = Trim$(Str$(tlSEE.lEventID))
        Do While Len(slStr) < tgNoCharAFE.iEventID
            slStr = "0" & slStr
        Loop
        Mid(smExportStr, tgStartColAFE.iEventID, tgNoCharAFE.iEventID) = slStr
    End If
    Print #hmExport, smExportStr

End Sub

Private Sub mSpotMatch(tlSEE As SEE)
'    Dim ilLoop As Integer
'
'    tlSEE.lCode = 0
'    For ilLoop = 0 To UBound(tgSpotCurrSEE) - 1 Step 1
'        If tlSEE.iBdeCode = tgSpotCurrSEE(ilLoop).iBdeCode Then
'            If tlSEE.lTime = tgSpotCurrSEE(ilLoop).lTime Then
'                If tlSEE.lDuration = tgSpotCurrSEE(ilLoop).lDuration Then
'                    If tlSEE.lAreCode = tgSpotCurrSEE(ilLoop).lAreCode Then
'                        tlSEE.lCode = tgSpotCurrSEE(ilLoop).lCode
'                        tlSEE.sAudioItemIDChk = tgSpotCurrSEE(ilLoop).sAudioItemIDChk
'                        tlSEE.sProtItemIDChk = tgSpotCurrSEE(ilLoop).sProtItemIDChk
'                        Exit Sub
'                    End If
'                End If
'            End If
'        End If
'    Next ilLoop
End Sub

Private Function mExportRow(ilEteCode As Integer, slEventCategory As String, slEventAutoCode As String) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    
    mExportRow = False
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).iCode = ilEteCode Then
            slEventCategory = tgCurrETE(ilETE).sCategory
            slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
            If tgCurrETE(ilETE).sCategory = "A" Then
                Exit Function
            End If
            For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                If tgCurrEPE(ilEPE).sType = "E" Then
                    If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                        If tgCurrEPE(ilEPE).sBus = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBusControl = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                        'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                        If tgCurrEPE(ilEPE).sTime = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStartType = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sFixedTime = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sEndType = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sDuration = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sMaterialType = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioName = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioItemID = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioISCI = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioControl = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBkupAudioName = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBkupAudioControl = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioName = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioItemID = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioISCI = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioControl = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sRelay1 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sRelay2 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sFollow = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilenceTime = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence1 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence2 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence3 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence4 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStartNetcue = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStopNetcue = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sTitle1 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sTitle2 = "Y" Then
                            mExportRow = True
                            Exit Function
                        End If
                        If (sgClientFields = "A") Then
                            If tgCurrEPE(ilEPE).sABCFormat = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCPgmCode = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCXDSMode = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCRecordItem = "Y" Then
                                mExportRow = True
                                Exit Function
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next ilEPE
            Exit For
        End If
    Next ilETE
End Function




Private Function mLoadAsAirLog() As Integer
    Dim slFromFile As String
    Dim slDateTime As String
    Dim slName As String
    Dim slPath As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilRet As Integer
    Dim slAsAirDate As String
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slLine As String
    Dim ilEof As Integer
    Dim slDrive As String
    Dim slStr As String
    Dim slDrivePath As String
    
    On Error GoTo mLoadAsAirLogErr:
    
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    smFileNames = ""
    ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrImportAsAir-mInit", tgCurrAPE())
    For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
        If ((tgCurrAPE(ilLoop).sType = "CI") And (igRunningFrom = 1)) Or ((tgCurrAPE(ilLoop).sType = "SI") And (igRunningFrom = 0)) Then
            smFileNames = Trim$(tgCurrAPE(ilLoop).sNewFileName) & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            ilPos = InStr(1, smFileNames, "Date", vbTextCompare)
            If ilPos > 0 Then
                If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                    slDate = Format$(sgAsAirLogDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                Else
                    slDate = Format$(sgAsAirLogDate, "yymmdd")
                End If
                smFileNames = Left$(smFileNames, ilPos - 1) & slDate & Mid(smFileNames, ilPos + 4)
            End If
            slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            If slPath <> "" Then
                If right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            lbcLogFile.Pattern = "*" & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
            Exit For
        End If
    Next ilLoop
    ilPos = InStr(slPath, ":")
    If ilPos > 0 Then
        slDrive = Left$(slPath, ilPos)
        slPath = Mid$(slPath, ilPos + 1)
        If right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        cbcLogDrive.Drive = slDrive
        lbcLogPath.Path = slPath
        slStr = lbcLogPath.Path
        If right$(slStr, 1) <> "\" Then
            slStr = slStr & "\"
        End If
        lbcLogFile.fileName = slStr & smFileNames
    ElseIf Left(slPath, 2) = "\\" Then
        ilPos = InStr(3, slPath, "\", vbTextCompare)
        ilPos = InStr(ilPos + 1, slPath, "\", vbTextCompare)
        slDrive = Left$(slPath, ilPos - 1)
        'slPath = Mid$(slPath, ilPos + 1)
        If right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        End If
        cbcLogDrive.Drive = slDrive
        lbcLogPath.Path = slPath
        slStr = lbcLogPath.Path
        If right$(slStr, 1) <> "\" Then
            slStr = slStr & "\"
        End If
        lbcLogFile.fileName = slStr & smFileNames
    End If
    If lbcLogFile.ListCount <= 0 Then
        btrDestroy hmSEE
        mLoadAsAirLog = True
        Exit Function
    End If
    
    
    slDrivePath = lbcLogPath.Path
    ReDim smRenameFile(0 To 0) As String
    For ilLoop = 0 To lbcLogFile.ListCount - 1 Step 1
        slName = lbcLogFile.List(ilLoop)
        
        If tgNoCharAFE.iDate = 8 Then
            slAsAirDate = Mid$(slName, 5, 2) & "/" & Mid$(slName, 7, 2) & "/" & Left$(slName, 4)
        ElseIf tgNoCharAFE.iDate = 6 Then
            slAsAirDate = Mid$(slName, 3, 2) & "/" & Mid$(slName, 5, 2) & "/" & Left$(slName, 2)
        End If
        ilRet = 0
        slFromFile = slDrivePath & slName
        slDateTime = FileDateTime(slFromFile)
        If ilRet = 0 Then
            ilRet = gLoadAsAirLog(slFromFile, slAsAirDate, hmSEE)
            If ilRet Then
                smRenameFile(UBound(smRenameFile)) = slDrivePath & slName
                ReDim Preserve smRenameFile(0 To UBound(smRenameFile) + 1) As String
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(smRenameFile) - 1 Step 1
        slName = smRenameFile(ilLoop)
        ilPos = InStr(1, slName, ".", vbTextCompare)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos) & "Old"
        End If
        ilRet = 0
        slDateTime = FileDateTime(slName)
        If ilRet = 0 Then
            Kill slName
        End If
        Name smRenameFile(ilLoop) As slName
    Next ilLoop
    mLoadAsAirLog = True
    btrDestroy hmSEE
    On Error GoTo 0
    Exit Function
mLoadAsAirLogErr:
    ilRet = 1
    Resume Next
End Function



Private Sub mSortTime(tlSEE() As SEE)
    Dim llSEE As Long
    Dim slEventCategory As String
    Dim slTime As String
    Dim slBusName As String
    Dim slSpotTime As String
    Dim ilETE As Integer
    Dim ilBDE As Integer
    
    ReDim tmSeeTimeSort(0 To UBound(tlSEE)) As SEETIMESORT
    For llSEE = 0 To UBound(tlSEE) - 1 Step 1
        slTime = Trim$(Str$(tlSEE(llSEE).lTime))
        Do While Len(slTime) < 10
            slTime = "0" & slTime
        Loop
        slBusName = ""
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            If tlSEE(llSEE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                slBusName = Trim$(tgCurrBDE(ilBDE).sName)
                Exit For
            End If
        Next ilBDE
        Do While Len(slBusName) < 10
            slBusName = slBusName & " "
        Loop
        slSpotTime = Trim$(Str$(tlSEE(llSEE).lSpotTime))
        Do While Len(slSpotTime) < 10
            slSpotTime = "0" & slSpotTime
        Loop
        slEventCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tlSEE(llSEE).iEteCode = tgCurrETE(ilETE).iCode Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If (slEventCategory = "S") Then
            tmSeeTimeSort(llSEE).sKey = slSpotTime & slBusName
        Else
            tmSeeTimeSort(llSEE).sKey = slTime & slBusName
        End If
        tmSeeTimeSort(llSEE).tSEE = tlSEE(llSEE)
    Next llSEE
    'Sort by Time
    If UBound(tmSeeTimeSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmSeeTimeSort(), 0), UBound(tmSeeTimeSort), 0, LenB(tmSeeTimeSort(0)), 0, LenB(tmSeeTimeSort(0).sKey), 0
    End If
    For llSEE = 0 To UBound(tmSeeTimeSort) - 1 Step 1
        tlSEE(llSEE) = tmSeeTimeSort(llSEE).tSEE
    Next llSEE
End Sub

Private Sub mUpdateServiceTime()
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    
    If tgMie.lCode <= 0 Then
        Exit Sub
    End If
    slNowDate = Format(Now, sgShowDateForm)  'Format(gNow(), sgShowDateForm)
    slNowTime = Format(Now, sgShowTimeWSecForm)  'Format(gNow(), sgShowTimeWSecForm)
    tgMie.sEnteredDate = Format$(slNowDate, sgShowDateForm)
    tgMie.sEnteredTime = Format$(slNowTime, sgShowTimeWSecForm)
    ilRet = gPutUpdate_MIE_MessageInfo(tgMie, "EngrServiveMain")
End Sub

Private Sub mArchiveLoad()
    Dim ilPass As Integer
    Dim ilAPE As Integer
    Dim ilFile As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim llFileDate As Long
    Dim llNowMinus7 As Long
    Dim slArchivePath As String
    Dim slPath As String
    Dim slName As String
    Dim slExt As String
    Dim ilYearS As Integer
    Dim ilNoYear As Integer
    Dim ilMnthS As Integer
    Dim ilNoMnth As Integer
    Dim ilDayS As Integer
    Dim ilNoDay As Integer
    Dim ilPos As Integer
    Dim slName0 As String
    Dim slExt0 As String
    Dim ilYearS0 As Integer
    Dim ilNoYear0 As Integer
    Dim ilMnthS0 As Integer
    Dim ilNoMnth0 As Integer
    Dim ilDayS0 As Integer
    Dim ilNoDay0 As Integer
    Dim slName1 As String
    Dim slExt1 As String
    Dim ilYearS1 As Integer
    Dim ilNoYear1 As Integer
    Dim ilMnthS1 As Integer
    Dim ilNoMnth1 As Integer
    Dim ilDayS1 As Integer
    Dim ilNoDay1 As Integer
    Dim ilIndex As Integer
    Dim slLocations As String
    
    On Error Resume Next

    If Not igTestSystem Then
        slLocations = "Locations"
    Else
        slLocations = "TestLocations"
    End If
    If igRunningFrom = 0 Then
        If Not gLoadOption(slLocations, "ServerArchivePath", slArchivePath) Then
            Exit Sub
        End If
    Else
        If Not gLoadOption(slLocations, "ArchivePath", slArchivePath) Then
            Exit Sub
        End If
    End If
    slArchivePath = gSetPathEndSlash(slArchivePath)
    llNowMinus7 = gDateValue(Format(Now, "ddddd")) - 7
    slPath = ""
    For ilAPE = 0 To UBound(tgCurrAPE) - 1 Step 1
        If ((tgCurrAPE(ilAPE).sType = "CE") And (igRunningFrom = 1)) Or ((tgCurrAPE(ilAPE).sType = "SE") And (igRunningFrom = 0)) Then
            If (Not igTestSystem) And ((tgCurrAPE(ilAPE).sSubType = "P") Or (Trim$(tgCurrAPE(ilAPE).sSubType) = "")) Then
                slPath = Trim$(tgCurrAPE(ilAPE).sPath)
            ElseIf (igTestSystem) And (tgCurrAPE(ilAPE).sSubType = "T") Then
                slPath = Trim$(tgCurrAPE(ilAPE).sPath)
            End If
            If slPath <> "" Then
                If right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            If ((tgCurrAPE(ilAPE).sSubType = "P") Or (Trim$(tgCurrAPE(ilAPE).sSubType) = "")) Then
                slName1 = Trim$(tgCurrAPE(ilAPE).sChgFileName)
                slExt1 = "*." & Trim$(tgCurrAPE(ilAPE).sChgFileExt)
                ilYearS1 = -1
                ilNoYear1 = 0
                ilMnthS1 = -1
                ilNoMnth1 = 0
                ilDayS1 = -1
                ilNoDay1 = 0
                ilPos = InStr(1, slName1, "Date", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilAPE).sDateFormat) <> "" Then
                        'slDate = Format$(slAirDate, Trim$(tgCurrAPE(ilAPE).sDateFormat))
                        slStr = Trim$(tgCurrAPE(ilAPE).sDateFormat)
                        For ilIndex = 1 To Len(slStr) Step 1
                            If UCase(Mid(slStr, ilIndex, 1)) = "Y" Then
                                If ilNoYear1 = 0 Then
                                    ilYearS1 = ilIndex
                                End If
                                ilNoYear1 = ilNoYear1 + 1
                            ElseIf UCase(Mid(slStr, ilIndex, 1)) = "M" Then
                                If ilNoMnth1 = 0 Then
                                    ilMnthS1 = ilIndex
                                End If
                                ilNoMnth1 = ilNoMnth1 + 1
                            ElseIf UCase(Mid(slStr, ilIndex, 1)) = "D" Then
                                If ilNoDay1 = 0 Then
                                    ilDayS1 = ilIndex
                                End If
                                ilNoDay1 = ilNoDay1 + 1
                            End If
                        Next ilIndex
                    Else
                        'slDate = Format$(slAirDate, "yymmdd")
                        ilYearS1 = ilPos
                        ilNoYear1 = 2
                        ilMnthS1 = 3
                        ilNoMnth1 = 2
                        ilDayS1 = 5
                        ilNoDay1 = 2
                    End If
                End If
                slName0 = Trim$(tgCurrAPE(ilAPE).sNewFileName)
                slExt0 = "*." & Trim$(tgCurrAPE(ilAPE).sNewFileExt)
                ilYearS0 = -1
                ilNoYear0 = 0
                ilMnthS0 = -1
                ilNoMnth0 = 0
                ilDayS0 = -1
                ilNoDay0 = 0
                ilPos = InStr(1, slName0, "Date", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilAPE).sDateFormat) <> "" Then
                        'slDate = Format$(slAirDate, Trim$(tgCurrAPE(ilAPE).sDateFormat))
                        slStr = Trim$(tgCurrAPE(ilAPE).sDateFormat)
                        For ilIndex = 1 To Len(slStr) Step 1
                            If UCase(Mid(slStr, ilIndex, 1)) = "Y" Then
                                If ilNoYear0 = 0 Then
                                    ilYearS0 = ilIndex
                                End If
                                ilNoYear0 = ilNoYear0 + 1
                            ElseIf UCase(Mid(slStr, ilIndex, 1)) = "M" Then
                                If ilNoMnth0 = 0 Then
                                    ilMnthS0 = ilIndex
                                End If
                                ilNoMnth0 = ilNoMnth0 + 1
                            ElseIf UCase(Mid(slStr, ilIndex, 1)) = "D" Then
                                If ilNoDay0 = 0 Then
                                    ilDayS0 = ilIndex
                                End If
                                ilNoDay0 = ilNoDay0 + 1
                            End If
                        Next ilIndex
                    Else
                        'slDate = Format$(slAirDate, "yymmdd")
                        ilYearS0 = ilPos
                        ilNoYear0 = 2
                        ilMnthS0 = 3
                        ilNoMnth0 = 2
                        ilDayS0 = 5
                        ilNoDay0 = 2
                    End If
                End If
                'Exit For
            End If
        End If
    Next ilAPE
    For ilPass = 0 To 1 Step 1
        If ilPass = 1 Then
            slName = slName1
            slExt = slExt1
            ilYearS = ilYearS1
            ilNoYear = ilNoYear1
            ilMnthS = ilMnthS1
            ilNoMnth = ilNoMnth1
            ilDayS = ilDayS1
            ilNoDay = ilNoDay1
        Else
            slName = slName0
            slExt = slExt0
            ilYearS = ilYearS0
            ilNoYear = ilNoYear0
            ilMnthS = ilMnthS0
            ilNoMnth = ilNoMnth0
            ilDayS = ilDayS0
            ilNoDay = ilNoDay0
        End If
        lbcFile.Path = Left$(slPath, Len(slPath) - 1)
        lbcFile.Pattern = slExt
        For ilFile = 0 To lbcFile.ListCount - 1 Step 1
            On Error GoTo mArchiveFileErr
            ilRet = True
            If (ilYearS = -1) Or (ilMnthS = -1) Or (ilDayS = -1) Then
                slStr = FileDateTime(slPath & lbcFile.List(ilFile))
            Else
                slStr = Mid(lbcFile.List(ilFile), ilMnthS, ilNoMnth) & "/"
                slStr = slStr & Mid(lbcFile.List(ilFile), ilDayS, ilNoDay) & "/"
                slStr = slStr & Mid(lbcFile.List(ilFile), ilYearS, ilNoYear)
            End If
            If ilRet Then
                llFileDate = gDateValue(Format(slStr, "ddddd"))
                If llFileDate < llNowMinus7 Then
                    Name slPath & lbcFile.List(ilFile) As slArchivePath & lbcFile.List(ilFile)
                End If
            End If
        Next ilFile
    Next ilPass
    Exit Sub
mArchiveFileErr:
    ilRet = False
    Resume Next
End Sub

Private Function mIsServiceRunning() As Boolean
'
'mCheckServiceStatus (O)
'
    Dim ilRet As Integer
    Dim llServiceDate As Long
    Dim llServiceTime As Long
    Dim ilCount As Integer
    
    ilCount = 0
    Do
        Sleep lmSleepTime
        ilRet = gGetServiceStatus_MIE_MessageInfo("Main", tgMie)
        If (ilRet) And (tgMie.lCode > 0) Then
            llServiceDate = gDateValue(tgMie.sEnteredDate)
            llServiceTime = gTimeToLong(tgMie.sEnteredTime, True)
            If lgLastServiceTime < 0 Then
                lgLastServiceDate = llServiceDate
                lgLastServiceTime = llServiceTime
            ElseIf lgLastServiceTime <> llServiceTime Then
                If lgLastServiceTime + 300 >= llServiceTime Then
                    mIsServiceRunning = True
                    Exit Function
                End If
                If (lgLastServiceTime > 85800) And (llServiceTime < 600) Then
                    mIsServiceRunning = True
                    Exit Function
                End If
                If lm1970 <> llServiceDate Then
                    mIsServiceRunning = True
                    Exit Function
                End If
            End If
        End If
        ilCount = ilCount + 1
    Loop While ilCount < 120
    mIsServiceRunning = False
End Function


