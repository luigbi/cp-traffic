VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MonitorSetup 
   Caption         =   "Monitor Setup"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   Icon            =   "MonitorSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   10665
   Begin VB.TextBox edcDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   3540
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox pbcToggle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   3180
      ScaleHeight     =   180
      ScaleWidth      =   765
      TabIndex        =   5
      Top             =   1695
      Visible         =   0   'False
      Width           =   765
   End
   Begin V81MonitorSetup.CSI_DayPicker dpcDay 
      Height          =   210
      Left            =   1410
      TabIndex        =   6
      Top             =   2370
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   370
      BorderStyle     =   1
      CSI_ShowSelectRangeButtons=   -1  'True
      CSI_AllowMultiSelection=   -1  'True
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_DayOnColor  =   4638790
      CSI_DayOffColor =   -2147483633
      CSI_RangeFGColor=   0
      CSI_RangeBGColor=   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   105
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   120
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   4815
      Width           =   60
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   120
      Picture         =   "MonitorSetup.frx":08CA
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10230
      Top             =   3045
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10230
      Top             =   3480
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   10260
      Top             =   3990
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5205
      FormDesignWidth =   10665
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3045
      TabIndex        =   0
      Top             =   4650
      Width           =   1890
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5730
      TabIndex        =   1
      Top             =   4650
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSetup 
      Height          =   4155
      Left            =   345
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   300
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
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
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "MonitorSetup"
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
Private imCtrlVisible As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on
Private imFieldChgd As Integer

Private imGuideUstCode As Integer
Private bmNoPervasive As Boolean

Private imSaving As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer

Private imLastColSorted As Integer
Private imLastSort As Integer
Private imColPos(0 To 8) As Integer 'Save column position because of merge

Private lm1970 As Long

Private smService As String    'C=CSI_Service; T=Task Scheduler; N=None
Private smRun As String        'P=Periodic; C=Continuous
Private smPeriod As String     'SE=Standard- End; CE=Calendar- End; IE=Invoice- End

Private Const FORMNAME As String = "MonitorSetup"
Private tmf_rst As ADODB.Recordset

Const TASKCODEINDEX = 0
Const TASKNAMEINDEX = 1
Const SERVICEINDEX = 2
Const RUNINDEX = 3
Const DAILYINDEX = 4
Const PERIODINDEX = 5
Const DAYSAFTERINDEX = 6
Const SORTINDEX = 7
Const TMFCODEINDEX = 8


Private Sub cmcSave_Click()
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If imSaving = True Then
        Exit Sub
    End If
    imSaving = True
    ilRet = mSave()
    imSaving = False
    If ilRet Then
        imFieldChgd = False
        cmcCancel.Caption = "&Done"
    End If
    imTerminate = False
    Exit Sub
cmcSaveErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "MonitorSetup.txt", "MonitorSetup: cmcSave_Click"
End Sub

Private Sub cmcCancel_Click()
    If imSaving Then
        imTerminate = True
        Exit Sub
    End If
    Unload MonitorSetup
End Sub

Private Sub Form_Activate()
    If imFirstTime Then
        'Place here to get correct ColPos instead of in resize as resize calls the routines twice.
        mSetGridColumns
        mSetGridTitles
        imFirstTime = False
    End If
End Sub

Private Sub Form_GotFocus()
    cmcSave.Caption = "&Save"
    cmcCancel.Caption = "&Cancel"
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.5
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts MonitorSetup
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim slAffToWebDate As String
    Dim slWebToAffDate As String
        
    Screen.MousePointer = vbHourglass
    imTerminate = False
    imSaving = False
    imFirstTime = True
    
    mInit
    If Not imTerminate Then
        mInitMonitorSetup
        gInitTaskInfo
        ilRet = mVerifyTask()
        mPopulated
        Screen.MousePointer = vbDefault
    Else
        tmcTerminate.Enabled = True
    End If
End Sub



Private Sub Form_Resize()
    'mSetGridColumns
    'mSetGridTitles
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imSaving Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    tmf_rst.Close
    
    cnn.Close
    
    Set MonitorSetup = Nothing
    End
End Sub

Private Sub grdSetup_EnterCell()
    mSetShow
End Sub

Private Sub grdSetup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdSetup.TopRow
    grdSetup.Redraw = False
End Sub

Private Sub grdSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    grdSetup.ToolTipText = ""
    If (grdSetup.MouseRow >= grdSetup.FixedRows) And (grdSetup.TextMatrix(grdSetup.MouseRow, grdSetup.MouseCol)) <> "" Then
        grdSetup.ToolTipText = grdSetup.TextMatrix(grdSetup.MouseRow, grdSetup.MouseCol)
    End If
End Sub

Private Sub grdSetup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim ilType As Integer
    
    If Y < grdSetup.RowHeight(0) Then
        grdSetup.Row = 0   'grdSetup.MouseRow
        grdSetup.Col = grdSetup.MouseCol
'        If grdSetup.CellBackColor = LIGHTBLUE Then
'            gSetMousePointer grdSetup, grdSetup, vbHourglass
'            mStatusSortCol grdSetup.Col
'            grdSetup.Row = 0
'            grdSetup.Col = ABFCODEINDEX
'            gSetMousePointer grdSetup, grdSetup, vbDefault
'        End If
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdSetup, X, Y)
    If Not ilFound Then
        grdSetup.Redraw = True
        On Error Resume Next
        cmcCancel.SetFocus
        Exit Sub
    End If
    
    If Trim$(grdSetup.TextMatrix(grdSetup.Row, TASKCODEINDEX)) = "" Then
        grdSetup.Redraw = True
        On Error Resume Next
        cmcCancel.SetFocus
        Exit Sub
    End If
    If Not mColOk() Then
        grdSetup.Redraw = True
        On Error Resume Next
        cmcCancel.SetFocus
        Exit Sub
    End If
    lmTopRow = grdSetup.TopRow
    grdSetup.Redraw = True
    mEnableBox
End Sub

Private Sub grdSetup_Scroll()
    If grdSetup.Redraw = False Then
        grdSetup.Redraw = True
        grdSetup.TopRow = lmTopRow
        grdSetup.Refresh
        grdSetup.Redraw = False
    End If
    mSetShow
    cmcCancel.SetFocus

End Sub

Private Sub pbcToggle_KeyPress(KeyAscii As Integer)
    If lmEnableCol = SERVICEINDEX Then
        If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            smService = "C"
        ElseIf KeyAscii = Asc("T") Or (KeyAscii = Asc("t")) Then
            smService = "T"
        '7967
        ElseIf KeyAscii = Asc("W") Or (KeyAscii = Asc("w")) Then
            smService = "W"
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            smService = "N"
        End If
    ElseIf lmEnableCol = RUNINDEX Then
        If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            smRun = "C"
        ElseIf KeyAscii = Asc("P") Or (KeyAscii = Asc("p")) Then
            smRun = "P"
        End If
    ElseIf lmEnableCol = PERIODINDEX Then
        If KeyAscii = Asc("S") Or (KeyAscii = Asc("s")) Then
            smPeriod = "SE"
        ElseIf KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            smPeriod = "CE"
        ElseIf KeyAscii = Asc("I") Or (KeyAscii = Asc("i")) Then
            smPeriod = "IE"
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If lmEnableCol = SERVICEINDEX Then
            '7967
'            If smService = "C" Then
'                smService = "T"
'            ElseIf smService = "T" Then
'                smService = "N"
'            ElseIf smService = "N" Then
'                smService = "C"
'            Else
'                smService = "N"
'            End If
            If smService = "C" Then
                smService = "T"
            ElseIf smService = "T" Then
                smService = "W"
            ElseIf smService = "T" Then
                smService = "W"
            ElseIf smService = "W" Then
                smService = "N"
            Else
                smService = "N"
            End If

        ElseIf lmEnableCol = RUNINDEX Then
            If smRun = "P" Then
                smRun = "C"
            ElseIf smRun = "C" Then
                smRun = "P"
            Else
                smRun = "C"
            End If
        ElseIf lmEnableCol = PERIODINDEX Then
            If smPeriod = "SE" Then
                smPeriod = "CE"
            ElseIf smPeriod = "CE" Then
                smPeriod = "IE"
            ElseIf smPeriod = "IE" Then
                smPeriod = "SE"
            Else
                smPeriod = "SE"
            End If
        End If
    End If
    pbcToggle_Paint
End Sub

Private Sub pbcToggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lmEnableCol = SERVICEINDEX Then
'7967
'        If smService = "C" Then
'            smService = "T"
'        ElseIf smService = "T" Then
'            smService = "N"
'        ElseIf smService = "N" Then
'            smService = "C"
'        Else
'            smService = "N"
'        End If
        If smService = "C" Then
            smService = "T"
        ElseIf smService = "T" Then
            smService = "W"
        ElseIf smService = "W" Then
            smService = "N"
        ElseIf smService = "N" Then
            smService = "C"
        Else
            smService = "N"
        End If
    ElseIf lmEnableCol = RUNINDEX Then
        If smRun = "P" Then
            smRun = "C"
        ElseIf smRun = "C" Then
            smRun = "P"
        Else
            smRun = "C"
        End If
    ElseIf lmEnableCol = PERIODINDEX Then
        If smPeriod = "SE" Then
            smPeriod = "CE"
        ElseIf smPeriod = "CE" Then
            smPeriod = "IE"
        ElseIf smPeriod = "IE" Then
            smPeriod = "SE"
        Else
            smPeriod = "SE"
        End If
    End If
    pbcToggle_Paint
End Sub

Private Sub pbcToggle_Paint()
    pbcToggle.Cls
    pbcToggle.CurrentX = 15
    pbcToggle.CurrentY = 0 'fgBoxInsetY
    If lmEnableCol = SERVICEINDEX Then
        If smService = "C" Then
            pbcToggle.Print "CSI Service"
        ElseIf smService = "T" Then
            pbcToggle.Print "Task Scheduler"
        '7967
        ElseIf smService = "W" Then
            pbcToggle.Print "Web Service"
        ElseIf smService = "N" Then
            pbcToggle.Print "None"
        Else
            pbcToggle.Print ""
        End If
    ElseIf lmEnableCol = RUNINDEX Then
        If smRun = "P" Then
            pbcToggle.Print "Periodic"
        ElseIf smRun = "C" Then
            pbcToggle.Print "Continuous"
        Else
            pbcToggle.Print ""
        End If
    ElseIf lmEnableCol = PERIODINDEX Then
        If smPeriod = "SE" Then
            pbcToggle.Print "Standard- End"
        ElseIf smPeriod = "CE" Then
            pbcToggle.Print "Calendar- End"
        ElseIf smPeriod = "IE" Then
            pbcToggle.Print "Invoice- End"
        Else
            pbcToggle.Print ""
        End If
    End If
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload MonitorSetup
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
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
        Exit Sub
    Else
        sgLogoPath = gSetPathEndSlash(sgLogoPath, True)
    End If
    
    
    If Not gLoadOption("Database", "Name", sgDatabaseName) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Database] 'Name' key is missing.", vbCritical
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Reports", sgReportDirectory) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'Reports' key is missing.", vbCritical
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Export", sgExportDirectory) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'Export' key is missing.", vbCritical
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Exe", sgExeDirectory) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'Exe' key is missing.", vbCritical
        Exit Sub
    End If
    If Not gLoadOption("Locations", "Logo", sgLogoDirectory) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'Logo' key is missing.", vbCritical
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
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
        Exit Sub
    Else
        sgDBPath = gSetPathEndSlash(sgDBPath, True)
    End If
    
    'Set Message folder
    If Not gLoadOption("Locations", "DBPath", sgMsgDirectory) Then
        imTerminate = True
        gMsgBox "Affiliat.Ini [Locations] 'DBPath' key is missing.", vbCritical
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
        imTerminate = True
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
            gHandleError "MonitorSetup.txt", FORMNAME & "-Form_Load"
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
'  7-10-15  Allow setup to run even if affiliate not used; client must still setup dsn and have affiliat.ini
'        If UCase(rst!spfGUseAffSys) <> "Y" Then
'            imTerminate = True
'            gMsgBox "The Affiliate system has not been activated.  Please call Counterpoint.", vbCritical
'            Exit Sub
'        End If
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
        '        Unload MonitorSetup
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
    ilRet = gPopFormats() 'Moved 7/15/21 due to Change in TTP 10243, where Formats need to Load BEFORE STATIONS
    ilRet = gPopStations()
    ilRet = gPopVehicleOptions()
    ilRet = gPopVehicles()
    ilRet = gPopSellingVehicles()
    ilRet = gPopAdvertisers()
    ilRet = gPopReportNames()
    ilRet = gGetLatestRatecard()
    ilRet = gPopTimeZones()
    ilRet = gPopStates()
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
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tgStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub



Private Sub mSetControl()
    If imFieldChgd Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    
    grdSetup.ColWidth(SORTINDEX) = 0
    grdSetup.ColWidth(TMFCODEINDEX) = 0
    grdSetup.ColWidth(TASKCODEINDEX) = grdSetup.Width * 0.06
    grdSetup.ColWidth(SERVICEINDEX) = grdSetup.Width * 0.1
    grdSetup.ColWidth(RUNINDEX) = grdSetup.Width * 0.1
    grdSetup.ColWidth(DAILYINDEX) = grdSetup.Width * 0.15
    grdSetup.ColWidth(PERIODINDEX) = grdSetup.Width * 0.15
    grdSetup.ColWidth(DAYSAFTERINDEX) = grdSetup.Width * 0.15
    
    grdSetup.ColWidth(TASKNAMEINDEX) = grdSetup.Width - 30  ' - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To DAYSAFTERINDEX Step 1
        If (ilCol <> TASKNAMEINDEX) Then
            grdSetup.ColWidth(TASKNAMEINDEX) = grdSetup.ColWidth(TASKNAMEINDEX) - grdSetup.ColWidth(ilCol)
        End If
    Next ilCol
    gGrid_AlignAllColsLeft grdSetup
    For ilCol = 0 To grdSetup.Cols - 1 Step 1
        imColPos(ilCol) = grdSetup.ColPos(ilCol)
    Next ilCol
End Sub

Private Sub mSetGridTitles()
    Dim ilCol As Integer
    
    'For ilCol = 0 To grdSetup.Cols - 1 Step 1
    '    imColPos(ilCol) = grdSetup.ColPos(ilCol)
    'Next ilCol
    
    grdSetup.TextMatrix(0, TASKCODEINDEX) = "Task"
    grdSetup.TextMatrix(1, TASKCODEINDEX) = "Code"
    grdSetup.TextMatrix(0, TASKNAMEINDEX) = "Task"
    grdSetup.TextMatrix(1, TASKNAMEINDEX) = "Name"
    grdSetup.TextMatrix(0, SERVICEINDEX) = "Service"
    
    grdSetup.TextMatrix(0, RUNINDEX) = "Run"
    grdSetup.TextMatrix(0, DAILYINDEX) = "Days"
    grdSetup.TextMatrix(0, PERIODINDEX) = "Monthly"
    grdSetup.TextMatrix(1, PERIODINDEX) = "Period"
    grdSetup.TextMatrix(0, DAYSAFTERINDEX) = "Monthly"
    grdSetup.TextMatrix(1, DAYSAFTERINDEX) = "Days After"
    
    grdSetup.Row = 0
    grdSetup.MergeCells = 2    'flexMergeRestrictColumns
    grdSetup.MergeRow(0) = True
    
    
    grdSetup.Row = 0
    grdSetup.Col = TASKCODEINDEX
    grdSetup.CellAlignment = 4
    
    grdSetup.Row = 0
    grdSetup.Col = PERIODINDEX
    grdSetup.CellAlignment = 4

End Sub


Private Sub mPopulated()
    Dim slStr As String
    Dim llRow As Long
    Dim llCol As Long
    Dim ilTask As Integer
    
    On Error GoTo ErrHand
    llRow = grdSetup.FixedRows
    SQLQuery = "SELECT * FROM TMF_Task_Monitor"
    Set tmf_rst = cnn.Execute(SQLQuery)
    Do While Not tmf_rst.EOF
        If llRow >= grdSetup.Rows Then
            grdSetup.AddItem ""
        End If
        grdSetup.Row = llRow
        For llCol = TASKCODEINDEX To TASKNAMEINDEX Step 1
            grdSetup.Col = llCol
            grdSetup.CellBackColor = LIGHTYELLOW
        Next llCol
        grdSetup.TextMatrix(llRow, TASKCODEINDEX) = Trim$(tmf_rst!tmfTaskCode)
        grdSetup.TextMatrix(llRow, TASKNAMEINDEX) = Trim$(tmf_rst!tmfTaskName)
        slStr = ""
        smService = Trim$(tmf_rst!tmfService)
        Select Case smService
            Case "C"
                slStr = "CSI Service"
            Case "T"
                slStr = "Task Scheduler"
            '7967
            Case "W"
                slStr = "Web Service"
            Case "N"
                slStr = "None"
                For llCol = RUNINDEX To DAYSAFTERINDEX Step 1
                    grdSetup.Col = llCol
                    grdSetup.CellBackColor = LIGHTYELLOW
                Next llCol
        End Select
        grdSetup.TextMatrix(llRow, SERVICEINDEX) = slStr
        slStr = ""
        smRun = Trim$(tmf_rst!tmfRunMode)
        Select Case smRun
            Case "C"
                slStr = "Continuous"
            Case "P"
                slStr = "Periodic"
        End Select
        grdSetup.TextMatrix(llRow, RUNINDEX) = slStr
        If slStr <> "Continuous" Then
            If Trim$(tmf_rst!tmfMo) <> "" Then
                slStr = "NNNNNNN"
                If tmf_rst!tmfMo = "Y" Then
                    Mid(slStr, 1, 1) = "Y"
                End If
                If tmf_rst!tmfTu = "Y" Then
                    Mid(slStr, 2, 1) = "Y"
                End If
                If tmf_rst!tmfWe = "Y" Then
                    Mid(slStr, 3, 1) = "Y"
                End If
                If tmf_rst!tmfTh = "Y" Then
                    Mid(slStr, 4, 1) = "Y"
                End If
                If tmf_rst!tmfFr = "Y" Then
                    Mid(slStr, 5, 1) = "Y"
                End If
                If tmf_rst!tmfSa = "Y" Then
                    Mid(slStr, 6, 1) = "Y"
                End If
                If tmf_rst!tmfSu = "Y" Then
                    Mid(slStr, 7, 1) = "Y"
                End If
                slStr = gMapDays(slStr)
                grdSetup.TextMatrix(llRow, DAILYINDEX) = slStr
                grdSetup.TextMatrix(llRow, PERIODINDEX) = ""
                grdSetup.TextMatrix(llRow, DAYSAFTERINDEX) = ""
            Else
                grdSetup.TextMatrix(llRow, DAILYINDEX) = ""
                slStr = ""
                smPeriod = Trim$(tmf_rst!tmfMonthPeriod)
                Select Case smPeriod
                    Case "SE"
                        slStr = "Standard- End"
                    Case "CE"
                        slStr = "Calendar- End"
                    Case "IE"
                        slStr = "Invoice- End"
                End Select
                grdSetup.TextMatrix(llRow, PERIODINDEX) = slStr
                If slStr <> "" Then
                    grdSetup.TextMatrix(llRow, DAYSAFTERINDEX) = tmf_rst!tmfDaysAfter
                Else
                    grdSetup.TextMatrix(llRow, DAYSAFTERINDEX) = ""
                End If
            End If
        Else
            For llCol = DAILYINDEX To DAYSAFTERINDEX Step 1
                grdSetup.Col = llCol
                grdSetup.CellBackColor = LIGHTYELLOW
                grdSetup.Text = ""
            Next llCol
        End If
        For ilTask = 0 To UBound(tgTaskInfo) Step 1
            If Trim$(tgTaskInfo(ilTask).sTaskCode) = Trim$(tmf_rst!tmfTaskCode) Then
                grdSetup.TextMatrix(llRow, SORTINDEX) = tgTaskInfo(ilTask).sSortCode
                Exit For
            End If
        Next ilTask
        grdSetup.TextMatrix(llRow, TMFCODEINDEX) = tmf_rst!tmfCode
        llRow = llRow + 1
        tmf_rst.MoveNext
    Loop
    imLastColSorted = -1
    imLastSort = -1
    gGrid_SortByCol grdSetup, TASKCODEINDEX, SORTINDEX, imLastColSorted, imLastSort
    Exit Sub
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mPopulated"
    Resume Next
ErrHand1:
    gHandleError "MonitorSetup.txt", "MonitorSetup-mPopulated"
    Return
End Sub

Private Function mAddTmf(slTaskCode As String, slTaskName As String) As Long
    Dim llTmfCode As Long
    
    On Error GoTo ErrHand
    
    SQLQuery = "Insert Into TMF_Task_Monitor ( "
    SQLQuery = SQLQuery & "tmfCode, "
    SQLQuery = SQLQuery & "tmfTaskCode, "
    SQLQuery = SQLQuery & "tmfTaskName, "
    SQLQuery = SQLQuery & "tmfService, "
    SQLQuery = SQLQuery & "tmfRunMode, "
    SQLQuery = SQLQuery & "tmfRunningDate, "
    SQLQuery = SQLQuery & "tmfRunningTime, "
    SQLQuery = SQLQuery & "tmf1stStartRunDate, "
    SQLQuery = SQLQuery & "tmf1stStartRunTime, "
    SQLQuery = SQLQuery & "tmf1stEndRunDate, "
    SQLQuery = SQLQuery & "tmf1stEndRunTime, "
    SQLQuery = SQLQuery & "tmfStartRunDate, "
    SQLQuery = SQLQuery & "tmfStartRunTime, "
    SQLQuery = SQLQuery & "tmfEndRunDate, "
    SQLQuery = SQLQuery & "tmfEndRunTime, "
    SQLQuery = SQLQuery & "tmfMo, "
    SQLQuery = SQLQuery & "tmfTu, "
    SQLQuery = SQLQuery & "tmfWe, "
    SQLQuery = SQLQuery & "tmfTh, "
    SQLQuery = SQLQuery & "tmfFr, "
    SQLQuery = SQLQuery & "tmfSa, "
    SQLQuery = SQLQuery & "tmfSu, "
    SQLQuery = SQLQuery & "tmfMonthPeriod, "
    SQLQuery = SQLQuery & "tmfDaysAfter, "
    SQLQuery = SQLQuery & "tmfEMailReqDate, "
    SQLQuery = SQLQuery & "tmfEMailReqTime, "
    SQLQuery = SQLQuery & "tmfEMailSentDate, "
    SQLQuery = SQLQuery & "tmfStatus, "
    SQLQuery = SQLQuery & "tmfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & slTaskCode & "', "
    SQLQuery = SQLQuery & "'" & slTaskName & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & "" & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llTmfCode = gInsertAndReturnCode(SQLQuery, "TMF_Task_Monitor", "tmfCode", "Replace")
    If llTmfCode <= 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "MonitorSetup.txt", "MonitorSetup-mAddTmf"
        mAddTmf = -1
    End If
    Exit Function
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mAddTmf"
    Exit Function
ErrHand1:
    gHandleError "MonitorSetup.txt", "MonitorSetup-mAddTmf"
    Return
End Function

Private Function mUpdateTmf(ilGridRow As Integer, llTmfCode As Long) As Integer
    On Error GoTo ErrHand
    
'    SQLQuery = "Update TMF_Task_Monitor Set "
'    SQLQuery = SQLQuery & "tmfCode = " & tlTMF.lCode & ", "
'    SQLQuery = SQLQuery & "tmfTaskCode = '" & gFixQuote(tlTMF.sTaskCode) & "', "
'    SQLQuery = SQLQuery & "tmfTaskName = '" & gFixQuote(tlTMF.sTaskName) & "', "
'    SQLQuery = SQLQuery & "tmfService = '" & gFixQuote(tlTMF.sService) & "', "
'    SQLQuery = SQLQuery & "tmfRunMode = '" & gFixQuote(tlTMF.sRunMode) & "', "
'    SQLQuery = SQLQuery & "tmfRunningDate = '" & Format$(tlTMF.sRunningDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "tmfRunningTime = '" & Format$(tlTMF.sRunningTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & "tmfStartRunDate = '" & Format$(tlTMF.sStartRunDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "tmfStartRunTime = '" & Format$(tlTMF.sStartRunTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & "tmfEndRunDate = '" & Format$(tlTMF.sEndRunDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "tmfEndRunTime = '" & Format$(tlTMF.sEndRunTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & "tmfMo = '" & gFixQuote(tlTMF.sMo) & "', "
'    SQLQuery = SQLQuery & "tmfTu = '" & gFixQuote(tlTMF.sTu) & "', "
'    SQLQuery = SQLQuery & "tmfWe = '" & gFixQuote(tlTMF.sWe) & "', "
'    SQLQuery = SQLQuery & "tmfTh = '" & gFixQuote(tlTMF.sTh) & "', "
'    SQLQuery = SQLQuery & "tmfFr = '" & gFixQuote(tlTMF.sFr) & "', "
'    SQLQuery = SQLQuery & "tmfSa = '" & gFixQuote(tlTMF.sSa) & "', "
'    SQLQuery = SQLQuery & "tmfSu = '" & gFixQuote(tlTMF.sSu) & "', "
'    SQLQuery = SQLQuery & "tmfMonthPeriod = '" & gFixQuote(tlTMF.sMonthPeriod) & "', "
'    SQLQuery = SQLQuery & "tmfDaysAfter = " & tlTMF.iDaysAfter & ", "
'    SQLQuery = SQLQuery & "tmfEMailReqDate = '" & Format$(tlTMF.sEMailReqDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "tmfEMailReqTime = '" & Format$(tlTMF.sEMailReqTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & "tmfEMailSentDate = '" & Format$(tlTMF.sEMailSentDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "tmfUnused = '" & gFixQuote(tlTMF.sUnused) & "' "
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "MonitorSetup.txt", "MonitorSetup-mUpdateTmf"
        mUpdateTmf = False
    End If
    Exit Function
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mUpdateTmf"
    Exit Function
'ErrHand1:
'    gHandleError "MonitorSetup.txt", "MonitorSetup-mUpdateTmf"
'    Return
End Function


Private Function mVerifyTask() As Integer
    Dim ilTask As Integer
    Dim llTmfCode As Long
    
    On Error GoTo ErrHand
    mVerifyTask = False
    For ilTask = 0 To UBound(tgTaskInfo) Step 1
        SQLQuery = "SELECT * FROM TMF_Task_Monitor WHERE (tmfTaskCode = '" & tgTaskInfo(ilTask).sTaskCode & "'" & ")"
        Set tmf_rst = cnn.Execute(SQLQuery)
        If tmf_rst.EOF Then
            llTmfCode = mAddTmf(Trim$(tgTaskInfo(ilTask).sTaskCode), Trim$(tgTaskInfo(ilTask).sTaskName))
        Else
            If StrComp(Trim$(tmf_rst!tmfTaskName), Trim$(tgTaskInfo(ilTask).sTaskName), vbTextCompare) <> 0 Then
                SQLQuery = "Update TMF_Task_Monitor Set "
                SQLQuery = SQLQuery & "tmfTaskName = '" & gFixQuote(tgTaskInfo(ilTask).sTaskName) & "'"
                SQLQuery = SQLQuery + " WHERE (tmfCode = " & tmf_rst!tmfCode & ")"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gHandleError "MonitorSetup.txt", "MonitorSetup-mVerifyTask"
                    mVerifyTask = False
                    Exit Function
                End If
            End If
        End If
    Next ilTask
    mVerifyTask = True
    Exit Function
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mVerifyTask"
    Exit Function
'ErrHand1:
'    gHandleError "MonitorSetup.txt", "MonitorSetup-mUpdateTmf"
'    Return
End Function

Private Sub mInitMonitorSetup()
    imCtrlVisible = False
    imLastColSorted = -1
    imLastSort = -1
    imFieldChgd = False
    lm1970 = gDateValue("1/1/1970")
    pbcSTab.Left = -2 * pbcSTab.Width
    pbcTab.Left = -2 * pbcTab.Width
End Sub
Private Sub pbcTab_GotFocus()
    Dim slStr As String
    Dim ilNext As Integer
    Dim ilTestValue As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        lmEnableRow = llEnableRow
        lmEnableCol = llEnableCol
        'Branch
        ilTestValue = True
        Do
            ilNext = False
            Select Case grdSetup.Col
                Case RUNINDEX
                    If smRun = "C" Then
                        If cmcSave.Enabled Then
                            cmcSave.SetFocus
                        Else
                            cmcCancel.SetFocus
                        End If
                        Exit Sub
                    End If
                    grdSetup.Col = grdSetup.Col + 1
                Case DAYSAFTERINDEX
                    If grdSetup.Rows - 1 = lmEnableRow Then
                        If cmcSave.Enabled Then
                            cmcSave.SetFocus
                        Else
                            cmcCancel.SetFocus
                        End If
                        Exit Sub
                    Else
                        grdSetup.Row = lmEnableRow + 1
                        grdSetup.Col = SERVICEINDEX
                    End If
                Case PERIODINDEX
                    If smPeriod = "" Then
                        If grdSetup.Rows - 1 = lmEnableRow Then
                            If cmcSave.Enabled Then
                                cmcSave.SetFocus
                            Else
                                cmcCancel.SetFocus
                            End If
                            Exit Sub
                        Else
                            grdSetup.Row = lmEnableRow + 1
                            grdSetup.Col = SERVICEINDEX
                        End If
                    Else
                        grdSetup.Col = grdSetup.Col + 1
                    End If
                Case Else
                    grdSetup.Col = grdSetup.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        grdSetup.TopRow = grdSetup.FixedRows
        grdSetup.Col = SERVICEINDEX
        Do
            If grdSetup.Row <= grdSetup.FixedRows Then
                cmcCancel.SetFocus
                Exit Sub
            End If
            grdSetup.Row = grdSetup.Rows - 1
            Do
                If Not grdSetup.RowIsVisible(grdSetup.Row) Then
                    grdSetup.TopRow = grdSetup.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
            If mColOk() Then
                Exit Do
            End If
        Loop
    End If
    lmTopRow = grdSetup.TopRow
    mEnableBox
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        'Branch
        Do
            ilNext = False
            Select Case grdSetup.Col
                Case Else
                    grdSetup.Col = grdSetup.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        lmTopRow = -1
        grdSetup.TopRow = grdSetup.FixedRows
        grdSetup.Row = grdSetup.FixedRows
        grdSetup.Col = SERVICEINDEX
        Do
            If mColOk() Then
                Exit Do
            End If
            If grdSetup.Row + 1 >= grdSetup.Rows Then
                cmcCancel.SetFocus
                Exit Sub
            End If
            grdSetup.Row = grdSetup.Row + 1
            Do
                If Not grdSetup.RowIsVisible(grdSetup.Row) Then
                    grdSetup.TopRow = grdSetup.TopRow + 1
                Else
                    Exit Do
                End If
            Loop
        Loop
    End If
    lmTopRow = grdSetup.TopRow
    mEnableBox
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdSetup.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function

Private Sub mEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    
    If Not imFieldChgd Then
        cmcCancel.Caption = "&Cancel"
    End If
    If (grdSetup.Row >= grdSetup.FixedRows) And (grdSetup.Row < grdSetup.Rows) Then
        lmEnableRow = grdSetup.Row
        lmEnableCol = grdSetup.Col
        imCtrlVisible = True
        grdSetup.CellForeColor = vbBlack
        If grdSetup.Text = "Missing" Then
            grdSetup.Text = ""
        End If
        Select Case grdSetup.Col
            Case SERVICEINDEX
                If grdSetup.Text = "CSI Service" Then
                    smService = "C"
                ElseIf grdSetup.Text = "Task Scheduler" Then
                    smService = "T"
                '7967
                ElseIf grdSetup.Text = "Web Service" Then
                    smService = "W"
                ElseIf grdSetup.Text = "None" Then
                    smService = "N"
                Else
                    smService = "N"
                End If
                pbcToggle.Move grdSetup.Left + imColPos(lmEnableCol) + 30, grdSetup.Top + grdSetup.RowPos(grdSetup.Row) + 15, grdSetup.ColWidth(grdSetup.Col) - 30, grdSetup.RowHeight(grdSetup.Row) - 15
            Case RUNINDEX
                If grdSetup.Text = "Periodic" Then
                    smRun = "P"
                ElseIf grdSetup.Text = "Continuous" Then
                    smRun = "C"
                Else
                    smRun = "C"
                End If
                pbcToggle.Move grdSetup.Left + imColPos(lmEnableCol) + 30, grdSetup.Top + grdSetup.RowPos(grdSetup.Row) + 15, grdSetup.ColWidth(grdSetup.Col) - 30, grdSetup.RowHeight(grdSetup.Row) - 15
            Case DAILYINDEX
                dpcDay.Text = grdSetup.TextMatrix(lmEnableRow, lmEnableCol)
                dpcDay.Move grdSetup.Left + imColPos(lmEnableCol) + 30, grdSetup.Top + grdSetup.RowPos(grdSetup.Row) + 15, grdSetup.ColWidth(grdSetup.Col) - 30, grdSetup.RowHeight(grdSetup.Row) - 15
            Case PERIODINDEX
                If grdSetup.Text = "Standard- End" Then
                    smPeriod = "SE"
                ElseIf grdSetup.Text = "Calendar- End" Then
                    smPeriod = "CE"
                ElseIf grdSetup.Text = "Invoice- End" Then
                    smPeriod = "IE"
                Else
                    smPeriod = ""
                End If
                pbcToggle.Move grdSetup.Left + imColPos(lmEnableCol) + 30, grdSetup.Top + grdSetup.RowPos(grdSetup.Row) + 15, grdSetup.ColWidth(grdSetup.Col) - 30, grdSetup.RowHeight(grdSetup.Row) - 15
            Case DAYSAFTERINDEX
                edcDropdown.Text = grdSetup.TextMatrix(lmEnableRow, lmEnableCol)
                edcDropdown.Move grdSetup.Left + imColPos(lmEnableCol) + 30, grdSetup.Top + grdSetup.RowPos(grdSetup.Row) + 15, grdSetup.ColWidth(grdSetup.Col) - 30, grdSetup.RowHeight(grdSetup.Row) - 15
        End Select
    End If
    mSetFocus
End Sub

Private Sub mSetShow()
    Dim slStr As String
    Dim llSvRow As Long
    Dim llSvCol As Long
    Dim llCol As Long
    
    If (lmEnableRow >= grdSetup.FixedRows) And (lmEnableRow < grdSetup.Rows) Then
        Select Case lmEnableCol
            Case SERVICEINDEX
                If smService = "C" Then
                    slStr = "CSI Service"
                ElseIf smService = "T" Then
                    slStr = "Task Scheduler"
                '7967
                ElseIf smService = "W" Then
                    slStr = "Web Service"
                ElseIf smService = "N" Then
                    slStr = "None"
                Else
                    slStr = ""
                End If
                If slStr <> grdSetup.TextMatrix(lmEnableRow, lmEnableCol) Then
                    imFieldChgd = True
                    llSvRow = grdSetup.Row
                    llSvCol = grdSetup.Col
                    grdSetup.Row = lmEnableRow
                    If slStr = "None" Then
                        For llCol = RUNINDEX To DAYSAFTERINDEX Step 1
                            grdSetup.Col = llCol
                            grdSetup.CellBackColor = LIGHTYELLOW
                            grdSetup.Text = ""
                        Next llCol
                    Else
                        For llCol = RUNINDEX To DAYSAFTERINDEX Step 1
                            grdSetup.Col = llCol
                            grdSetup.CellBackColor = vbWhite
                        Next llCol
                    End If
                    grdSetup.Row = llSvRow
                    grdSetup.Col = llSvCol
                    grdSetup.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case RUNINDEX
                If smRun = "P" Then
                    slStr = "Periodic"
                ElseIf smRun = "C" Then
                    slStr = "Continuous"
                Else
                    slStr = ""
                End If
                If slStr <> grdSetup.TextMatrix(lmEnableRow, lmEnableCol) Then
                    imFieldChgd = True
                    llSvRow = grdSetup.Row
                    llSvCol = grdSetup.Col
                    grdSetup.Row = lmEnableRow
                    If slStr = "Continuous" Then
                        For llCol = DAILYINDEX To DAYSAFTERINDEX Step 1
                            grdSetup.Col = llCol
                            grdSetup.CellBackColor = LIGHTYELLOW
                            grdSetup.Text = ""
                        Next llCol
                    Else
                        For llCol = DAILYINDEX To DAYSAFTERINDEX Step 1
                            grdSetup.Col = llCol
                            grdSetup.CellBackColor = vbWhite
                        Next llCol
                    End If
                    grdSetup.Row = llSvRow
                    grdSetup.Col = llSvCol
                    grdSetup.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case DAILYINDEX
                slStr = dpcDay.Text
                If slStr <> grdSetup.TextMatrix(lmEnableRow, lmEnableCol) Then
                    imFieldChgd = True
                    grdSetup.TextMatrix(lmEnableRow, PERIODINDEX) = ""
                    grdSetup.TextMatrix(lmEnableRow, DAYSAFTERINDEX) = ""
                    grdSetup.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case PERIODINDEX
                If smPeriod = "SE" Then
                    slStr = "Standard- End"
                ElseIf smPeriod = "CE" Then
                    slStr = "Calendar- End"
                ElseIf smPeriod = "IE" Then
                    slStr = "Invoice- End"
                Else
                    slStr = ""
                End If
                If slStr <> grdSetup.TextMatrix(lmEnableRow, lmEnableCol) Then
                    imFieldChgd = True
                    grdSetup.TextMatrix(lmEnableRow, DAILYINDEX) = ""
                    grdSetup.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case DAYSAFTERINDEX
                slStr = edcDropdown.Text
                If Val(slStr) <> Val(grdSetup.TextMatrix(lmEnableRow, lmEnableCol)) Then
                    imFieldChgd = True
                    grdSetup.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
       End Select
    End If
    dpcDay.Visible = False
    pbcToggle.Visible = False
    edcDropdown.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetControl
End Sub
Private Sub mSetFocus()
    Dim ilIndex As Integer
    Dim slStr As String
    If (grdSetup.Row >= grdSetup.FixedRows) And (grdSetup.Row < grdSetup.Rows) And (imCtrlVisible) Then
        imCtrlVisible = True
        Select Case grdSetup.Col
            Case SERVICEINDEX
                pbcToggle.Visible = True
                pbcToggle.SetFocus
            Case RUNINDEX
                pbcToggle.Visible = True
                pbcToggle.SetFocus
            Case DAILYINDEX
                dpcDay.Visible = True
                dpcDay.SetFocus
            Case PERIODINDEX
                pbcToggle.Visible = True
                pbcToggle.SetFocus
            Case DAYSAFTERINDEX
                edcDropdown.Visible = True
                edcDropdown.SetFocus
        End Select
    End If
End Sub


Private Function mSave() As Boolean
    Dim slStr As String
    Dim slDays As String
    Dim llRow As Long
    
    If Not mTestFields() Then
        mSave = False
        Exit Function
    End If
    For llRow = grdSetup.FixedRows To grdSetup.Rows - 1 Step 1
        '4/7/15: Clear values if None specified
        'If (grdSetup.TextMatrix(llRow, SERVICEINDEX) <> "") And (grdSetup.TextMatrix(llRow, SERVICEINDEX) <> "None") Then
        If (grdSetup.TextMatrix(llRow, SERVICEINDEX) <> "") Then
            slStr = grdSetup.TextMatrix(llRow, SERVICEINDEX)
            If slStr = "CSI Service" Then
                smService = "C"
            ElseIf slStr = "Task Scheduler" Then
                smService = "T"
            ElseIf slStr = "None" Then
                smService = "N"
            '7967
            ElseIf slStr = "Web Service" Then
                smService = "W"
            Else
                smService = "C"
            End If
            slStr = grdSetup.TextMatrix(llRow, RUNINDEX)
            If slStr = "Periodic" Then
                smRun = "P"
            ElseIf slStr = "Continuous" Then
                smRun = "C"
            Else
                smRun = ""
            End If
            SQLQuery = "Update TMF_Task_Monitor Set "
            SQLQuery = SQLQuery & "tmfService = '" & gFixQuote(smService) & "', "
            SQLQuery = SQLQuery & "tmfRunMode = '" & gFixQuote(smRun) & "', "
            slStr = grdSetup.TextMatrix(llRow, DAILYINDEX)
            If slStr <> "" Then
                slDays = gCreateDayStr(slStr)
                SQLQuery = SQLQuery & "tmfMo = '" & gFixQuote(Mid(slDays, 1, 1)) & "', "
                SQLQuery = SQLQuery & "tmfTu = '" & gFixQuote(Mid(slDays, 2, 1)) & "', "
                SQLQuery = SQLQuery & "tmfWe = '" & gFixQuote(Mid(slDays, 3, 1)) & "', "
                SQLQuery = SQLQuery & "tmfTh = '" & gFixQuote(Mid(slDays, 4, 1)) & "', "
                SQLQuery = SQLQuery & "tmfFr = '" & gFixQuote(Mid(slDays, 5, 1)) & "', "
                SQLQuery = SQLQuery & "tmfSa = '" & gFixQuote(Mid(slDays, 6, 1)) & "', "
                SQLQuery = SQLQuery & "tmfSu = '" & gFixQuote(Mid(slDays, 7, 1)) & "', "
                SQLQuery = SQLQuery & "tmfMonthPeriod = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfDaysAfter = " & 0 & ", "
            Else
                SQLQuery = SQLQuery & "tmfMo = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfTu = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfWe = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfTh = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfFr = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfSa = '" & gFixQuote("") & "', "
                SQLQuery = SQLQuery & "tmfSu = '" & gFixQuote("") & "', "
                slStr = grdSetup.TextMatrix(llRow, PERIODINDEX)
                If slStr = "Standard- End" Then
                    smPeriod = "SE"
                ElseIf slStr = "Calendar- End" Then
                    smPeriod = "CE"
                ElseIf slStr = "Invoice- End" Then
                    smPeriod = "IE"
                Else
                    smPeriod = ""
                End If
                SQLQuery = SQLQuery & "tmfMonthPeriod = '" & gFixQuote(smPeriod) & "', "
                SQLQuery = SQLQuery & "tmfDaysAfter = " & Val(grdSetup.TextMatrix(llRow, DAYSAFTERINDEX)) & ", "
            End If
            SQLQuery = SQLQuery & "tmfUnused = '" & gFixQuote("") & "' "
            SQLQuery = SQLQuery + " WHERE (tmfCode = " & grdSetup.TextMatrix(llRow, TMFCODEINDEX) & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "MonitorSetup.txt", "MonitorSetup-mSave"
                mSave = False
                Exit Function
            End If
        End If
    Next llRow
    mSave = True
    Exit Function
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mSave"
    Resume Next
ErrHand1:
    gHandleError "MonitorSetup.txt", "MonitorSetup-mSave"
    Return
End Function

Private Function mTestFields() As Boolean
    Dim llRow As Long
    Dim slStr As String
    Dim blError As Boolean
    
    On Error GoTo ErrHand
    blError = False
    For llRow = grdSetup.FixedRows To grdSetup.Rows - 1 Step 1
        grdSetup.Row = llRow
        If (grdSetup.TextMatrix(llRow, SERVICEINDEX) <> "") And (grdSetup.TextMatrix(llRow, SERVICEINDEX) <> "None") Then
            slStr = grdSetup.TextMatrix(llRow, RUNINDEX)
            If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                blError = True
                grdSetup.TextMatrix(llRow, RUNINDEX) = "Missing"
                grdSetup.Col = RUNINDEX
                grdSetup.CellForeColor = vbRed
            End If
            If slStr <> "Continuous" Then
                If ((grdSetup.TextMatrix(llRow, DAILYINDEX) = "") Or (StrComp(grdSetup.TextMatrix(llRow, DAILYINDEX), "Missing", vbTextCompare) = 0)) Then
                    If ((grdSetup.TextMatrix(llRow, PERIODINDEX) = "") Or (StrComp(grdSetup.TextMatrix(llRow, PERIODINDEX), "Missing", vbTextCompare) = 0)) Then
                        blError = True
                        grdSetup.TextMatrix(llRow, DAILYINDEX) = "Missing"
                        grdSetup.Col = DAILYINDEX
                        grdSetup.CellForeColor = vbRed
                    Else
                        slStr = grdSetup.TextMatrix(llRow, DAYSAFTERINDEX)
                        If (slStr = "") Or (StrComp(slStr, "Missing", vbTextCompare) = 0) Then
                            blError = True
                            grdSetup.TextMatrix(llRow, DAYSAFTERINDEX) = "Missing"
                            grdSetup.Col = DAYSAFTERINDEX
                            grdSetup.CellForeColor = vbRed
                        End If
                    End If
                End If
            End If
        End If
    Next llRow
    If blError Then
        mTestFields = False
    Else
        mTestFields = True
    End If
    Exit Function
ErrHand:
    gHandleError "MonitorSetup.Txt", "MonitorSetup-mTestFields"
    mTestFields = False
End Function
