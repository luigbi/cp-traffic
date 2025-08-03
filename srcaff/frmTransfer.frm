VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransfer 
   Caption         =   "Counterpoint Transfer"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize"
      Height          =   330
      Left            =   3885
      TabIndex        =   10
      Top             =   6210
      Width           =   1380
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   330
      Left            =   5700
      TabIndex        =   9
      Top             =   6210
      Width           =   1380
   End
   Begin VB.Timer tmcStart 
      Left            =   405
      Top             =   7005
   End
   Begin VB.Frame frcError 
      Caption         =   "csiTransfer Issue"
      Height          =   2775
      Left            =   10980
      TabIndex        =   4
      Top             =   810
      Width           =   3645
      Begin VB.Label lbcError 
         Caption         =   "Label1"
         Height          =   2085
         Left            =   330
         TabIndex        =   5
         Top             =   390
         Width           =   2925
      End
   End
   Begin VB.Frame frcImport 
      Caption         =   "Import"
      Height          =   4875
      Left            =   915
      TabIndex        =   3
      Top             =   3000
      Width           =   10260
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdImportMain 
         Height          =   1380
         Left            =   360
         TabIndex        =   7
         Top             =   390
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   2434
         _Version        =   393216
         Rows            =   3
         Cols            =   6
         BackColorSel    =   16777215
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         ScrollBars      =   0
         SelectionMode   =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdImportFiles 
         Height          =   3150
         Left            =   4815
         TabIndex        =   8
         Top             =   255
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   5556
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Frame frcExport 
      Caption         =   "frcExport"
      Height          =   4890
      Left            =   270
      TabIndex        =   0
      Top             =   480
      Width           =   10245
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExportMain 
         Height          =   1380
         Left            =   180
         TabIndex        =   1
         Top             =   435
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   2434
         _Version        =   393216
         Rows            =   3
         Cols            =   6
         BackColorSel    =   16777215
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         ScrollBars      =   0
         SelectionMode   =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExportFiles 
         Height          =   3150
         Left            =   4785
         TabIndex        =   6
         Top             =   390
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   5556
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.Timer tmcRun 
      Interval        =   60000
      Left            =   5370
      Top             =   4770
   End
   Begin MSComctlLib.TabStrip tbcMain 
      Height          =   5600
      Left            =   165
      TabIndex        =   2
      Top             =   135
      Width           =   11035
      _ExtentX        =   19473
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   """1"""
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   """2"""
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin CsiTransfer.TelnetTTYClient ttcControl 
      Left            =   6945
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   1085
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "Test Connection"
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuIPumpExport 
            Caption         =   "iPump"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuMarketronExport 
            Caption         =   "Marketron"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuIPumpImport 
            Caption         =   "iPump"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuSeeFile 
         Caption         =   "See File"
      End
      Begin VB.Menu mnuLogAll 
         Caption         =   "Log All"
      End
   End
   Begin VB.Menu MnuForce 
      Caption         =   "Force To Run"
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bmCancelled As Boolean
Private lmInterval As Long
Private myLog As CLogger
Private bmLogAll As Boolean
Private bmSkipAtClose As Boolean
'create new log if cross midnight
Private dmCurrentDay As Date

Private Const ININAME As String = "Transfer.ini"
Private Const LOGNAME As String = "Transfer"
Private Const CLEANLOGS As Integer = 30
Private Const TABIMPORT As Integer = 1
Private Const TABEXPORT As Integer = 2
Private Const INTERVALDEFAULT As Long = 60000
'for grid
Private Const MAININDEXSTATUS As Integer = 0
Private Const MAININDEXNAME As Integer = 1
Private Const MAININDEXPROCESSED As Integer = 2
Private Const MAININDEXPENDING As Integer = 3
Private Const MAININDEXERROR As Integer = 4
Private Const MAININDEXTRANSFER As Integer = 5
Private Const PROCESSALL As Integer = -3
Private Const PROCESSNONE As Integer = -2
Private Enum ProcessChoices
    running
    pending
    paused
    blank
    warning
End Enum
Private imExportViewFile As Integer
Private imImportViewFile As Integer
'change as add new transfer type. Also change 'sendChoices'
Private Const CHOICES As Integer = 4
Private Enum SendChoices
    ipump = 0
    idc = 1
    marketron = 2
    GenericTelNet = 3
End Enum
Private MyExports() As ITransfer
Private myImports() As ITransfer

Private Sub cmcMin_Click()
    Me.WindowState = vbMinimized
End Sub
Private Sub cmdMin_Click()
    Me.WindowState = vbMinimized
End Sub
Private Sub cmdStop_Click()
    bmCancelled = True
    If bmCancelled Then
        Unload Me
    End If
End Sub
Private Sub grdExportMain_Click()
    Dim slStatus As String
    Dim ilRow As Integer
    Dim llCols As Long
    Dim c As Integer
    Dim llColor As Long
    
    ilRow = grdExportMain.MouseRow
    If grdExportMain.MouseCol = MAININDEXSTATUS Then
        slStatus = grdExportMain.TextMatrix(ilRow, MAININDEXSTATUS)
        Select Case slStatus
            'running?
            Case "Waiting", "Pending", "Running", "Warning"
                grdExportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Paused"
            Case "Paused"
                If Len(grdExportMain.TextMatrix(ilRow, MAININDEXPENDING)) > 0 Then
                    If grdExportMain.TextMatrix(ilRow, MAININDEXPENDING) > 0 Then
                        grdExportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Pending"
                    Else
                        ' first to blank, then send to see if need 'warning'
                        grdExportMain.TextMatrix(ilRow, MAININDEXSTATUS) = ""
                        mSetStatus blank, grdExportMain.TextMatrix(ilRow, MAININDEXTRANSFER)
                    End If
                End If
            'block empty rows
            Case ""
                If grdExportMain.TextMatrix(ilRow, MAININDEXPENDING) <> "" Then
                    grdExportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Paused"
                End If
        End Select
    ElseIf Len(grdExportMain.TextMatrix(ilRow, MAININDEXNAME)) > 0 Then
        grdExportFiles.Visible = True
        If ilRow > 0 Then
            With grdExportMain
                imExportViewFile = ilRow - 1
                For c = 1 To .Rows - 1
                    If c <> ilRow Then
                        llColor = LIGHTYELLOW
                    Else
                        llColor = GRAY
                    End If
                     .Row = c
                    For llCols = 0 To .Cols - 1 Step 1
                         .Col = llCols
                         .CellBackColor = llColor
                         If llColor = LIGHTYELLOW And llCols = 0 And grdExportMain.TextMatrix(c, MAININDEXPENDING) <> "" Then
                            .CellBackColor = vbWhite
                        End If
                     Next llCols
                Next c
            End With
            MyExports(imExportViewFile).FillGrid grdExportFiles
        End If
    End If
End Sub

Private Sub grdExportMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdExportMain
        .ToolTipText = ""
        If (.MouseRow >= .FixedRows) And (.TextMatrix(.MouseRow, .MouseCol)) <> "" Then
            .ToolTipText = .TextMatrix(.MouseRow, MAININDEXERROR)
        End If
    End With
End Sub
Private Sub grdImportMain_click()
    Dim slStatus As String
    Dim ilRow As Integer
    Dim llCols As Long
    Dim c As Integer
    Dim llColor As Long
    
    ilRow = grdImportMain.MouseRow
    If grdImportMain.MouseCol = MAININDEXSTATUS Then
        slStatus = grdImportMain.TextMatrix(ilRow, MAININDEXSTATUS)
        Select Case slStatus
            'running?
            Case "Waiting", "Pending", "Running", "Warning"
                grdImportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Paused"
            Case "Paused"
                If Len(grdImportMain.TextMatrix(ilRow, MAININDEXPENDING)) > 0 Then
                    If grdImportMain.TextMatrix(ilRow, MAININDEXPENDING) > 0 Then
                        grdImportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Pending"
                    Else
                        ' first to blank, then send to see if need 'warning'
                        grdImportMain.TextMatrix(ilRow, MAININDEXSTATUS) = ""
                        mSetStatus blank, grdImportMain.TextMatrix(ilRow, MAININDEXTRANSFER)
                    End If
                End If
            'block empty rows
            Case ""
                If grdImportMain.TextMatrix(ilRow, MAININDEXPENDING) <> "" Then
                    grdImportMain.TextMatrix(ilRow, MAININDEXSTATUS) = "Paused"
                End If
        End Select
    ElseIf Len(grdImportMain.TextMatrix(ilRow, MAININDEXNAME)) > 0 Then
        grdImportFiles.Visible = True
        If ilRow > 0 Then
            With grdImportMain
                imImportViewFile = ilRow - 1
                For c = 1 To .Rows - 1
                    If c <> ilRow Then
                        llColor = LIGHTYELLOW
                    Else
                        llColor = GRAY
                    End If
                     .Row = c
                    For llCols = 0 To .Cols - 1 Step 1
                         .Col = llCols
                         .CellBackColor = llColor
                         If llColor = LIGHTYELLOW And llCols = 0 And grdImportMain.TextMatrix(c, MAININDEXPENDING) <> "" Then
                            .CellBackColor = vbWhite
                        End If
                     Next llCols
                Next c
            End With
            myImports(imImportViewFile).FillGrid grdImportFiles
        End If
    End If
End Sub
Private Sub grdImportMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With grdImportMain
        .ToolTipText = ""
        If (.MouseRow >= .FixedRows) And (.TextMatrix(.MouseRow, .MouseCol)) <> "" Then
            .ToolTipText = .TextMatrix(.MouseRow, MAININDEXERROR)
        End If
    End With
End Sub

Private Sub tbcMain_Click()
    If tbcMain.SelectedItem.Index = TABEXPORT Then
        frcExport.Visible = True
        frcImport.Visible = False
    Else
        frcImport.Visible = True
        frcExport.Visible = False
    End If
End Sub

Private Sub tmcRun_Timer()
    tmcRun.Enabled = False
    mRun
End Sub
'menu items
Private Sub MnuForce_Click()
    tmcRun.Enabled = False
    tmcStart.Enabled = False
    mRun
    tmcRun.Enabled = True
End Sub
Private Sub mnuIPumpExport_Click()
    Dim c As Integer
    Dim ilFound As Integer
    
    Screen.MousePointer = vbHourglass
    tmcRun.Enabled = False
    tmcStart.Enabled = False
    ilFound = PROCESSNONE
    For c = 0 To UBound(MyExports)
        If MyExports(c).Name = "iPump" Then
            ilFound = c
            Exit For
        End If
    Next c
    If ilFound > PROCESSNONE Then
        mTestTelNet c
    End If
    Screen.MousePointer = vbDefault
    tmcRun.Enabled = True
End Sub
Private Sub mnuIPumpImport_Click()
    Dim c As Integer
    Dim ilFound As Integer
    
    Screen.MousePointer = vbHourglass
    tmcRun.Enabled = False
    tmcStart.Enabled = False
    ilFound = PROCESSNONE
    For c = 0 To UBound(myImports)
        If myImports(c).Name = "iPump" Then
            ilFound = c
            Exit For
        End If
    Next c
    If ilFound > PROCESSNONE Then
        If myImports(c).Connect Then
            MsgBox "Connected successfully.", vbOKOnly, "TelNet Connection"
            myLog.WriteFacts "User successfully tested connection to " & myImports(c).Name
        Else
            MsgBox "Connection failed.", vbOKOnly, "TelNet Connection"
            myLog.WriteFacts "User could not connect to " & myImports(c).Name & " " & myImports(c).ErrorMessage
        End If
    End If
    Screen.MousePointer = vbDefault
    tmcRun.Enabled = True
End Sub
Private Sub mnuLogAll_Click()
    If mnuLogAll.Checked = vbChecked Then
        mnuLogAll.Checked = False
        bmLogAll = False
    Else
        mnuLogAll.Checked = vbChecked
        bmLogAll = True
    End If
End Sub
Private Sub MnuSeeFile_Click()
On Error Resume Next
    Shell "notepad.exe " & myLog.LogPath, vbNormalFocus
End Sub
Private Sub mAllowMenus(blAllow As Boolean)
    mnuConnection.Enabled = blAllow
    mnuFile.Enabled = blAllow
    MnuForce.Enabled = blAllow
End Sub

'end menu items
Private Sub Form_Load()
    mInit
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRet As Integer
If Not bgCantClose Then
    tmcRun.Enabled = False
    If MsgBox("Stop Counterpoint Transfer?", vbQuestion + vbYesNo, "Stop Service") = vbNo Then
        Cancel = 1
        bmCancelled = False
        tmcRun.Enabled = True
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim slIniPath As String
    Dim slNewValue As String
    Dim c As Integer
    Dim slName As String
    Dim ilTransfer As Integer
    
    If Not bmSkipAtClose Then
        If Not bgCantClose Then
            myLog.WriteFacts "CsiTransfer stopped", True
            slIniPath = gXmlIniPath(ININAME)
            For c = 0 To UBound(MyExports) - 1
                ilTransfer = grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                If grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = "Paused" Then
                    slNewValue = "True"
                Else
                    slNewValue = "False"
                End If
                slName = MyExports(ilTransfer).Name
                gWriteIni slIniPath, slName, "ExportStartPaused", slNewValue
            Next c
            For c = 0 To UBound(myImports) - 1
                ilTransfer = grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                If grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = "Paused" Then
                    slNewValue = "True"
                Else
                    slNewValue = "False"
                End If
                slName = myImports(ilTransfer).Name
                gWriteIni slIniPath, slName, "ImportStartPaused", slNewValue
            Next c
            Set myLog = Nothing
            Erase MyExports
        Else
            Cancel = True
        End If
    Else
    'no ini file
        Set myLog = Nothing
        Erase MyExports
    End If
End Sub
Private Sub mInit()
    Dim slError As String
    Dim slWriteError As String
    
    If App.PrevInstance Then
        MsgBox "Only one copy of Transfer can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        End
    End If
    Me.Width = 11035
    Me.Height = 7600
    tmcRun.Enabled = False
    tmcStart.Enabled = False
    tmcStart.Interval = INTERVALDEFAULT
    frcError.Visible = False
    slError = ""
    slWriteError = ""
    bmSkipAtClose = False
    If App.PrevInstance Then
        End
    End If
    imExportViewFile = -1
    imImportViewFile = -1
   ' mReadCommandLine
    gCenterStdAlone Me
    ' these 3 so can use gXmlIniPath
    sgDbPath = ""
    sgStartupDirectory = CurDir()
    sgStartupDirectory = gSetPathEndSlash(sgStartupDirectory, False)
    sgExeDirectory = sgStartupDirectory
    'for logging. Change most values after ini file read in
    igExportSource = 2
    sgUserName = "CsiTransferQueue"
    'for cleaning log files
    sgImportDirectory = sgStartupDirectory
    sgExportDirectory = sgStartupDirectory
    'only one used: set in mIniValues
    sgMsgDirectory = sgStartupDirectory
    Set myLog = New CLogger
    If mIniValues(slError) Then
        dmCurrentDay = Date
        If Not myLog.isLog Then
            myLog.LogPath = myLog.CreateLogName(sgMsgDirectory & LOGNAME)
        End If
        mInitTabs
        mGetTransfers
        mInitGrid
        mLoadTransfers
        mSetPaused
        mAllowMenus True
        mCleanDataFiles
        'done by mrun...but don't want to process yet
        mUpdateExports
        mTestAllGridExport "", "Pending", True, False
        mTestAllGridImport "", "Pending", True, False
        myLog.CleanThisFolder = messages
        myLog.CleanFolder LOGNAME, , CLEANLOGS
        myLog.WriteFacts " CsiTransfer started.", True
        If Len(slError) > 0 Then
            myLog.WriteWarning slError
        End If
        tmcRun.Interval = lmInterval
        tmcStart.Enabled = True
    Else
        'no processes available, or ini doesn't exist
        If Len(slError) > 0 Then
            slWriteError = "Could not start program: issue with " & ININAME & ": " & slError
        Else
            slWriteError = "Could not start program: issue with " & ININAME
        End If
        lbcError.Caption = slWriteError
        mHideAll
        myLog.WriteError slWriteError
        bmSkipAtClose = True
    End If
End Sub
Private Sub mReadCommandLine()
'    Dim slCommand As String
'    Dim ilPos As String
    
'    slCommand = Command$
'    ilPos = InStr(1, slCommand, "D:/")
'    If ilPos > 0 Then
'        If InStr(ilPos, slCommand, "Debug", vbTextCompare) > 0 Then
'            bmIsDebug = True
'        Else
'            bmIsDebug = False
'        End If
'    End If

End Sub
Private Sub mHideAll()
    frcExport.Visible = False
    frcImport.Visible = False
    tbcMain.Enabled = False
    frcError.Left = frcExport.Left
    frcError.Top = frcExport.Top
    frcError.Visible = True
End Sub
Private Sub mInitTabs()
    With tbcMain
        frcExport.BorderStyle = 0
        frcExport.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        frcImport.BorderStyle = 0
        frcImport.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        .Tabs(TABEXPORT).Selected = True
    End With
    
End Sub
Private Sub mLoadTransfers()
    Dim c As Integer
On Error GoTo ERRORBOX
    With grdExportMain
        .Col = 0
        For c = 0 To UBound(MyExports) - 1
           .TextMatrix(c + 1, MAININDEXNAME) = MyExports(c).Name
           .TextMatrix(c + 1, MAININDEXTRANSFER) = c
           .Row = c + 1
           .CellBackColor = vbWhite
        Next c
    End With
    With grdImportMain
        .Col = 0
        For c = 0 To UBound(myImports) - 1
           .TextMatrix(c + 1, MAININDEXNAME) = MyExports(c).Name
           .TextMatrix(c + 1, MAININDEXTRANSFER) = c
           .Row = c + 1
           .CellBackColor = vbWhite
        Next c
    End With
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mLoadTransfers:" & Err.Description, True, True
End Sub
Private Sub mSetPaused()
    'set as paused from ini file
    Dim c As Integer
    Dim ilTransferIndex As Integer
    
On Error GoTo ERRORBOX
    With grdExportMain
        For c = 0 To UBound(MyExports) - 1
            ilTransferIndex = .TextMatrix(c + 1, MAININDEXTRANSFER)
            If MyExports(ilTransferIndex).StartPaused Then
                .TextMatrix(c + 1, MAININDEXSTATUS) = "Paused"
            End If
        Next c
    End With
    With grdImportMain
        For c = 0 To UBound(myImports) - 1
            ilTransferIndex = .TextMatrix(c + 1, MAININDEXTRANSFER)
            If myImports(ilTransferIndex).StartPaused Then
                .TextMatrix(c + 1, MAININDEXSTATUS) = "Paused"
            End If
        Next c
    End With
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mSetPaused. " & Err.Description, True, True
    Resume Next
End Sub
Private Sub mGetTransfers()
    Dim ilUpper As String
    
On Error GoTo ERRORBOX
    ReDim MyExports(0)
    If Not myIPump Is Nothing Then
        If myIPump.IsLoaded Then
            ilUpper = UBound(MyExports)
            ReDim Preserve MyExports(ilUpper + 1)
            Set MyExports(ilUpper) = myIPump
        End If
    End If
    If Not myGeneric Is Nothing Then
        If myGeneric.IsLoaded Then
            ilUpper = UBound(MyExports)
            ReDim Preserve MyExports(ilUpper + 1)
            Set MyExports(ilUpper) = myGeneric
        End If
    End If
'    If Not myIDC Is Nothing Then
'        If myIDC.IsLoaded Then
'            ilUpper = UBound(MyExports)
'            ReDim Preserve MyExports(ilUpper + 1)
'            Set MyExports(ilUpper) = myIDC
'        End If
'    End If
'    If Not myMarketron Is Nothing Then
'        If myMarketron.IsLoaded Then
'            ilUpper = UBound(MyExports)
'            ReDim Preserve MyExports(ilUpper + 1)
'            Set MyExports(ilUpper) = myMarketron
'        End If
'    End If
    ReDim myImports(0)
    If Not myIPumpImport Is Nothing Then
        If myIPumpImport.IsLoaded Then
            ilUpper = UBound(myImports)
            ReDim Preserve myImports(ilUpper + 1)
            Set myImports(ilUpper) = myIPumpImport
        End If
    End If
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mGetTransfers:" & Err.Description, True, True
End Sub
Private Function mIniValues(slError As String) As Boolean
    Dim slIniPath As String
    Dim ilchoice As SendChoices
    Dim blRet As Boolean
    Dim blChoice As Boolean
    Dim slPath As String
'    Dim blIsTelNet As Boolean
    
On Error GoTo ERRORBOX
  '  blIsTelNet = False
    blRet = False
    slError = ""
    slIniPath = gXmlIniPath(ININAME)
    If LenB(slIniPath) > 0 Then
        'change myLog only if found.
        gLoadFromIni "General", "LogPath", slIniPath, slPath
        If slPath <> NOTFOUND Then
            sgMsgDirectory = gSetPathEndSlash(slPath, False)
            myLog.LogPath = myLog.CreateLogName(sgMsgDirectory & LOGNAME)
        End If
        gLoadFromIni "General", "Interval", slIniPath, slPath
        If slPath <> NOTFOUND Then
    On Error GoTo NUMBERBOX
            lmInterval = CInt(slPath)
            If Not (lmInterval > 0 And lmInterval < 10) Then
                lmInterval = INTERVALDEFAULT
            Else
                lmInterval = lmInterval * 60000
            End If
        Else
            lmInterval = INTERVALDEFAULT
        End If
On Error GoTo ERRORBOX
        For ilchoice = 0 To CHOICES - 1
            Select Case ilchoice
                Case ipump
                    Set myIPump = New CTransferTelNet
                    With myIPump
                         .Name = "iPump"
                         .ExtensionToFind = ".weg"
                         blChoice = .LoadIni(slIniPath)
                         If Not blChoice Then
                            If UCase(.ErrorMessage) <> "NOT FOUND" Then
                                slError = slError & " " & .ErrorMessage
                            End If
                         Else
                            mnuIPumpExport.Enabled = True
                            blRet = True
                         End If
                    End With
                    Set myIPumpImport = New CTransferFTPI
                    With myIPumpImport
                        .Name = "iPump"
                        .ExtensionToFind = ".weg"
                        blChoice = .LoadIni(slIniPath)
                        If Not blChoice Then
                            If UCase(.ErrorMessage) <> "NOT FOUND" Then
                                slError = slError & " " & .ErrorMessage
                            End If
                        Else
                            mnuIPumpImport.Enabled = True
                            blRet = True
                       End If
                    End With
                Case GenericTelNet
                    Set myGeneric = New CTransferTelNet
                    With myGeneric
                        .Name = "GenericTelNet"
                        .ExtensionToFind = "gen"
                        blChoice = .LoadIni(slIniPath)
                        If Not blChoice Then
                            If UCase(.ErrorMessage) <> "NOT FOUND" Then
                                slError = slError & " " & .ErrorMessage
                            End If
                        Else
                            blRet = True
                        End If
                    End With
'                Case idc
'                    'must change!
''                    Set myIDC = New CTransferTelNet
''                    myIDC.Name = "IDC"
''                    blChoice = myIDC.LoadIni(slIniPath)
''                    If Not blChoice Then
''                        slError = slError & " " & myIDC.ErrorMessage
''                    End If
'                Case marketron
'                    Set myIDC = New CTransferTelNet
'                    With myIDC
'                        .Name = "Marketron"
'                        blChoice = .LoadIni(slIniPath)
'                        If Not blChoice Then
'                            slError = slError & " " & .ErrorMessage
'                        End If
'                       ' Set .LogFile = myLog
'                       ' Set .TelNetControl = ttcControl
'                        mnuMarketronExport.Enabled = True
'                    End With
            End Select
            ' 1 true? blRet is true
        Next ilchoice
    Else
        slError = " Cannot find."
    End If
    mIniValues = blRet
    Exit Function
NUMBERBOX:
    lmInterval = INTERVALDEFAULT
    Resume Next
ERRORBOX:
    myLog.WriteError "Error in mIniValues:" & Err.Description, True, True
    blRet = False
End Function
Private Sub mRun()
    Dim dlNow As Date
    
    DoEvents
        'check to see if log needs to write for a new day...did we cross midnight?
    dlNow = Date
    If DateDiff("d", dmCurrentDay, dlNow) > 0 Then
        myLog.WriteFacts "Closing this log file because of change of day.", True
        myLog.LogPath = myLog.CreateLogName(sgMsgDirectory & LOGNAME)
        myLog.CleanFolder LOGNAME, , CLEANLOGS
        myLog.WriteFacts " New log started for change of day.", True
        mCleanDataFiles
    End If

    mUpdateExports
    If bmCancelled Then
        Exit Sub
    End If
    mAllowMenus False
    DoEvents
    mProcessImports
    mProcessExports
    If bmCancelled Then
        Exit Sub
    End If
    mUpdateImports
    DoEvents
    If bmCancelled Then
        Exit Sub
    End If
    mAllowMenus True
    tmcRun.Enabled = True
End Sub
Private Sub mUpdateExports()
    Dim c As Integer
    Dim ilFiles As Integer
    Dim ilIndex As Integer
    Dim slOldDate As String
    
On Error GoTo ERRORBOX
    If bmCancelled Then
        Exit Sub
    End If
    ' use c for number of grid rows.  get actual transfer index from row
    For c = 0 To UBound(MyExports) - 1
        ilIndex = grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
        slOldDate = MyExports(ilIndex).DateStored
        If DateDiff("d", slOldDate, gNow()) <> 0 Then
            'we only refresh if the day has changed
            MyExports(ilIndex).DateStored = gNow()
            ilFiles = MyExports(ilIndex).FilesProcessed(True)
        Else
            ilFiles = MyExports(ilIndex).FilesProcessed(False)
        End If
        grdExportMain.TextMatrix(c + 1, MAININDEXPROCESSED) = ilFiles
        ilFiles = MyExports(ilIndex).FilesWaiting(True)
        grdExportMain.TextMatrix(c + 1, MAININDEXPENDING) = ilFiles
        If ilFiles = 0 Then
            If Len(MyExports(ilIndex).ErrorMessage) > 0 Then
                mSetStatus warning, ilIndex
                mSetExportGridLine c, MAININDEXERROR, MyExports(ilIndex).ErrorMessage
            ElseIf grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = "Pending" Then
                mSetStatus blank, ilIndex
            End If
        End If
    Next c
    If grdExportFiles.Visible And imExportViewFile > -1 Then
        MyExports(imExportViewFile).FillGrid grdExportFiles
    End If
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mUpdateExports:" & Err.Description, True, True
End Sub
Private Sub mUpdateImports()
    Dim c As Integer
    Dim ilFiles As Integer
    Dim ilIndex As Integer
    Dim slOldDate As String
    
On Error GoTo ERRORBOX
    If bmCancelled Then
        Exit Sub
    End If
    For c = 0 To UBound(myImports) - 1
        ilIndex = grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
        slOldDate = myImports(ilIndex).DateStored
         If DateDiff("d", slOldDate, gNow()) <> 0 Then
            'we only refresh if the day has changed
            myImports(ilIndex).DateStored = gNow()
            ilFiles = myImports(ilIndex).FilesProcessed(True)
        Else
            ilFiles = myImports(ilIndex).FilesProcessed(False)
        End If
        grdImportMain.TextMatrix(c + 1, MAININDEXPROCESSED) = ilFiles
        ilFiles = myImports(ilIndex).FilesWaiting(True)
        grdImportMain.TextMatrix(c + 1, MAININDEXPENDING) = ilFiles
        If ilFiles = 0 Then
            If Len(myImports(ilIndex).ErrorMessage) > 0 Then
                mSetStatusImports warning, ilIndex
                mSetImportGridLine c, MAININDEXERROR, myImports(ilIndex).ErrorMessage
            ElseIf grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = "Pending" Then
                mSetStatusImports blank, ilIndex
            End If
        End If
    Next c
    If grdImportFiles.Visible And imImportViewFile > -1 Then
        myImports(imImportViewFile).FillGrid grdImportFiles
    End If
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mUpdateImports:" & Err.Description, True, True
End Sub
Private Sub mProcessImports()
Dim ilRowToProcess As Integer
Dim ilFilesWaiting As Integer
Dim blMarkAsUnread As Boolean
Dim slFileName As String
Dim ilFileRow As Integer
Dim blRet As Boolean
'ftp doesn't have connection errors

On Error GoTo ERRORBOX
    DoEvents
    Screen.MousePointer = vbHourglass
    ilRowToProcess = PROCESSALL
    mSetStatusImports warning
    Do While ilRowToProcess <> PROCESSNONE
        'find one marked as running, or next to run
        ilRowToProcess = mSetStatusImports(running)
        If ilRowToProcess > PROCESSNONE Then
            DoEvents
            'set others to pending
            mSetStatusImports pending
            With myImports(ilRowToProcess)
                ilFilesWaiting = -1
                'process all files for this row
                Do While ilFilesWaiting <> 0
                    DoEvents
                    If imImportViewFile = ilRowToProcess Then
                        slFileName = .NextFile
                        ilFileRow = mFindFileInGridImport(slFileName)
                        If ilFileRow > 0 Then
                            grdImportFiles.TextMatrix(ilFileRow, FILESINDEXSTATUS) = "Processing"
                        End If
                    End If
                    blRet = .Process()
                    If Not blRet Then
                        blMarkAsUnread = True
                        mSetStatusImports warning, ilRowToProcess
                        myLog.WriteWarning .ErrorMessage
                        mSetImportGridLine ilRowToProcess, MAININDEXERROR, .ErrorMessage
                    Else
                        blMarkAsUnread = False
                        myLog.WriteFacts "Files Processed--" & .StatusMessage
                    End If
                    ilFilesWaiting = .FilesWaiting(False)
                    grdImportMain.TextMatrix(ilRowToProcess + 1, MAININDEXPENDING) = ilFilesWaiting
                    If imImportViewFile = ilRowToProcess Then
                        .FillGrid grdImportFiles
                    End If
                    'changed to paused during running.  Stop processing
                    If grdImportMain.TextMatrix(ilRowToProcess + 1, MAININDEXSTATUS) = "Paused" Then
                        ilFilesWaiting = 0
                    End If
                Loop
                mSetStatusImports blank, ilRowToProcess
            End With
        End If
    Loop
    Screen.MousePointer = vbDefault
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mProcessimports:" & Err.Description, True, True
End Sub

Private Sub mProcessExports()

Dim ilRowToProcess As Integer
Dim ilFilesWaiting As Integer
Dim blMarkAsUnread As Boolean
Dim slFileName As String
Dim ilFileRow As Integer
Dim blRet As Boolean
Dim blHaltRow As Boolean

On Error GoTo ERRORBOX
    DoEvents
    Screen.MousePointer = vbHourglass
    ilRowToProcess = PROCESSALL
    mSetStatus warning
    Do While ilRowToProcess <> PROCESSNONE
        'find one marked as running, or next to run
        ilRowToProcess = mSetStatus(running)
        If ilRowToProcess > PROCESSNONE Then
            DoEvents
            blHaltRow = False
            'set others to pending
            mSetStatus pending
            With MyExports(ilRowToProcess)
                ilFilesWaiting = -1
                'process all files for this row
                Do While ilFilesWaiting <> 0
                    DoEvents
                    If imExportViewFile = ilRowToProcess Then
                        slFileName = .NextFile
                        ilFileRow = mFindFileInGridExport(slFileName)
                        If ilFileRow > 0 Then
                            grdExportFiles.TextMatrix(ilFileRow, FILESINDEXSTATUS) = "Processing"
                        End If
                    End If
                    blRet = .Process()
                    If Not blRet Then
                        blMarkAsUnread = True
                        mSetStatus warning, ilRowToProcess
                        If InStr(1, .ErrorMessage, "connect", vbTextCompare) > 0 Then
                            blHaltRow = True
                            myLog.WriteError .ErrorMessage
                        Else
                            myLog.WriteWarning .ErrorMessage
                        End If
                        mSetExportGridLine ilRowToProcess, MAININDEXERROR, .ErrorMessage
                    Else
                        blMarkAsUnread = False
                        myLog.WriteFacts "Files Processed--" & .StatusMessage
                    End If
                    'don't move if connection error!
                    If Not blHaltRow Then
                        If Not .Move(blMarkAsUnread) Then
                            myLog.WriteError .ErrorMessage
                            mSetStatus warning, ilRowToProcess
                            mSetExportGridLine ilRowToProcess, MAININDEXERROR, .ErrorMessage
                        End If
                    End If
                    ilFilesWaiting = .FilesWaiting(False)
                    grdExportMain.TextMatrix(ilRowToProcess + 1, MAININDEXPENDING) = ilFilesWaiting
                    If imExportViewFile = ilRowToProcess Then
                        .FillGrid grdExportFiles
                    End If
                    'changed to paused during running.  Stop processing
                    'stop process for row if connection error
                    If blHaltRow Or grdExportMain.TextMatrix(ilRowToProcess + 1, MAININDEXSTATUS) = "Paused" Then
                        ilFilesWaiting = 0
                    End If
                Loop
                'this time through, don't run files again.
                If blHaltRow = True Then
                    blHaltRow = False
                    mSetExportGridLine ilRowToProcess, MAININDEXSTATUS, "warning"
                'warning and Pause will remain!
                Else
                    mSetStatus blank, ilRowToProcess
                End If
            End With
        End If
    Loop
    Screen.MousePointer = vbDefault
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mProcessExports:" & Err.Description, True, True
End Sub
Private Function mSetExportGridLine(ilTransferIndex As Integer, ilColIndex As Integer, slValue As String)
    Dim c As Integer
    
    For c = 0 To UBound(MyExports) - 1
        If grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER) = ilTransferIndex Then
            grdExportMain.TextMatrix(c + 1, ilColIndex) = slValue
            Exit For
        End If
    Next c
End Function
Private Function mSetImportGridLine(ilTransferIndex As Integer, ilColIndex As Integer, slValue As String)
    Dim c As Integer
    
    For c = 0 To UBound(myImports) - 1
        If grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER) = ilTransferIndex Then
            grdImportMain.TextMatrix(c + 1, ilColIndex) = slValue
            Exit For
        End If
    Next c
End Function
Private Function mFindFileInGridExport(slName As String) As Integer
    Dim ilRow As Integer
    Dim c As Integer
    
    ilRow = 0
    If Len(slName) > 0 Then
        For c = 1 To grdExportFiles.Rows
            If grdExportFiles.TextMatrix(c, FILESINDEXFILENAME) = slName Then
                ilRow = c
                Exit For
            End If
        Next c
    End If
    mFindFileInGridExport = ilRow
End Function
Private Function mFindFileInGridImport(slName As String) As Integer
    Dim ilRow As Integer
    Dim c As Integer
    
    ilRow = 0
    If Len(slName) > 0 Then
        For c = 1 To grdImportFiles.Rows
            If grdImportFiles.TextMatrix(c, FILESINDEXFILENAME) = slName Then
                ilRow = c
                Exit For
            End If
        Next c
    End If
    mFindFileInGridImport = ilRow
End Function
Private Function mTestAllGridExport(slFind As String, slChange As String, blTestFiles As Boolean, Optional blFirstOnly As Boolean = True) As Integer
    Dim ilRet As Integer
    Dim c As Integer
    
On Error GoTo ERRORBOX
    ilRet = PROCESSNONE
    For c = 0 To UBound(MyExports) - 1
        If grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slFind Then
            'files can't be 0
            If blTestFiles Then
                If grdExportMain.TextMatrix(c + 1, MAININDEXPENDING) > 0 Then
                    grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slChange
                    ilRet = grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                    If blFirstOnly Then
                        Exit For
                    End If
                End If
            Else
                grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slChange
                ilRet = grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                If blFirstOnly Then
                    Exit For
                End If
            End If
        End If
    Next c
    mTestAllGridExport = ilRet
    Exit Function
ERRORBOX:
    myLog.WriteError "Error in mTestAllGridExport:" & Err.Description, True, True
    mTestAllGridExport = PROCESSNONE
End Function
Private Function mTestAllGridImport(slFind As String, slChange As String, blTestFiles As Boolean, Optional blFirstOnly As Boolean = True) As Integer
    Dim ilRet As Integer
    Dim c As Integer
    
On Error GoTo ERRORBOX
    ilRet = PROCESSNONE
    For c = 0 To UBound(myImports) - 1
        If grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slFind Then
            'files can't be 0
            If blTestFiles Then
                If grdImportMain.TextMatrix(c + 1, MAININDEXPENDING) > 0 Then
                    grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slChange
                    ilRet = grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                    If blFirstOnly Then
                        Exit For
                    End If
                End If
            Else
                grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slChange
                ilRet = grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER)
                If blFirstOnly Then
                    Exit For
                End If
            End If
        End If
    Next c
    mTestAllGridImport = ilRet
    Exit Function
ERRORBOX:
    myLog.WriteError "Error in mTestAllGridImport:" & Err.Description, True, True
    mTestAllGridImport = PROCESSNONE
End Function
Private Function mSetStatus(ilProcess As ProcessChoices, Optional ilTransferIndex As Integer = PROCESSNONE) As Integer
    'return the transfer index!
    'in all cases below, must have files!
    ' Running: change waiting to running. If not waiting, change first pending.  No pending?  First blank. After blank, warning
    ' Pending: change all blank to Pending.
    Dim c As Integer
    Dim ilRet As Integer
    Dim slStatus As String
    Dim blAll As Boolean
    Dim ilCols As Integer
    
On Error GoTo ERRORBOX
    Select Case ilProcess
        Case running
            slStatus = "Running"
        Case pending
            slStatus = "Pending"
        Case paused
            slStatus = "Paused"
        Case blank
            slStatus = ""
        Case warning
            slStatus = "Warning"
    End Select
    'set indivdual line as told
    If ilTransferIndex > -1 Then
        'go through grid lines and find the one with the matching transfer index
        For c = 0 To UBound(MyExports) - 1
            If grdExportMain.TextMatrix(c + 1, MAININDEXTRANSFER) = ilTransferIndex Then
                grdExportMain.Row = c + 1
                'pause never changes.
                If grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) <> "Paused" Then
                    'don't turn to blank if was previously 'warning'
                    If ilProcess = blank Then
                        grdExportMain.Col = 0
                        If grdExportMain.CellForeColor = vbRed Then
                            slStatus = "Warning"
                        End If
                    End If
                    grdExportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slStatus
                End If
                'change font to red to mark that there was an issue
                If ilProcess = warning Then
                    For ilCols = 0 To grdExportMain.Cols - 1
                        grdExportMain.Col = ilCols
                        grdExportMain.CellForeColor = vbRed
                    Next ilCols
                End If
                ilRet = ilTransferIndex
                Exit For
            End If
        Next c
    'find the line to set
    Else
        Select Case ilProcess
            Case running
                slStatus = "Running"
                ilRet = mTestAllGridExport("Pending", slStatus, True)
                If ilRet = PROCESSNONE Then
                    ilRet = mTestAllGridExport("", slStatus, True)
                End If
                If ilRet = PROCESSNONE Then
                    ilRet = mTestAllGridExport("Warning", slStatus, True)
                End If
'                End If
                mTestAllGridExport "", "Pending", True, False
            Case pending
                slStatus = "Pending"
                ilRet = mTestAllGridExport("", slStatus, True)
            Case paused
            
            Case blank
            'change 'warning' to 'Warning' this is how I blocked it from continuing to run when connection error discovered.
            Case warning
                slStatus = "Warning"
                ilRet = mTestAllGridExport("warning", slStatus, False, False)
        End Select
    End If
    mSetStatus = ilRet
    Exit Function
ERRORBOX:
    myLog.WriteError "Error in mSetStatus:" & Err.Description, True, True
    mSetStatus = PROCESSNONE
End Function
Private Function mSetStatusImports(ilProcess As ProcessChoices, Optional ilTransferIndex As Integer = PROCESSNONE) As Integer
    'return the transfer index!
    'in all cases below, must have files!
    ' Running: change waiting to running. If not waiting, change first pending.  No pending?  First blank. After blank, warning
    ' Pending: change all blank to Pending.
    Dim c As Integer
    Dim ilRet As Integer
    Dim slStatus As String
    Dim blAll As Boolean
    Dim ilCols As Integer
    
On Error GoTo ERRORBOX
    Select Case ilProcess
        Case running
            slStatus = "Running"
        Case pending
            slStatus = "Pending"
        Case paused
            slStatus = "Paused"
        Case blank
            slStatus = ""
        Case warning
            slStatus = "Warning"
    End Select
    'set indivdual line as told
    If ilTransferIndex > -1 Then
        'go through grid lines and find the one with the matching transfer index
        For c = 0 To UBound(myImports) - 1
            If grdImportMain.TextMatrix(c + 1, MAININDEXTRANSFER) = ilTransferIndex Then
                grdImportMain.Row = c + 1
                'pause never changes.
                If grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) <> "Paused" Then
                    'don't turn to blank if was previously 'warning'
                    If ilProcess = blank Then
                        grdImportMain.Col = 0
                        If grdImportMain.CellForeColor = vbRed Then
                            slStatus = "Warning"
                        End If
                    End If
                    grdImportMain.TextMatrix(c + 1, MAININDEXSTATUS) = slStatus
                End If
                'change font to red to mark that there was an issue
                If ilProcess = warning Then
                    For ilCols = 0 To grdImportMain.Cols - 1
                        grdImportMain.Col = ilCols
                        grdImportMain.CellForeColor = vbRed
                    Next ilCols
                End If
                ilRet = ilTransferIndex
                Exit For
            End If
        Next c
    'find the line to set
    Else
        Select Case ilProcess
            Case running
                slStatus = "Running"
                ilRet = mTestAllGridImport("Pending", slStatus, True)
                If ilRet = PROCESSNONE Then
                    ilRet = mTestAllGridImport("", slStatus, True)
                End If
                If ilRet = PROCESSNONE Then
                    ilRet = mTestAllGridImport("Warning", slStatus, True)
                End If
'                End If
                mTestAllGridImport "", "Pending", True, False
            Case pending
                slStatus = "Pending"
                ilRet = mTestAllGridImport("", slStatus, True)
            Case paused
            
            Case blank
            'change 'warning' to 'Warning' this is how I blocked it from continuing to run when connection error discovered.
            Case warning
                slStatus = "Warning"
                ilRet = mTestAllGridImport("warning", slStatus, False, False)
        End Select
    End If
    mSetStatusImports = ilRet
    Exit Function
ERRORBOX:
    myLog.WriteError "Error in mSetStatusImports:" & Err.Description, True, True
    mSetStatusImports = PROCESSNONE
End Function

Private Sub mInitGrid()

On Error GoTo ERRORBOX
    grdExportFiles.Width = 4875
    grdExportFiles.ScrollBars = flexScrollBarVertical
    grdImportFiles.Width = grdExportFiles.Width
    grdImportFiles.ScrollBars = grdExportFiles.ScrollBars
    mSetGridHeight
    mSetGridColumns
    mSetGridTitles
    mClearGrid
    gGrid_IntegralHeight grdExportMain
    gGrid_FillWithRows grdExportMain
    gGrid_IntegralHeight grdExportFiles
    gGrid_FillWithRows grdExportFiles
    gGrid_IntegralHeight grdImportMain
    gGrid_FillWithRows grdImportMain
    gGrid_IntegralHeight grdImportFiles
    gGrid_FillWithRows grdImportFiles
    mSetStatusGridColor
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in mInitGrid:" & Err.Description, True, True
End Sub
Private Sub mSetGridColumns()
    With grdExportMain
        .Width = 4530
        .ColWidth(MAININDEXNAME) = .Width * 0.2
        .ColWidth(MAININDEXPENDING) = .Width * 0.2
        .ColWidth(MAININDEXPROCESSED) = .Width * 0.2
        .ColWidth(MAININDEXSTATUS) = .Width * 0.2
        .ColWidth(MAININDEXERROR) = 0
        .ColWidth(MAININDEXTRANSFER) = 0
        .Width = .ColWidth(MAININDEXSTATUS) + .ColWidth(MAININDEXNAME) + .ColWidth(MAININDEXPENDING) + .ColWidth(MAININDEXPROCESSED)
    End With
    With grdImportMain
        .Width = 4530
        .ColWidth(MAININDEXSTATUS) = .Width * 0.2
        .ColWidth(MAININDEXNAME) = .Width * 0.2
        .ColWidth(MAININDEXPENDING) = .Width * 0.2
        .ColWidth(MAININDEXPROCESSED) = .Width * 0.2
        .ColWidth(MAININDEXERROR) = 0
        .ColWidth(MAININDEXTRANSFER) = 0
        .Width = .ColWidth(MAININDEXSTATUS) + .ColWidth(MAININDEXNAME) + .ColWidth(MAININDEXPENDING) + .ColWidth(MAININDEXPROCESSED)
    End With

    With grdExportFiles
        .ColWidth(FILESINDEXFILENAME) = .Width * 0.35
        .ColWidth(FILESINDEXDATE) = .Width * 0.2
        .ColWidth(FILESINDEXTIME) = .Width * 0.17
        .ColWidth(FILESINDEXSTATUS) = .Width * 0.27
        .Width = .ColWidth(FILESINDEXFILENAME) + .ColWidth(FILESINDEXDATE) + .ColWidth(FILESINDEXTIME) + .ColWidth(FILESINDEXSTATUS)
    End With
    With grdImportFiles
        .ColWidth(FILESINDEXFILENAME) = .Width * 0.35
        .ColWidth(FILESINDEXDATE) = .Width * 0.2
        .ColWidth(FILESINDEXTIME) = .Width * 0.17
        .ColWidth(FILESINDEXSTATUS) = .Width * 0.27
        .Width = .ColWidth(FILESINDEXFILENAME) + .ColWidth(FILESINDEXDATE) + .ColWidth(FILESINDEXTIME) + .ColWidth(FILESINDEXSTATUS)
    End With

End Sub


Private Sub mSetGridTitles()
'    grdExportMain.TextMatrix(0, MAININDEXPENDING) = "Files"
'    grdExportMain.TextMatrix(0, MAININDEXNAME) = "Export"
'    grdExportMain.TextMatrix(0, MAININDEXSTATUS) = "Status"
    With grdExportMain
        .TextMatrix(0, MAININDEXPENDING) = "Pending"
        .TextMatrix(0, MAININDEXPROCESSED) = "Processed"
        .TextMatrix(0, MAININDEXNAME) = "Export"
        .TextMatrix(0, MAININDEXSTATUS) = "Status"
    End With
    With grdImportMain
        .TextMatrix(0, MAININDEXPENDING) = "Pending"
        .TextMatrix(0, MAININDEXPROCESSED) = "Processed"
        .TextMatrix(0, MAININDEXNAME) = "Import"
        .TextMatrix(0, MAININDEXSTATUS) = "Status"
    End With
    With grdExportFiles
        .TextMatrix(0, 0) = "File Name"
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Time"
        .TextMatrix(0, 3) = "Status"
    End With
    With grdImportFiles
        .TextMatrix(0, 0) = "File Name"
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Time"
        .TextMatrix(0, 3) = "Status"
    End With
End Sub

Private Sub mSetStatusGridColor()
    Dim llRow As Long
    Dim llCol As Long
    
    With grdExportMain
        For llRow = .FixedRows To .Rows - 1 Step 1
            For llCol = 0 To .Cols - 1 Step 1
                .Row = llRow
                .Col = llCol
                .CellBackColor = LIGHTYELLOW
            Next llCol
        Next llRow
    End With
    With grdImportMain
        For llRow = .FixedRows To .Rows - 1 Step 1
            For llCol = 0 To .Cols - 1 Step 1
                .Row = llRow
                .Col = llCol
                .CellBackColor = LIGHTYELLOW
            Next llCol
        Next llRow
    End With

    With grdExportFiles
        For llRow = .FixedRows To .Rows - 1 Step 1
            For llCol = 0 To 3 Step 1
                .Row = llRow
                .Col = llCol
                .CellBackColor = LIGHTYELLOW
            Next llCol
        Next llRow
    End With
    With grdImportFiles
        For llRow = .FixedRows To .Rows - 1 Step 1
            For llCol = 0 To 3 Step 1
                .Row = llRow
                .Col = llCol
                .CellBackColor = LIGHTYELLOW
            Next llCol
        Next llRow
    End With
End Sub

Private Sub mClearGrid()
    gGrid_Clear grdExportMain, True
    gGrid_Clear grdExportFiles, True
    gGrid_Clear grdImportMain, True
    gGrid_Clear grdImportFiles, True
End Sub
Private Sub mSetGridHeight()
    Dim ilCount As Integer
    'show a row for each possible transfer, or only for those defined in ini?
    ilCount = CHOICES + 2
    grdExportMain.Height = ilCount * grdExportMain.RowHeight(0)
    grdImportMain.Height = ilCount * grdImportMain.RowHeight(0)
    grdImportMain.Left = grdExportMain.Left
    grdImportMain.Top = grdExportMain.Top
    grdImportFiles.Top = grdExportFiles.Top
    grdImportFiles.Left = grdExportFiles.Left
End Sub
Private Sub mTestTelNet(ilTransferChoice As Integer)
    With MyExports(ilTransferChoice)
        If .Connect() Then
            MsgBox "Connected successfully.", vbOKOnly, "TelNet Connection"
            myLog.WriteFacts "User successfully tested connection to " & .Name
            .DisConnect
        Else
            MsgBox "Connection failed.", vbOKOnly, "TelNet Connection"
            If Len(.ErrorMessage) > 0 Then
                myLog.WriteWarning "Failed to connect to " & .Name & ":" & .ErrorMessage
            Else
                myLog.WriteWarning "Unknown error trying to connect to " & .Name
            End If
        End If
    End With
End Sub
Private Sub mCleanDataFiles()
    Dim c As Integer
    Dim ilDays As Integer
    Dim slExt As String
    Dim slPath As String
    Dim slFileName As String
    Dim myCurrent As File
    Dim slCurrentDate As String
    
    For c = 0 To UBound(MyExports) - 1
        With MyExports(c)
            ilDays = .DaysToSave
            slExt = .ExtensionToFind
            slPath = .SaveFolder
        End With
        slFileName = ""
        slFileName = Dir(slPath & "*" & slExt)
        Do While slFileName > ""
            If InStr(1, slFileName, "UNREAD", vbTextCompare) = 0 Then
                If myLog.myFile.FileExists(slPath & slFileName) Then
                    Set myCurrent = myLog.myFile.GetFile(slPath & slFileName)
                    slCurrentDate = myCurrent.DateCreated
                    If DateDiff("d", slCurrentDate, Now()) > ilDays Then
                        myLog.myFile.DeleteFile slPath & slFileName
                    End If
                End If
            End If
           slFileName = Dir()
        Loop
    Next c
Cleanup:
    Set myCurrent = Nothing
    Exit Sub
ERRORBOX:
    myLog.WriteError "mCleanDataFiles: " & Err.Description, , True
End Sub
Private Sub tmcStart_Timer()
    mRun
End Sub

'TELNET CONTROLS
Public Sub ttcControl_Connect()
On Error GoTo ERRORBOX
'dan here
  ' ttcControl.Echo True
    'ttcControl.Echo False
     ttcControl.Echo bgEcho
    Exit Sub
ERRORBOX:
    myLog.WriteError "Error in ttcControl_Connect: " & Err.Description, True, False
End Sub
Public Sub ttcControl_DataArrival()
    Dim SlData As String
    
    'probably not needed here. I set it earlier.
    sgTelNetReturn = ""
    SlData = Replace$(ttcControl.GetData(), vbCrLf, vbFormFeed)
    SlData = Replace$(SlData, LFCR, vbFormFeed)
    SlData = Replace$(SlData, vbLf, vbFormFeed)
    SlData = Replace$(SlData, vbFormFeed, vbCrLf)
    sgTelNetReturn = SlData
    If bmLogAll Then
        myLog.WriteFacts "RETURNED: " & sgTelNetReturn, False
    End If
End Sub

Public Sub ttcControl_Disconnect()
    myLog.WriteFacts "disconnected", False
End Sub

Public Sub ttcControl_Error(ByVal Number As Long, ByVal Description As String)
    sgTelNetReturn = "Error"
    myLog.WriteWarning "Error in TelNet connection:" & Description
End Sub
'END TELNETCONTROLS
