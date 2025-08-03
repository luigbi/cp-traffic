VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportWegener 
   Caption         =   "Import Wegener-Compel Spots"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6675
   Icon            =   "AffImportWegener.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6675
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   480
      Top             =   4920
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   4920
   End
   Begin VB.CommandButton cmcStationInfo 
      Caption         =   "Browse..."
      Height          =   300
      Left            =   5550
      TabIndex        =   11
      Top             =   585
      Width           =   1065
   End
   Begin VB.TextBox txtStationInfo 
      Height          =   300
      Left            =   1065
      TabIndex        =   10
      Top             =   570
      Width           =   4335
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5490
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   5145
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox edcMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   645
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1095
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   1065
      TabIndex        =   6
      Top             =   150
      Width           =   4335
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse..."
      Height          =   300
      Left            =   5550
      TabIndex        =   5
      Top             =   165
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   2205
      ItemData        =   "AffImportWegener.frx":08CA
      Left            =   150
      List            =   "AffImportWegener.frx":08CC
      TabIndex        =   1
      Top             =   2160
      Width           =   6315
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5805
      Top             =   5025
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5475
      FormDesignWidth =   6675
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   4995
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   4995
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6255
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacWFileInfo 
      Caption         =   "(Import information from: JNS_RecSerialNum.Csv; Port B-C.Csv)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1035
      TabIndex        =   13
      Top             =   885
      Width           =   7320
   End
   Begin VB.Label lacStationInfo 
      Caption         =   "Station Info"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   585
      Width           =   900
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   165
      Width           =   780
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   210
      TabIndex        =   4
      Top             =   4410
      Width           =   6285
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   1800
      Width           =   6300
   End
   Begin VB.Menu mnuGuide 
      Caption         =   "Tools"
      Begin VB.Menu mnuHaltWeb 
         Caption         =   "Halt Web Export"
      End
      Begin VB.Menu mnuFakeWeb 
         Caption         =   "Fake Web Export"
      End
   End
End
Attribute VB_Name = "frmImportWegener"
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
Option Compare Text
'
'   All return spot records are in Eastern time zone
'   All station should be setup to air the same spot at the same time
'   (i.e. 6:15am spot in Eastern will run at 6:15 Central and at 6:15 Mountain and 6:15am Pacific.  Spots are delayed and retransmitted)
'   Therefore, when a single date is returned it will have spots that aired in the previous day in Eastern zone
'              A spot transmitted at 9:20 on the Pacific channel, actually airs at 6:20am Pacific time
'              A spot transmitted at 2:40am on the Pacific channel, actually airs at 11:40pm on previous day in the Pacific time
'              The same spot would air at 11:40pm on all zones as stated above.
'   To match up with breaks the Lst must be retained by zone.
'   For Pacific zone, lst

Private smDate As String     'Import Date
Private smMODate As String
Private imImporting As Integer
Private imTerminate As Integer
Private lmMaxWidth As Long
'Private hmMsg As Integer
Private hmResult As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private cptt_rst As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Private vpf_rst As ADODB.Recordset
Private vff_rst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private gsf_rst As ADODB.Recordset
Private cif_rst As ADODB.Recordset
'Private smFields(1 To 100) As String
Private smFields(0 To 99) As String
Private tmCPDat() As DAT
Private tmAirSpotInfo() As AIRSPOTINFO
Private tmWegenerImport() As WEGENERIMPORT
Private imGameNo() As Integer
Private lmGsfCode() As Long
Private hmAst As Integer
Private tmAstInfo() As ASTINFO
Private lmAttCode() As Long
'11/5/13: change from local to module so that mBuildBreakArray can be moved
Private smETLstFrom As String
Private smCTLstFrom As String
Private smMTLstFrom As String
Private smPTLstFrom As String

Private AstInfo_rst As ADODB.Recordset
Private Type COMPELEXPORTINFO
            sVefName As String * 40
            sCallLetters As String * 40
            sGame As String * 20
            sAiredDate As String * 10
            sAiredTime As String * 8
            sAdvertiser As String * 30
            sISCI As String * 20
            sStatus As String * 10
            ilBreakNo As Integer
            lRow As Long
        End Type
Private tmCompelExportInfo() As COMPELEXPORTINFO

Private Type EXPORTASTINFO
            iMissReason As Integer
            iShttCode As Integer
            iVefCode As Integer
            lAttCode As Long
            lAstCode As Long
            sAiredDate As String * 10
            sFeedDate As String * 10
            sAiredTime As String * 8
            sISCI As String * 20
            sType As String * 1
            '11/18/14: Add Event code
            gsfCode As Long
        End Type
Private tmExportWebSpot() As EXPORTASTINFO

Private BreakInfo_rst As ADODB.Recordset
Private ISCIInfo_rst As ADODB.Recordset
Private PlayCmmdInfo_rst As ADODB.Recordset
Private ImportSpotInfo_rst As ADODB.Recordset
Private PledgeError_rst As ADODB.Recordset
Private AgreementError_rst As ADODB.Recordset
Private DayError_rst As ADODB.Recordset
Private MissingStation_rst As ADODB.Recordset
Private LSTMap_rst As ADODB.Recordset
'Web Site Info
Private lmWebTtlComments As Long
Private lmWebTtlMultiUse As Long
Private lmWebTtlHeaders As Long
Private lmWebTtlSpots As Long
Private lmWebTtlEmail As Long
Private lmTtlEventSpots As Long
Private lmWebTtlEventSpots As Long
Private imIdx As Integer
Private rstWebQ As ADODB.Recordset
Private lmTotalRecordsProcessed As Long
Private tmCsiFtpInfo As CSIFTPINFO
Private tmCsiFtpStatus As CSIFTPSTATUS
Private tmCsiFtpErrorInfo As CSIFTPERRORINFO
Private smWebImports As String
Private smCheckIniReIndex As String
Private smMinSpotsToReIndex As String
Private smMaxWaitMinutes As String
Private imFTPEvents As Boolean
Private imFtpInProgress As Boolean
Private mFtpArray() As String
Private smAttWebInterface As String

Private imWebExporting As Integer
Private smWebWorkStatus As String
Private smWebSpots As String
Private smWebToFileDetail As String
Private smWebHeader As String
Private smWebToFileHeader As String
Private smWebCopyRot As String
Private smWebToCpyRot As String
Private smWebMultiUse As String
Private smWebToMultiUse As String
Private smWebExports As String
Private smWebEventInfo As String
Private smWebToEventInfo As String
Private hmWebToDetail As Integer
Private hmWebToHeader As Integer
Private hmWebToCpyRot As Integer
Private hmWebToMultiUse As Integer
Private hmWebToEventInfo As Integer
Private bmFoundRecs As Boolean
Private lmTtlMultiUse As Long
Private lmFileMultiUseCount As Long
Private smCurDir As String
Private smCurDrive As String
Private lmMaxRecs As Long
Private lmExportWebCount As Long
Private lmMaxCmpRecs As Long
Private lmExptCmpCnt As Long
Private imExporting As Integer
Private imNeedToSend As Integer
'Dan for 7458
Private bmFTPIsOn As Boolean
'Private FTPIsOn As Integer
Private imWaiting As Integer
Private imSomeThingToDo As Integer
Const cmOneMegaByte As Long = 1000000
Const cmOneSecond As Long = 1000
Private smStatus As String
Private smMsg1 As String
Private smMsg2 As String
Private smDTStamp As String
Private smFileName As String
Private lmTotalAddSpotCount As Long
Private lmTotalDeleteSpotCount As Long
'7458
Private cmPathForgLogMsg As String  ' = "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt"
Private myEnt As CENThelper
'Dan M 4/15/15
Private bmHaltWeb As Boolean
Private bmFakeWeb As Boolean
'2/12/18: Add field that indicates which import form is being used
Dim smImportForm As String  'Search for vehicle/day/nreak on which row type
'Wegener Vars.
Dim smCompelDays2Retain As String
Dim smCompelSavePath As String
Dim smCompelImpPath As String
Dim smCompelStnInfoPath As String
Dim smCompelFileNames() As String
Dim slCompelStatus As String
Dim lmExpCmpCnt As Long
Dim slCompelBaseName As String
Dim imNewFile As Integer



Private Sub cmcBrowse_Click()
    'Wegner-Compel Import
    smCurDrive = Left$(CurDir$, 1)
    smCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    'CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv|OCSV Files (*.ocsv)|*.ocsv|OASV Files (*.oasv)|*.oasv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    CommonDialog1.fileName = ""
    CommonDialog1.DialogTitle = "Select Import File"
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDrive smCurDir
    ChDir smCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmcStationInfo_Click()
    'Import wegener-compel
    'TTP 10523 - Affiliate exports: replace Browse button with Windows Browse button
    'smCurDir = CurDir
    'igPathType = 0
    'sgGetPath = txtStationInfo.Text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    txtStationInfo.Text = sgGetPath
    'End If
    'ChDir smCurDir
    gBrowseForFolder CommonDialog1, txtStationInfo
    Exit Sub
End Sub

Private Sub cmdImport_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String
    Dim ilRet As Integer

    On Error GoTo ErrHand
    lbcMsg.Clear
    DoEvents
    Screen.MousePointer = vbHourglass
    gLogMsgWODT "O", hmResult, sgMsgDirectory & "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt"
    gLogMsgWODT "W", hmResult, "Import Started: " & Format(Now, "mm-dd-yy")
    If Not mCheckFile() Then
        Screen.MousePointer = vbDefault
        gLogMsgWODT "C", hmResult, ""
        If Not igCompelAutoImport Then
            txtFile.SetFocus
        End If
        Exit Sub
    End If
'    If Not mOpenMsgFile(sMsgFileName) Then
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
    imImporting = True
    On Error GoTo 0
    lbcMsg.Enabled = True
    lacResult.Caption = ""
    
    edcMsg.Text = "Generating Advertiser/ISCI cross reference...."
    edcMsg.Visible = True
    DoEvents
    gLogMsgWODT "W", hmResult, "  Building Adverister/ISCI List: " & " " & Format(Now, "mm-dd-yy")
    iRet = mBuildISCIInfo()
    edcMsg.Visible = False
    If Not iRet Then
        gLogMsgWODT "W", hmResult, "See WegenerImportLog for error messages: " & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "C", hmResult, ""
        imImporting = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If imTerminate Then
        gLogMsgWODT "W", hmResult, "User Terminated Import: " & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "C", hmResult, ""
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    
    edcMsg.Text = "Reading Station Info from Wegener...."
    edcMsg.Visible = True
    DoEvents
    gLogMsgWODT "W", hmResult, "  Building Station List from Compel: " & " " & Format(Now, "mm-dd-yy")
    iRet = mReadStationReceiverRecords()
    edcMsg.Visible = False
    If iRet <> 0 Then
        imImporting = False
        Screen.MousePointer = vbDefault
        If iRet = 1 Then
            gLogMsgWODT "W", hmResult, "See WegenerImportLog for error messages: " & " " & Format(Now, "mm-dd-yy")
            gLogMsgWODT "C", hmResult, ""
            Exit Sub
        ElseIf iRet = 2 Then
            gLogMsgWODT "W", hmResult, "See WegenerImportLog for error messages: " & " " & Format(Now, "mm-dd-yy")
            gLogMsgWODT "C", hmResult, ""
            Exit Sub
        Else
            If Not igCompelAutoImport Then
                iRet = gMsgBox("Some Stations Not Defined within the Affiliate system, Continue anyway", vbYesNo + vbQuestion, "Information")
                iRet = vbYes
            Else
                'gLogMsgWODT "W", hmResult, "Some Stations Not Defined within the Affiliate system" & " " & Format(Now, "mm-dd-yy")
            End If
            If iRet = vbNo Then
                gLogMsgWODT "W", hmResult, "User Terminated to View Stations Not in Affiliate: " & " " & Format(Now, "mm-dd-yy")
                gLogMsgWODT "C", hmResult, ""
                Exit Sub
            End If
        End If
        imImporting = True
        Screen.MousePointer = vbHourglass
    End If
    If imTerminate Then
        gLogMsgWODT "W", hmResult, "User Terminated Import: " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "C", hmResult, ""
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    ''''lbcMsg.Clear
    'edcMsg.Text = "Generating Station Spots...."
    'edcMsg.Visible = True
    'gLogMsgWODT "W", hmResult, "  Creating Affiliate Spots: " & Format(Now, "mm-dd-yy")
    '11/5/15: Moved createAst
    'iRet = mCreateAST()
    'edcMsg.Visible = False
    '11/5/13: Moved createAst
    'If Not iRet Then
    '    gLogMsgWODT "W", hmResult, "See WegenerImportLog for error messages: " & Format(Now, "mm-dd-yy")
    '    gLogMsgWODT "C", hmResult, ""
    '    imImporting = False
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'If imTerminate Then
    '    gLogMsgWODT "W", hmResult, "User Terminated Import: " & Format(Now, "mm-dd-yy")
    '    gLogMsgWODT "C", hmResult, ""
    '    'Print #hmMsg, "** User Terminated **"
    '    'Close #hmMsg
    '    Close #hmTo
    '    imImporting = False
    '    Screen.MousePointer = vbDefault
    '    cmdCancel.SetFocus
    '    Exit Sub
    'End If
    edcMsg.Text = "Reading Wegener Aired Spots...."
    edcMsg.Visible = True
    gLogMsgWODT "W", hmResult, "  Reading Wegener Aired Spots: " & " " & Format(Now, "mm-dd-yy")
    bgTaskBlocked = False
    sgTaskBlockedName = "Wegener Import"
    iRet = mImportSpots()
    edcMsg.Visible = False
    If (iRet = False) Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsgWODT "W", hmResult, "See WegenerImportLog for error messages: " & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "C", hmResult, ""
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsgWODT "W", hmResult, "User Terminated Import: " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "C", hmResult, ""
        Close #hmTo
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    'Clear old aet records out
    On Error GoTo ErrHand
    If bgTaskBlocked Then
        SetResults "Some spots were blocked during Import.", RGB(255, 0, 0)
        If Not igCompelAutoImport Then
            gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
        Else
            gLogMsgWODT "W", hmResult, "Import Started: " & Format(Now, "mm-dd-yy")
            gLogMsgWODT "W", hmResult, "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information." & " " & Format(Now, "mm-dd-yy")
        End If
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imImporting = False
    gLogMsgWODT "W", hmResult, "Wegener Import Completed: " & " " & Format(Now, "mm-dd-yy")
    gLogMsgWODT "W", hmResult, "** Completed Import Aired Station Spots" & " **"
    gLogMsgWODT "W", hmResult, ""
    lacResult.Caption = "Results: " & sgMsgDirectory & "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt"
    cmdImport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsgWODT "C", hmResult, ""
    '7458
    Set myEnt = Nothing
    Exit Sub
cmdImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener -cmdImport_Click"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    'debug
    'Resume Next
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    
    If imImporting Then
        imTerminate = True
    End If
    If Not igCompelAutoImport Then
        Unload frmImportWegener
    End If
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    Dim FTPIsOn As String
    Dim ilIdx As Integer
    Dim ilPos As Integer
    
'    If gFileExist(slPath) = FILEEXISTS Then
'    End If
    imNewFile = True
    Screen.MousePointer = vbHourglass
    cmPathForgLogMsg = "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt"
    frmImportWegener.Caption = "Import Wegener-Compel Aired Spots "
    imTerminate = False
    imImporting = False
    
    txtFile.Text = ""   'sgImportDirectory & "CSISpots.txt"
    If Len(sgImportDirectory) > 0 Then
        txtStationInfo.Text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
    Else
        txtStationInfo.Text = ""
    End If
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")

    Call gLoadOption(sgWebServerSection, "FTPIsOn", FTPIsOn)
    'Dan for 7458
    If FTPIsOn = "1" Then
        bmFTPIsOn = True
    Else
        bmFTPIsOn = False
    End If
    
    ReDim tmExportWebSpot(0 To 0) As EXPORTASTINFO
    ReDim tmCompelExportInfo(0 To 0) As COMPELEXPORTINFO
    Call gLoadOption(sgWebServerSection, "WebExports", smWebExports)
    smWebExports = gSetPathEndSlash(smWebExports, True)
    Call gLoadOption(sgWebServerSection, "WebImports", smWebImports)
    smWebImports = gSetPathEndSlash(smWebImports, True)
    
    ilRet = gPopVff()
    ilRet = gPopShttInfo()
    Screen.MousePointer = vbDefault
    ilRet = mInitFTP()
    '************************** Start Compel Auto Import *****************************
    If igCompelAutoImport Then
        Call gLoadOption("Wegener", "Days_to_Retain", smCompelDays2Retain)
        Call gLoadOption("Wegener", "Import", smCompelImpPath)
        Call gLoadOption("Wegener", "Save", smCompelSavePath)
        Call gLoadOption("Wegener", "StationInfo", smCompelStnInfoPath)
        sgTimeZone = Left$(gGetLocalTZName(), 1)
        smCurDrive = Left$(CurDir$, 1)
        smCurDir = CurDir
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
        gUpdateTaskMonitor 1, "CAI"
        frmImportWegener.Show
        'Find all files in the Import folder
        ilRet = mFindFile(smCompelImpPath)
        If ilRet > 0 Then
            txtStationInfo.Text = smCompelStnInfoPath
            For ilIdx = 0 To UBound(smCompelFileNames) - 1 Step 1
                txtFile.Text = smCompelImpPath & "\" & smCompelFileNames(ilIdx)
                ilPos = InStr(1, smCompelFileNames(ilIdx), ".", vbTextCompare)
                If ilPos > 0 Then
                    slCompelBaseName = Left(smCompelFileNames(ilIdx), ilPos - 1)
                End If
                'slCompelBaseName = smCompelFileNames(ilIdx)
                'Move and Rename files
                cmdImport_Click
                ilRet = mMoveFileAndRename(smCompelFileNames(ilIdx))
            Next ilIdx
        End If
        ilRet = mFindFile(smCompelSavePath)
        'Delete files in a given folder after X number of days
        ilRet = mDeleteFilesByDate(smCompelSavePath, smCompelDays2Retain)
        gUpdateTaskMonitor 2, "CAI"
        cmdCancel_Click
    End If
    
    '************************** End Compel Auto Import *******************************
    'dan M 4/15/15
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuGuide.Visible = True
    Else
        mnuGuide.Visible = False
    End If
    bmHaltWeb = False
    
End Sub

Private Sub Form_Resize()
    lacWFileInfo.FontSize = 7
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    mCloseBreakInfo
    mCloseAstInfo
    mCloseISCIInfo
    mClosePlayCmmdInfo
    mClosePledgeError
    mCloseAgreementError
    mCloseDayError
    mCloseMissingStation
    mCloseLSTMap
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    
    cptt_rst.Close
    ast_rst.Close
    att_rst.Close
    vpf_rst.Close
    vff_rst.Close
    lst_rst.Close
    gsf_rst.Close
    cif_rst.Close
    AstInfo_rst.Close
    Erase tmWegenerImport
    Erase tmAirSpotInfo
    Erase tmCPDat
    Erase imGameNo
    Erase lmGsfCode
    Erase tmAstInfo
    Erase lmAttCode
    'Erase tmImportAstInfo
    'Erase tmBreakInfo
    Erase tmExportWebSpot
    Erase tmCompelExportInfo
    Set frmImportWegener = Nothing
End Sub

Private Function mImportSpots() As Integer
    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilRecordType As Integer
    Dim slPortNo As String
    Dim slBreakNo As String
    Dim ilBreakNo As Integer
    Dim slPlaylistName As String
    Dim slAirStatus As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llAirTime As Long
    Dim slSerialNo As String
    Dim ilStationIndex As Integer
    Dim ilLoop As Integer
    Dim slExportID As String
    Dim ilVefCode As Integer
    Dim ilMP2Pos As Integer
    Dim ilSlashPos As Integer
    Dim slISCI As String
    Dim ilAdfCode As Integer
    Dim ilShttCode As Integer
    Dim slCallLetters As String
    Dim llAstCode As Long
    Dim ilAdjTime As Integer
    Dim slZone As String
    Dim llShttRet As Long
    Dim ilPos As Integer
    Dim llUpdateCount As Long
    Dim llNoPlayCmmdCount As Long
    Dim llNoAstCount As Long
    Dim llDayCount As Long
    Dim llAttCode As Long
    Dim llAtt As Long
    Dim slMoDate As String
    Dim slSuDate As String
    Dim ilSpotsAired As Integer
    Dim llRow As Long
    Dim ilPledgeStatus As Integer
    Dim llVef As Long
    Dim slVehicleName As String
    Dim blPledgeError As Boolean
    Dim blPlayError As Boolean
    Dim blCmmdError As Boolean
    Dim blAgreementError As Boolean
    Dim blDayError As Boolean
    Dim slDay As String
    Dim slImportDay As String
    Dim llDate As Long
    Dim blFindSpot As Boolean
    Dim llOnAirCount As Long
    Dim llOnAirStatus As Long
    Dim llType2Count As Long
    Dim llStartOnAir As Long
    Dim llMissingStation As Long
    Dim blMissingStation As Boolean
    Dim bmFoundRecs As Boolean
    Dim mixeduse_rst As ADODB.Recordset
    Dim ilMnfCode As Integer
    Dim slISCICode As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llTotalTime As Long
    Dim llISCIErrorCount As Long
    Dim llWebExpCount As Long
    '11/18/14: Add Event Code
    Dim slSportInfo As String
    Dim llGsfCode As Long
    Dim blCmmdFound As Boolean
    
    mImportSpots = False
    llStartTime = timeGetTime
    blPledgeError = False
    blPlayError = False
    blCmmdError = False
    blAgreementError = False
    blDayError = False
    blMissingStation = False
    mCloseImportSpotInfo
    Set ImportSpotInfo_rst = mInitImportSpotInfo()
    mClosePlayCmmdInfo
    Set PlayCmmdInfo_rst = mInitPlayCmmdInfo()
    mClosePledgeError
    Set PledgeError_rst = mInitPledgeError()
    mCloseAgreementError
    Set AgreementError_rst = mInitAgreementError()
    mCloseDayError
    Set DayError_rst = mInitDayError()
    llDate = gDateValue(smDate)
    Select Case Weekday(smDate)
        Case vbSunday
            slImportDay = "Su"
        Case vbMonday
            slImportDay = "Mo"
        Case vbTuesday
            slImportDay = "Tu"
        Case vbWednesday
            slImportDay = "We"
        Case vbThursday
            slImportDay = "Th"
        Case vbFriday
            slImportDay = "Fr"
        Case vbSaturday
            slImportDay = "Sa"
        Case Else
            slImportDay = "??"
    End Select

    slFromFile = txtFile.Text
    If fs.FILEEXISTS(slFromFile) Then
        Set tlTxtStream = fs.OpenTextFile(slFromFile, ForReading, False)
    Else
        If Not igCompelAutoImport Then
            Beep
            gMsgBox "Unable to open the Import file ", vbCritical
        Else
            gLogMsgWODT "W", hmResult, "Unable to open the Import file:" & slFromFile & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
    '2/12/18:Determine which import form
    If UCase$(right$(slFromFile, 4)) = ".CSV" Then
        smImportForm = "E"  'Search for Vehicle/day/break in End_of_File_Play row (Bypass the Command_To_Play_Playlist)
    Else
        smImportForm = "C"  'Search for Vehicle/day/break in Command_to_Play_Playlist row
    End If
    ReDim lmAttCode(0 To 0) As Long
    llUpdateCount = 0
    llISCIErrorCount = 0
    llNoPlayCmmdCount = 0
    llNoAstCount = 0
    llDayCount = 0
    llRow = 0
    llOnAirStatus = 0
    llOnAirCount = 0
    llStartOnAir = 0
    llType2Count = 0
    llMissingStation = 0
    Do While tlTxtStream.AtEndOfStream <> True
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Cancelled Import"
            mImportSpots = False
            Exit Function
        End If
        llRow = llRow + 1
        slLine = tlTxtStream.ReadLine
        slLine = UCase(Trim$(slLine))
        If Len(slLine) > 0 Then
            'Process Input
            gParseCDFields slLine, False, smFields()
            ilRecordType = -1
            slPortNo = ""
            slPlaylistName = ""
            slAirStatus = ""
            llGsfCode = 0
            slSportInfo = ""
            For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                smFields(ilLoop) = Trim$(smFields(ilLoop))
                If smFields(ilLoop) <> "" Then
                    If InStr(1, smFields(ilLoop), "COMMAND TO PLAY PLAYLIST", vbTextCompare) > 0 Then
                        ilRecordType = 1
                    ElseIf InStr(1, smFields(ilLoop), "END OF FILE PLAY", vbTextCompare) > 0 Then
                        ilRecordType = 2
                        llType2Count = llType2Count + 1
                    End If
                    If InStr(1, smFields(ilLoop), "OUTPUT_DECODER_NUMBER", vbTextCompare) > 0 Then
                        slPortNo = smFields(ilLoop + 1)
                    End If
                    If InStr(1, smFields(ilLoop), "INPUT_PLAYLIST_NAME", vbTextCompare) > 0 Then
                        slPlaylistName = smFields(ilLoop + 1)
                    End If
                    If InStr(1, smFields(ilLoop), "ON_AIR_STATUS", vbTextCompare) > 0 Then
                        slAirStatus = smFields(ilLoop + 1)
                        llOnAirStatus = llOnAirStatus + 1
                        If slAirStatus = "ON AIR" Then
                            llStartOnAir = llStartOnAir + 1
                        End If
                    End If
                    If InStr(1, smFields(ilLoop), "USER_TEXT", vbTextCompare) > 0 Then
                        ilRecordType = 2
                        llType2Count = llType2Count + 1
                        slSportInfo = smFields(ilLoop + 1)
                    End If
                    '7496
                    'ilMP2Pos = InStr(1, smFields(ilLoop), ".MP2", vbTextCompare)
                    ilMP2Pos = InStr(1, smFields(ilLoop), UCase(sgAudioExtension), vbTextCompare)
                    If ilMP2Pos > 0 Then
                        ilSlashPos = InStrRev(smFields(ilLoop), "/")
                        If ilSlashPos > 0 Then
                            slISCI = Mid(smFields(ilLoop), ilSlashPos + 1)
                        Else
                            ilSlashPos = InStrRev(smFields(ilLoop), "\")
                            If ilSlashPos > 0 Then
                                slISCI = Mid(smFields(ilLoop), ilSlashPos + 1)
                            Else
                                slISCI = smFields(ilLoop)
                            End If
                        End If
                        slISCI = Left(slISCI, Len(slISCI) - 4)
                    End If
                End If
            Next ilLoop
            ilStationIndex = -1
            If (ilRecordType = 1) Or (ilRecordType = 2) Then
                'Obtain station
                'slSerialNo = smFields(3)
                slSerialNo = smFields(2)
                ilStationIndex = mFindStationIndex(slSerialNo, slPortNo)
            End If
            If ilStationIndex <> -1 Then
                slCallLetters = Trim$(tmWegenerImport(ilStationIndex).sCallLetters)
                ilShttCode = tmWegenerImport(ilStationIndex).iShttCode
                If ilRecordType = 1 Then
                    'Obtain Vehicle
                    If Mid(slPlaylistName, Len(slPlaylistName) - 2, 2) = "BK" Then
                        slBreakNo = right(slPlaylistName, 1)
                        slDay = Mid(slPlaylistName, Len(slPlaylistName) - 4, 2)
                        slExportID = Left(slPlaylistName, Len(slPlaylistName) - 5)
                    ElseIf Mid(slPlaylistName, Len(slPlaylistName) - 3, 2) = "BK" Then
                        slBreakNo = right(slPlaylistName, 2)
                        slDay = Mid(slPlaylistName, Len(slPlaylistName) - 5, 2)
                        slExportID = Left(slPlaylistName, Len(slPlaylistName) - 6)
                    End If
                    ilVefCode = mFindVefCode(slExportID)
                    'add or update array of "command of play playlist"
                    '(Vehicle Code, Station Code, Break Number)
                    '2/12/18: Only create Command to Play Playlist record if form is C
                    If smImportForm = "C" Then
                        mBuildPlayCmmdInfo ilShttCode, ilVefCode, slBreakNo, slDay, Val(slSerialNo), slPortNo
                    End If
                ElseIf ilRecordType = 2 Then
                    If slAirStatus = "ON AIR" Then
                        llOnAirCount = llOnAirCount + 1
                        'slAirDate = smFields(1)
                        slAirDate = smFields(0)
                        'slAirTime = smFields(2)
                        slAirTime = smFields(1)
                        ilAdfCode = mFindAdfCode(slISCI)
                        '11/18/14: Add Event code
                        'If mFindPlayCmmd(ilShttCode, Val(slSerialNo), slPortNo, ilVefCode, ilBreakNo, slDay) Then
                        ''If slSportInfo = "" Then
                        If (slSportInfo = "") Or (InStr(1, slSportInfo, ":", vbBinaryCompare) <= 0) Then
                            '2/12/18: Determine vehicle/Day/Break#
                            If smImportForm = "C" Then
                                blCmmdFound = mFindPlayCmmd(ilShttCode, Val(slSerialNo), slPortNo, ilVefCode, ilBreakNo, slDay)
                            Else
                                blCmmdFound = mFindEndPlay(slPlaylistName, slExportID, ilVefCode, ilBreakNo, slDay)
                            End If
                        Else
                            blCmmdFound = mFindEventCmmd(slSportInfo, slAirDate, ilVefCode, llGsfCode, slDay)
                        End If
                        If blCmmdFound Then
                            
                            llVef = gBinarySearchVef(CLng(ilVefCode))
                            If llVef <> -1 Then
                                slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
                            Else
                                slVehicleName = "Vehicle Missing: " & ilVefCode
                            End If
                            'Map Date and Time from Eastern to station zone
                            'as all time are recorded in Eastern zone times
                            'Spot running at 6a PT was transmitted at 9a ET
                            llShttRet = gBinarySearchShtt(ilShttCode)
                            'If Not rst_Shtt.EOF Then
                            If llShttRet <> -1 Then
                                slZone = Trim$(tgShttInfo1(llShttRet).shttTimeZone)
                            Else
                                slZone = ""
                            End If
                            Select Case Left(slZone, 1)
                                Case "E"
                                    ilAdjTime = 0
                                Case "C"
                                    ilAdjTime = -1
                                Case "M"
                                    ilAdjTime = -2
                                Case "P"
                                    ilAdjTime = -3
                                Case Else
                                    ilAdjTime = 0
                            End Select
                            ilAdjTime = ilAdjTime + mGetZoneAdj(ilVefCode, slZone)
                            'Remove tenths of a second
                            ilPos = InStr(1, slAirTime, ".", vbTextCompare)
                            If ilPos > 0 Then
                                slAirTime = Left(slAirTime, ilPos - 1)
                            End If
                            slAirTime = Format(slAirTime, "h:mm:ssAM/PM")
                            llAirTime = gTimeToLong(slAirTime, False) + 3600 * ilAdjTime
                            If llAirTime < 0 Then
                                llAirTime = llAirTime + 86400
                                slAirTime = gLongToTime(llAirTime)
                                slAirDate = DateAdd("d", -1, slAirDate)
                            Else
                                slAirTime = gLongToTime(llAirTime)
                            End If
                            blFindSpot = False
                            Select Case Left(slZone, 1)
                                Case "E"
                                    blFindSpot = True
                                Case "C"
                                    If slImportDay = slDay Then
                                        blFindSpot = True
                                    End If
                                    If gDateValue(slAirDate) = llDate - 1 Then
                                        blFindSpot = True
                                    End If
                                Case "M"
                                    If slImportDay = slDay Then
                                        blFindSpot = True
                                    End If
                                    If gDateValue(slAirDate) = llDate - 1 Then
                                        blFindSpot = True
                                    End If
                                Case "P"
                                    If slImportDay = slDay Then
                                        blFindSpot = True
                                    End If
                                    If gDateValue(slAirDate) = llDate - 1 Then
                                        blFindSpot = True
                                    End If
                                Case Else
                                    blFindSpot = False
                            End Select
                            If blFindSpot Then
'11/5/13: Save Import spot information to be processed later
'                                'day test
'                                llAstCode = mFindAstCode(ilShttCode, ilVefCode, ilBreakNo, ilAdfCode, llAttCode, ilPledgeStatus)
'                                If llAstCode <> -1 Then
'                                    'Update ast
'                                    llUpdateCount = llUpdateCount + 1
'                                    SQLQuery = "UPDATE ast SET astCPStatus = 1, "
'                                    SQLQuery = SQLQuery & "astAirDate = '" & Format$(slAirDate, sgSQLDateForm) & "', "
'                                    SQLQuery = SQLQuery & "astAirTime = '" & Format$(slAirTime, sgSQLTimeForm) & "'"
'                                    SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
'                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                                        GoSub ErrHand:
'                                    End If
'                                    If tgStatusTypes(gGetAirStatus(ilPledgeStatus)).iPledged = 2 Then
'                                        If mAddPledgeError(ilShttCode, ilVefCode) Then
'                                            gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Spot Aired but Pledge set to Not Carried"
'                                        End If
'                                        If Not blPledgeError Then
'                                            SetResults "Spot Aired but Pledge set to Not Carried"
'                                            blPledgeError = True
'                                        End If
'                                    End If
'                                    mSaveAttCode llAttCode
'                                Else
'                                    'Output error message
'                                    llNoAstCount = llNoAstCount + 1
'                                    If mAttExist(ilVefCode, ilShttCode) Then
'                                        gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Unable to Find Matching Spot (Row " & llRow & ")"
'                                        If Not blPlayError Then
'                                            SetResults "Unable to Find Matching Affiliate Spot for Import Spot"
'                                            blPlayError = True
'                                        End If
'                                    Else
'                                        If mAddAgreementError(ilShttCode, ilVefCode) Then
'                                            gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Agreement missing (Row " & llRow & ")"
'                                        End If
'                                        If Not blAgreementError Then
'                                            SetResults "Agreement Missing"
'                                            blAgreementError = True
'                                        End If
'                                    End If
'                                End If
                                '11/18/14: Add Event code
                                'ImportSpotInfo_rst.AddNew Array("VefCode", "ShttCode", "AdfCode", "BreakNo", "AirDate", "AirTime", "ISCI", "Row"), Array(ilVefCode, ilShttCode, ilAdfCode, ilBreakNo, gDateValue(slAirDate), gTimeToLong(slAirTime, False), slISCI, llRow)
                                ImportSpotInfo_rst.AddNew Array("VefCode", "ShttCode", "AdfCode", "BreakNo", "GsfCode", "AirDate", "AirTime", "ISCI", "Row"), Array(ilVefCode, ilShttCode, ilAdfCode, ilBreakNo, llGsfCode, gDateValue(slAirDate), gTimeToLong(slAirTime, False), slISCI, llRow)
                            Else
                                llDayCount = llDayCount + 1
                                If mAddDayError(ilShttCode, ilVefCode) Then
                                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Spot Aired On Day not being Processed"
                                End If
                                If Not blDayError Then
                                    SetResults "Spot Aired On Day not being Processed", 0
                                    DoEvents
                                    blDayError = True
                                End If
                            End If
                        Else
                            'Output error message
                            llNoPlayCmmdCount = llNoPlayCmmdCount + 1
                            '2/12/18: Output correct error message
                            If smImportForm = "C" Then
                                gLogMsgWODT "W", hmResult, "  " & slCallLetters & ": Unable to Find Matching 'Command To Play Playlist' for line- " & slLine
                                If Not blCmmdError Then
                                    SetResults "Unable to Find Matching 'Command To Play Playlist' for Import Spot", 0
                                    DoEvents
                                    blCmmdError = True
                                End If
                            Else
                                gLogMsgWODT "W", hmResult, "  " & slCallLetters & ": Unable to parse 'Input Playlist Name' for line- " & slLine
                                If Not blCmmdError Then
                                    SetResults "Unable to parse 'Input Playlist Name' for Import Spot", 0
                                    DoEvents
                                    blCmmdError = True
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If (ilRecordType = 2) And (slAirStatus = "ON AIR") Then
                    llMissingStation = llMissingStation + 1
                    slCallLetters = mFindMissingStation(slSerialNo, slPortNo)
                    If slCallLetters <> "" Then
                        gLogMsgWODT "W", hmResult, "  " & slCallLetters & ": Spots Airing but Affiliate Not Defined in Affiliate System"
                        If Not blMissingStation Then
                            SetResults slCallLetters & ": Spots Airing but Affiliate Not Defined in Affiliate System", 0
                            DoEvents
                            blMissingStation = True
                        End If
                    End If
                End If
            End If
        End If
    Loop
    tlTxtStream.Close
    '11/5/13: Add Processing of Imported spots
    edcMsg.Visible = False
    edcMsg.Text = "Matching Wegener Aired Spots with Affiliate Spots...."
    edcMsg.Visible = True
    DoEvents
    ilRet = mProcessImportedSpots(llUpdateCount, llNoAstCount, llISCIErrorCount, blPlayError, blPledgeError, blAgreementError)
    edcMsg.Visible = False
    SetResults "Total Spots Posted = " & llUpdateCount, 0
    gLogMsgWODT "W", hmResult, "Total Spots Posted = " & llUpdateCount
    SetResults "Total Spots Posted but ISCI not Matching = " & llISCIErrorCount, 0
    gLogMsgWODT "W", hmResult, "Total Spots Posted but ISCI not Matching = " & llISCIErrorCount
    SetResults "Total Spots Not Posted because Matching Affiliate Spot not found = " & llNoAstCount, 0
    gLogMsgWODT "W", hmResult, "Total Spots Not Posted because Matching Affiliate Spot not found = " & llNoAstCount
    SetResults "Total 'End of File Play' without matching 'Command to Play Playlist' = " & llNoPlayCmmdCount, 0
    gLogMsgWODT "W", hmResult, "Total 'End of File Play' without matching 'Command to Play Playlist' = " & llNoPlayCmmdCount
    DoEvents

'    'Set any Not Aired to received as they are not imported
'    llPrevAttCode = -1
'    For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
'        If llPrevAttCode <> tmAirSpotInfo(ilLoop).lAtfCode Then
'            slSDate = tmAirSpotInfo(ilLoop).sStartDate
'            slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
'            Do
'                slSuDate = DateAdd("d", 6, slMoDate)
'                For ilStatus = 0 To UBound(tgStatusTypes) Step 1
'                    If (tgStatusTypes(ilStatus).iPledged = 2) Then
'                        SQLQuery = "UPDATE ast SET "
'                        SQLQuery = SQLQuery + "astCPStatus = " & "1"    'Received
'                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
'                        SQLQuery = SQLQuery + " AND astCPStatus = 0"
'                        SQLQuery = SQLQuery + " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
'                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
'                        cnn.BeginTrans
'                        'cnn.Execute SQLQuery, rdExecDirect
'                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                            GoSub ErrHand:
'                        End If
'                        cnn.CommitTrans
'                    End If
'                Next ilStatus
'                slMoDate = DateAdd("d", 7, slMoDate)
'            Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate))
'        End If
'        llPrevAttCode = tmAirSpotInfo(ilLoop).lAtfCode
'    Next ilLoop
'
'    'D.S. 02/25/11 Start new compliant code
'    SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attWebInterface"
'    SQLQuery = SQLQuery & " FROM shtt, cptt, att"
'    SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
'    SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
'    SQLQuery = SQLQuery & " AND attExportType = 2"
'    SQLQuery = SQLQuery & " AND cpttVefCode = " & rst!astVefCode
'    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
'    SQLQuery = SQLQuery & " Order by shttcallletters"
'    Set cptt_rst = gSQLSelectCall(SQLQuery)
'    If Not cptt_rst.EOF Then
'        'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
'        ReDim tgCPPosting(0 To 1) As CPPOSTING
'        tgCPPosting(0).lCpttCode = cptt_rst!cpttCode
'        tgCPPosting(0).iStatus = cptt_rst!cpttStatus
'        tgCPPosting(0).iPostingStatus = cptt_rst!cpttPostingStatus
'        tgCPPosting(0).lAttCode = cptt_rst!cpttatfCode
'        tgCPPosting(0).iAttTimeType = cptt_rst!attTimeType
'        tgCPPosting(0).iVefCode = rst!astVefCode
'        tgCPPosting(0).iShttCode = cptt_rst!shttCode
'        tgCPPosting(0).sZone = cptt_rst!shttTimeZone
'        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
'        tgCPPosting(0).sAstStatus = cptt_rst!cpttAstStatus
'        ilSchdCount = 0
'        ilAiredCount = 0
'        ilCompliantCount = 0
'        igTimes = 1 'By Week
'        ilAdfCode = -1
'        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True)
'        For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
'            ilAnyAstExist = True
'            gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, tmAstInfo(ilAst).iStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
'        Next ilAst
'        If ilAiredCount <> ilCompliantCount Then
'            ilAnyNotCompliant = True
'        End If
'        SQLQuery = "Update cptt Set "
'        SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
'        SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
'        SQLQuery = SQLQuery & "cpttNoCompliant = " & ilCompliantCount & " "
'        SQLQuery = SQLQuery & " Where cpttCode = " & cptt_rst!cpttCode
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'    End If
'    'D.S. 02/25/11 End new code
'
    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
    For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
        slMoDate = gAdjYear(gObtainPrevMonday(smDate))
        slSuDate = DateAdd("d", 6, slMoDate)
        'Test to see if any spots aired or were they all not aired
        ilSpotsAired = gDidAnySpotsAir(lmAttCode(llAtt), slMoDate, slSuDate)
        If ilSpotsAired Then
            'We know at least one spot aired
           ilSpotsAired = True
        Else
            'no aired spots were found
            ilSpotsAired = False
        End If

        'Check for any spots that have not aired - astCPStatus = 0 = not aired
        SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
        SQLQuery = SQLQuery + " AND astAtfCode = " & lmAttCode(llAtt)
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            'Set CPTT as complete
            SQLQuery = "UPDATE cptt SET "
            If ilSpotsAired Then
                SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete spots aired
            Else
                SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete NO spots aired
            End If
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & lmAttCode(llAtt)
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mImportSpots"
                cnn.RollbackTrans
                mImportSpots = False
                Exit Function
            End If
            cnn.CommitTrans
        Else
            'Set CPTT as Partial
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & lmAttCode(llAtt)
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mImportSpots"
                cnn.RollbackTrans
                mImportSpots = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
    Next llAtt
    gFileChgdUpdate "cptt.mkd", True
    'D.S. Export Posted Spots to the Web
    'dan M added bmhaltweb so can test without sending to web
  '  If lmExportWebCount > 0 Then
    If lmExportWebCount > 0 And Not bmHaltWeb Then
        Call mWebOpenFiles("_C")
        mWebExportSpots
    ElseIf lmExportWebCount > 0 Then
        If Not myEnt.UpdateIncompleteByFilename(EntError) Then
            gLogMsg myEnt.ErrorMessage, myEnt.ErrorLog, False
        End If
    End If
    If Not mixeduse_rst Is Nothing Then
        If (mixeduse_rst.State And adStateOpen) <> 0 Then
            mixeduse_rst.Close
        End If
        Set mixeduse_rst = Nothing
    End If
    
    llEndTime = timeGetTime
    llTotalTime = llEndTime - llStartTime
    SetResults "Total Run Time = " & gTimeString(llTotalTime / 1000, True), 0
    gLogMsgWODT "W", hmResult, "Total Run Time = " & gTimeString(llTotalTime / 1000, True)
    SetResults "*** Wegener Export Posted Spots Completed Successfully ***", 0
    gLogMsgWODT "W", hmResult, "Total Run Time = " & gTimeString(llTotalTime / 1000, True)
    DoEvents
    mImportSpots = True
    ChDir smCurDir
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mImportSpots"
'    mImportSpots = False
'    Exit Function

End Function




'*******************************************************
'*                                                     *
'*      Procedure Name:mReadStationReceiverRecords     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File to get wegener       *
'*                      stations                       *
'*                                                     *
'*******************************************************
Private Function mReadStationReceiverRecords() As Integer
    Dim ilEof As Integer
    Dim slPath As String
    Dim slLine As String
    Dim slChar As String
    Dim slWord As String
    Dim ilRet As Integer
    Dim slCallLettersA As String
    Dim slCallLettersB As String
    Dim slCallLettersC As String
    Dim slSerialNo As String
    Dim slFromFile As String
    Dim ilImport As Integer
    Dim ilPosStart As Integer
    Dim ilPosEnd As Integer
    Dim ilVehPosStart As Integer
    Dim ilTRNPosStart As Integer
    Dim slGroup As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slMainGroup As String
    Dim slPort As String
    Dim ilFound As Integer
    Dim slTrueCallLetters As String
    Dim slCallLetters As String
    Dim il4600RX As Integer
    Dim ilPortFound As Integer
    Dim ilPass As Integer
    Dim ilStationFound As Integer
    Dim ilPort As Integer
    Dim blAx_Calls As Boolean
    
    mCloseMissingStation
    Set MissingStation_rst = mInitMissingStation()
    edcMsg.Text = "Reading Station Info from JNS_RecSerialNum.Csv...."
    mReadStationReceiverRecords = 0
    ReDim tmWegenerImport(0 To 0) As WEGENERIMPORT
    On Error GoTo mReadStationReceiverRecordsErr:
    slPath = Trim$(txtStationInfo.Text)
    If right$(slPath, 1) <> "\" Then
        slPath = slPath & "\"
    End If
'    'ilRet = 0
'    slFromFile = slPath & "JNS_RecSerialNum.Csv"
'    'hmFrom = FreeFile
'    'Open slFromFile For Input Access Read As hmFrom
'    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
'    If ilRet <> 0 Then
'        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
'        mReadStationReceiverRecords = 1
'        Exit Function
'    End If
'    Do While Not EOF(hmFrom)
'        ilRet = 0
'        On Error GoTo mReadStationReceiverRecordsErr:
'        slLine = ""
'        Do While Not EOF(hmFrom)
'            slChar = Input(1, #hmFrom)
'            If slChar = sgLF Then
'                Exit Do
'            ElseIf slChar <> sgCR Then
'                slLine = slLine & slChar
'            End If
'        Loop
'        On Error GoTo 0
'        If ilRet = 62 Then
'            ilRet = 0
'            Exit Do
'        End If
'        If imTerminate Then
'            mAddMsgToList "User Cancelled Import"
'            mReadStationReceiverRecords = 2
'            Exit Function
'        End If
'        slLine = Trim$(slLine)
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            Else
'                'Process Input
'                gParseCDFields slLine, True, smFields()
'                'slCallLettersA = Trim$(smFields(1))
'                slCallLettersA = Trim$(smFields(0))
'                'slSerialNo = smFields(2)
'                slSerialNo = smFields(1)
'                slChar = Left$(slSerialNo, 1)
'                If (slChar >= "A") And (slChar <= "Z") Then
'                    slSerialNo = Mid$(slSerialNo, 2)
'                End If
'                slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersA)
'                ilRet = gBinarySearchStation(slTrueCallLetters)
'                If ilRet = -1 Then
'                    '5/12/11: Ignore Call Letters without -AM or -FM or -HD
'                    If (InStr(1, UCase(slCallLettersA), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-HD", vbBinaryCompare) > 0) Then
'                        mAddMsgToList slCallLettersA & " not defined"
'                        MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "A", slCallLettersA, False)
'                        mReadStationReceiverRecords = 3
'                    Else
'                        'Save so that call letters can be shown if airing on this factious station
'                        MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "A", slCallLettersA, False)
'                    End If
'                Else
'                    tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersA
'                    tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
'                    tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
'                    tmWegenerImport(UBound(tmWegenerImport)).sPort = "A"
'                    tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
'                    tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
'                    tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
'                    tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
'                    tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
'                    tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
'                    tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
'                    ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
'                End If
'            End If
'        End If
'        ilRet = 0
'    Loop
'    Close hmFrom
'
'    'ilRet = 0
'    On Error GoTo mReadStationReceiverRecordsErr:
'    edcMsg.Text = "Reading Station Info from PortB-C.Csv...."
'    'hmFrom = FreeFile
'    slFromFile = slPath & "Port B-C.Csv"
'    'Open slFromFile For Input Access Read As hmFrom
'    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
'    If ilRet <> 0 Then
'        'ilRet = 0
'        slFromFile = slPath & "PortB-C.Csv"
'        'Open slFromFile For Input Access Read As hmFrom
'        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
'        If ilRet <> 0 Then
'            'ilRet = 0
'            slFromFile = slPath & "Port_B-C.Csv"
'            'Open slFromFile For Input Access Read As hmFrom
'            ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
'            If ilRet <> 0 Then
'                mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
'                mReadStationReceiverRecords = 1
'                Exit Function
'            End If
'        End If
'    End If
'    Do While Not EOF(hmFrom)
'        ilRet = 0
'        On Error GoTo mReadStationReceiverRecordsErr:
'        slLine = ""
'        Do While Not EOF(hmFrom)
'            slChar = Input(1, #hmFrom)
'            If slChar = sgLF Then
'                Exit Do
'            ElseIf slChar <> sgCR Then
'                slLine = slLine & slChar
'            End If
'        Loop
'        On Error GoTo 0
'        If ilRet = 62 Then
'            ilRet = 0
'            Exit Do
'        End If
'        If imTerminate Then
'            mAddMsgToList "User Cancelled Import"
'            mReadStationReceiverRecords = 2
'            Exit Function
'        End If
'        slLine = Trim$(slLine)
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                Exit Do
'            Else
'                'Process Input
'                gParseCDFields slLine, True, smFields()
'                'slSerialNo = Trim$(smFields(1))
'                slSerialNo = Trim$(smFields(0))
'                slChar = Left$(slSerialNo, 1)
'                If (slChar >= "A") And (slChar <= "Z") Then
'                    slSerialNo = Mid$(slSerialNo, 2)
'                End If
'                'slCallLettersB = Trim$(smFields(3))
'                slCallLettersB = Trim$(smFields(2))
'                'slCallLettersC = Trim$(smFields(5))
'                slCallLettersC = Trim$(smFields(4))
'                If (slCallLettersB <> "") And (slChar <> "<") Then
'                    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersB)
'                    ilRet = gBinarySearchStation(slTrueCallLetters)
'                    If ilRet = -1 Then
'                        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
'                        If (InStr(1, UCase(slCallLettersB), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-HD", vbBinaryCompare) > 0) Then
'                            mAddMsgToList slCallLettersB & " not defined"
'                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "B", slCallLettersB, False)
'                            mReadStationReceiverRecords = 3
'                        Else
'                            'Save so that call letters can be shown if airing on this factious station
'                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "B", slCallLettersB, False)
'                        End If
'                    Else
'                        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersB
'                        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
'                        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
'                        tmWegenerImport(UBound(tmWegenerImport)).sPort = "B"
'                        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
'                        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
'                        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
'                    End If
'                End If
'                If (slCallLettersC <> "") And (slChar <> "<") Then
'                    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersC)
'                    ilRet = gBinarySearchStation(slTrueCallLetters)
'                    If ilRet = -1 Then
'                        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
'                        If (InStr(1, UCase(slCallLettersC), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-HD", vbBinaryCompare) > 0) Then
'                            mAddMsgToList slCallLettersC & " not defined"
'                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "C", slCallLettersC, False)
'                            mReadStationReceiverRecords = 3
'                        Else
'                            'Save so that call letters can be shown if airing on this factious station
'                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "C", slCallLettersC, False)
'                        End If
'                    Else
'                        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersC
'                        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
'                        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
'                        tmWegenerImport(UBound(tmWegenerImport)).sPort = "C"
'                        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
'                        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
'                        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
'                        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
'                    End If
'                End If
'            End If
'        End If
'        ilRet = 0
'    Loop
'    Close hmFrom
    
    
    edcMsg.Text = "Reading Station Info from rx_calls.Csv...."
    'ilRet = 0
    slFromFile = slPath & "rx_calls.Csv"
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet = 0 Then
        blAx_Calls = True
        Do While Not EOF(hmFrom)
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slSerialNo = Trim$(smFields(1))
                    slSerialNo = Trim$(smFields(0))
                    If slSerialNo <> "SN" Then
                        'For ilPort = 2 To 5 Step 1
                        For ilPort = 1 To 4 Step 1
                            slPort = Chr(Asc("A") + ilPort - 1)
                            slCallLetters = Trim$(smFields(ilPort))
                            If slCallLetters <> "" Then
                                slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLetters)
                                ilRet = gBinarySearchStation(slTrueCallLetters)
                                If ilRet = -1 Then
                                    '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                                    If (InStr(1, UCase(slCallLetters), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-HD", vbBinaryCompare) > 0) Then
                                        mAddMsgToList slCallLetters & " not defined", False
                                        MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, slPort, slCallLetters, False)
                                        mReadStationReceiverRecords = 3
                                    Else
                                        'Save so that call letters can be shown if airing on this factious station
                                        MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, slPort, slCallLetters, False)
                                    End If
                                Else
                                    tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLetters
                                    tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                                    tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                                    'tmWegenerImport(UBound(tmWegenerImport)).sPort = Chr(Asc("A") + ilPort - 2)
                                    tmWegenerImport(UBound(tmWegenerImport)).sPort = slPort 'Chr(Asc("A") + ilPort - 1)
                                    tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                                    tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                                    tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                                    ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                                End If
                            End If
                        Next ilPort
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    Else
        blAx_Calls = False
        edcMsg.Text = "Reading Station Info from JNS_RecSerialNum.Csv...."
        'ilRet = 0
        slFromFile = slPath & "JNS_RecSerialNum.Csv"
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
            mReadStationReceiverRecords = 1
            Exit Function
        End If
        Do While Not EOF(hmFrom)
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slCallLettersA = Trim$(smFields(1))
                    slCallLettersA = Trim$(smFields(0))
                    'slSerialNo = smFields(2)
                    slSerialNo = smFields(1)
                    slChar = Left$(slSerialNo, 1)
                    If (slChar >= "A") And (slChar <= "Z") Then
                        slSerialNo = Mid$(slSerialNo, 2)
                    End If
                    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersA)
                    ilRet = gBinarySearchStation(slTrueCallLetters)
                    If ilRet = -1 Then
                        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                        If (InStr(1, UCase(slCallLettersA), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersA), "-HD", vbBinaryCompare) > 0) Then
                            mAddMsgToList slCallLettersA & " not defined", False
                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "A", slCallLettersA, False)
                            mReadStationReceiverRecords = 3
                        Else
                            'Save so that call letters can be shown if airing on this factious station
                            MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, "A", slCallLettersA, False)
                        End If
                    Else
                        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersA
                        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                        tmWegenerImport(UBound(tmWegenerImport)).sPort = "A"
                        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
        
        'ilRet = 0
        On Error GoTo mReadStationReceiverRecordsErr:
        edcMsg.Text = "Reading Station Info from PortB-C.Csv...."
        'hmFrom = FreeFile
        slFromFile = slPath & "Port B-C.Csv"
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            'ilRet = 0
            slFromFile = slPath & "PortB-C.Csv"
            'Open slFromFile For Input Access Read As hmFrom
            ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
            If ilRet <> 0 Then
                'ilRet = 0
                slFromFile = slPath & "Port_B-C.Csv"
                'Open slFromFile For Input Access Read As hmFrom
                ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                If ilRet <> 0 Then
                    'ilRet = 0
                    slFromFile = slPath & "Port B-D.Csv"
                    'Open slFromFile For Input Access Read As hmFrom
                    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                    If ilRet <> 0 Then
                        'ilRet = 0
                        'slFromFile = slPath & "PortB-D.Csv"
                        'Open slFromFile For Input Access Read As hmFrom
                        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                        If ilRet <> 0 Then
                            'ilRet = 0
                            slFromFile = slPath & "Port_B-D.Csv"
                            'Open slFromFile For Input Access Read As hmFrom
                            ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
                            If ilRet <> 0 Then
                                mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
                                mReadStationReceiverRecords = 1
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Do While Not EOF(hmFrom)
            ilRet = 0
            On Error GoTo mReadStationReceiverRecordsErr:
            slLine = ""
            Do While Not EOF(hmFrom)
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mReadStationReceiverRecords = 2
                Exit Function
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, True, smFields()
                    'slSerialNo = Trim$(smFields(1))
                    slSerialNo = Trim$(smFields(0))
                    slChar = Left$(slSerialNo, 1)
                    If (slChar >= "A") And (slChar <= "Z") Then
                        slSerialNo = Mid$(slSerialNo, 2)
                    End If
                    'slCallLettersB = Trim$(smFields(3))
                    'slCallLettersC = Trim$(smFields(5))
                    'If slCallLettersB <> "" Then
                    '    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersB)
                    '    ilRet = gBinarySearchStation(slTrueCallLetters)
                    '    If ilRet = -1 Then
                    '        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                    '        If (InStr(1, UCase(slCallLettersB), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersB), "-HD", vbBinaryCompare) > 0) Then
                    '            mAddMsgToList slCallLettersB & " not defined"
                    '            mReadStationReceiverRecords = 3
                    '        End If
                    '    Else
                    '        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersB
                    '        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).icode
                    '        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPort = "B"
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                    '        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                    '        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    '    End If
                    'End If
                    'If slCallLettersC <> "" Then
                    '    slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLettersC)
                    '    ilRet = gBinarySearchStation(slTrueCallLetters)
                    '    If ilRet = -1 Then
                    '        '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                    '        If (InStr(1, UCase(slCallLettersC), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLettersC), "-HD", vbBinaryCompare) > 0) Then
                    '            mAddMsgToList slCallLettersC & " not defined"
                    '            mReadStationReceiverRecords = 3
                    '        End If
                    '    Else
                    '        tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLettersC
                    '        tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).icode
                    '        tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPort = "C"
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                    '        tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                    '        tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                    '        ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                    '    End If
                    'End If
                    'For ilPort = 3 To 7 Step 2
                    For ilPort = 2 To 6 Step 2
                        If ilPort = 2 Then
                            slPort = "B"
                        ElseIf ilPort = 4 Then
                            slPort = "C"
                        Else
                            slPort = "D"
                        End If
                        slCallLetters = Trim$(smFields(ilPort))
                        If slCallLetters <> "" Then
                            slTrueCallLetters = mRemoveExtraFromCallLetters(slCallLetters)
                            ilRet = gBinarySearchStation(slTrueCallLetters)
                            If ilRet = -1 Then
                                '5/12/11: Ignore Call Letters without -AM or -FM or -HD
                                If (InStr(1, UCase(slCallLetters), "-AM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-FM", vbBinaryCompare) > 0) Or (InStr(1, UCase(slCallLetters), "-HD", vbBinaryCompare) > 0) Then
                                    mAddMsgToList slCallLetters & " not defined", False
                                    MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, slPort, slCallLetters, False)
                                    mReadStationReceiverRecords = 3
                                Else
                                    'Save so that call letters can be shown if airing on this factious station
                                    MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerialNo, slPort, slCallLetters, False)
                                End If
                            Else
                                tmWegenerImport(UBound(tmWegenerImport)).sCallLetters = slCallLetters
                                tmWegenerImport(UBound(tmWegenerImport)).iShttCode = tgStationInfo(ilRet).iCode
                                tmWegenerImport(UBound(tmWegenerImport)).sSerialNo1 = slSerialNo
                                tmWegenerImport(UBound(tmWegenerImport)).sPort = slPort
                                tmWegenerImport(UBound(tmWegenerImport)).iMktCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iMSAMktCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iFormatCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iTztCode = -1
                                tmWegenerImport(UBound(tmWegenerImport)).sPostalName = ""
                                tmWegenerImport(UBound(tmWegenerImport)).lVefCodeFirst = -1
                                tmWegenerImport(UBound(tmWegenerImport)).iRecGroupFd = False
                                ReDim Preserve tmWegenerImport(0 To UBound(tmWegenerImport) + 1) As WEGENERIMPORT
                            End If
                        End If
                    Next ilPort
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
    End If
    
    Exit Function
mReadStationReceiverRecordsErr:
    ilRet = Err.Number
    Resume Next
End Function

Private Sub mAddMsgToList(slMsg As String, Optional blWegImpResult As Boolean = True)
    'Add horizontal scroll if required and add message to list box
    'The control pbcArial is used to get the approximate width of the text as the list box does not has a TextWidth command
    Dim llValue As Long
    Dim llRg As Long
    Dim llMaxWidth
    Dim llRet As Long
    Dim llRow As Long
    
    llMaxWidth = (pbcArial.TextWidth(slMsg))
    If llMaxWidth > lmMaxWidth Then
        lmMaxWidth = llMaxWidth
    End If
    If lmMaxWidth > lbcMsg.Width Then
        'Divide by 15 to convert units and add 120 for little extra room
        'Scale Mode is in Twips
        llValue = lmMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slMsg)
    If llRow < 0 Then
        SetResults slMsg, 0
        If blWegImpResult Then
            gLogMsg slMsg, "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", False
        Else
            gLogMsg slMsg, "WegenerImportStationNotDefined" & ".txt", imNewFile
            If imNewFile Then
                gLogMsg "See WegenerImportStationNotDefined.txt for Stations Not Defined", "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", False
            End If
            imNewFile = False
        End If
    End If
End Sub
Private Function mRemoveExtraFromCallLetters(slCallLetters) As String
    Dim ilPos As Integer
    Dim slTrueCallLetters As String
    Dim slTestBand As String
    
    slTrueCallLetters = Trim$(slCallLetters)
    slTestBand = "-AM"
    ilPos = 1
    Do
        ilPos = InStr(ilPos, slTrueCallLetters, slTestBand, vbTextCompare)
        If ilPos <= 0 Then
            If slTestBand = "-HD" Then
                Exit Do
            ElseIf slTestBand = "-FM" Then
                slTestBand = "-HD"
            Else
                slTestBand = "-FM"
            End If
            ilPos = 1
        Else
            If ilPos + 2 < Len(slTrueCallLetters) Then
                If Mid$(slTrueCallLetters, ilPos + 3, 1) <> "/" Then
                    slTrueCallLetters = Left$(slTrueCallLetters, ilPos + 2) & Mid(slTrueCallLetters, ilPos + 4)
                Else
                    If slTestBand = "-HD" Then
                        Exit Do
                    ElseIf slTestBand = "-FM" Then
                        slTestBand = "-HD"
                    Else
                        slTestBand = "-FM"
                    End If
                End If
                ilPos = ilPos + 1
            Else
                Exit Do
            End If
        End If
    Loop
    mRemoveExtraFromCallLetters = slTrueCallLetters
End Function

Private Function mCreateAST(ilVefCode As Integer, ilShttCode As Integer) As Integer
    Dim ilRet As Integer
    Dim slMoDate As String
    Dim llLoop As Long
    Dim ilBreakNo As Integer
    Dim llDate As Long
    Dim slSQLQuery As String
    Dim blAddSpot As Boolean
    Dim llLstCode As Long
    Dim llFeedDate As Long
    Dim llFeedTime As Long
    Dim slFeedDate As String
    Dim slFeedTime As String
    Dim ilLocalAdj As Integer
    Dim slFindZone As String
    Dim slMapZone As String
    Dim slISCI As String
    '11/18/14: Add Event Code
    Dim llGsfCode As Long
    Dim blLstFound As Boolean
     
    On Error GoTo ErrHand
    'ReDim tmImportAstInfo(0 To 100000) As IMPORTASTINFO
    mCloseAstInfo
    Set AstInfo_rst = mInitAstInfo()
    llDate = gDateValue(smDate)
    smMODate = gObtainPrevMonday(smDate)
    '11/5/13: Remove test if wegener export
    'slSQLQuery = "SELECT DISTINCT vpfVefKCode FROM vpf_Vehicle_Options WHERE vpfWegenerExport = 'Y'"
    'Set vpf_rst = gSQLSelectCall(slSQLQuery)
    'Do While Not vpf_rst.EOF
        If imTerminate Then
            mAddMsgToList "User Cancelled Import"
            mCreateAST = False
            Exit Function
        End If
        '11/5/13: Moved
        'mBuildBreakArray vpf_rst!vpfVefKCode, slETLstFrom, slCTLstFrom, slMTLstFrom, slPTLstFrom
        mBuildBreakArray ilVefCode, smETLstFrom, smCTLstFrom, smMTLstFrom, smPTLstFrom
        If imTerminate Then
            mAddMsgToList "User Cancelled Import"
            mCreateAST = False
            Exit Function
        End If
        If BreakInfo_rst.RecordCount > 0 Then
            SQLQuery = "SELECT cpttShfCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attACName, shttCode, shttTimeZone"
            SQLQuery = SQLQuery + " FROM cptt LEFT OUTER JOIN att ON cpttAtfCode = attCode LEFT OUTER JOIN shtt On cpttShfCode = shttCode "
            'SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & vpf_rst!vpfVefKCode
            SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND cpttShfCode = " & ilShttCode
            'SQLQuery = SQLQuery + " AND shttUsedForWegener = 'Y'"
            SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(gObtainPrevMonday(smMODate), sgSQLDateForm) & "')"
            Set cptt_rst = gSQLSelectCall(SQLQuery)
            Do While Not cptt_rst.EOF
                If imTerminate Then
                    mAddMsgToList "User Cancelled Import"
                    mCreateAST = False
                    Exit Function
                End If
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cptt_rst!cpttCode
                tgCPPosting(0).iStatus = cptt_rst!cpttStatus
                tgCPPosting(0).iPostingStatus = cptt_rst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cptt_rst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cptt_rst!attTimeType
                tgCPPosting(0).iVefCode = cptt_rst!cpttvefcode
                tgCPPosting(0).iShttCode = cptt_rst!shttCode
                tgCPPosting(0).sZone = cptt_rst!shttTimeZone
                'tgCPPosting(0).sDate = Format$(smMODate, sgShowDateForm)
                tgCPPosting(0).sDate = Format$(smDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cptt_rst!cpttAstStatus
                '11/5/13
                'igTimes = 1 'By week
                igTimes = 3 'not By Week
                If Left(tgCPPosting(0).sZone, 1) <> "E" Then
                    tgCPPosting(0).iNumberDays = 0
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True, True, -1, True)
                Else
                    tgCPPosting(0).iNumberDays = 1
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True, True)
                End If
                'Might need to contvert to Eastern so that the spot match in import
                'Build array of station Code; Break Number; ast Code
                'Undo time zone conversion so that the correct LST can be found
                ilLocalAdj = mGetZoneAdj(cptt_rst!cpttvefcode, cptt_rst!shttTimeZone)
                If ilLocalAdj <> 0 Then
                    ilLocalAdj = -ilLocalAdj
                End If
                For llLoop = 0 To UBound(tmAstInfo) - 1 Step 1
                    If imTerminate Then
                        mAddMsgToList "User Cancelled Import"
                        mCreateAST = False
                        Exit Function
                    End If
                    'LST build without blackout or region copy.
                    'Get generic lst if blackout defined
                    llLstCode = tmAstInfo(llLoop).lLstCode
                    If tmAstInfo(llLoop).iRegionType = 2 Then
                        SQLQuery = "SELECT lstBkoutLstCode"
                        SQLQuery = SQLQuery & " FROM lst"
                        SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tmAstInfo(llLoop).lLstCode)
                        Set lst_rst = gSQLSelectCall(SQLQuery)
                        If Not lst_rst.EOF Then
                            llLstCode = lst_rst!lstBkoutLstCode
                        End If
                    End If
                    'Can't match on lst because some stations might have blackouts, so Break number must be used
                    'Add Zone test to retain only those spots mapped back to Eastern time Eastern zone as BreakNo by LST is
                    'from eastern zone
                    'AST generated without region copy, therefore LST should match
                    'Eastern: all
                    'Central llDate = 12am->11p; llDate-1=11p->12am
                    'Mountain llDate = 12am->10p; llDate-1=10p->12am
                    'Pacific llDate = 12am->9p; llDate-1=9p->12am
                    blAddSpot = False
                    llFeedDate = DateValue(gAdjYear(tmAstInfo(llLoop).sFeedDate))
                    llFeedTime = gTimeToLong(tmAstInfo(llLoop).sFeedTime, False)
                    
                    llFeedTime = llFeedTime + 3600 * ilLocalAdj
                    If llFeedTime < 0 Then
                        llFeedTime = llFeedTime + 86400
                        llFeedDate = llFeedDate - 1
                    ElseIf llFeedTime > 86400 Then
                        llFeedTime = llFeedTime - 86400
                        llFeedDate = llFeedDate + 1
                    End If
                    slFeedTime = Format$(gLongToTime(llFeedTime), "h:mm:ssAM/PM")
                    slFeedDate = Format$(llFeedDate, "m/d/yyyy")
                    'If (gDateValue(slFeedDate) = llDate) Then
                    '    blAddSpot = True
                    'End If
                    slFindZone = Left(cptt_rst!shttTimeZone, 1)
                    Select Case slFindZone
                        Case "E"
                            slMapZone = smETLstFrom
                            If llFeedDate = llDate Then
                                blAddSpot = True
                            End If
                        Case "C"
                            slMapZone = smCTLstFrom
                            If (llFeedDate = llDate) And (llFeedTime < gTimeToLong("11PM", False)) Then
                                blAddSpot = True
                            End If
                            If (llFeedDate = llDate - 1) And (llFeedTime >= gTimeToLong("11PM", False)) Then
                                blAddSpot = True
                            End If
                        Case "M"
                            slMapZone = smMTLstFrom
                            If (llFeedDate = llDate) And (llFeedTime < gTimeToLong("10PM", False)) Then
                                blAddSpot = True
                            End If
                            If (llFeedDate = llDate - 1) And (llFeedTime >= gTimeToLong("10PM", False)) Then
                                blAddSpot = True
                            End If
                        Case "P"
                            slMapZone = smPTLstFrom
                            If (llFeedDate = llDate) And (llFeedTime < gTimeToLong("9PM", False)) Then
                                blAddSpot = True
                            End If
                            If (llFeedDate = llDate - 1) And (llFeedTime >= gTimeToLong("9PM", False)) Then
                                blAddSpot = True
                            End If
                    End Select
                    If blAddSpot Then
                        '11/18/14: Add Event code
                        ''ilBreakNo = mFindBreakNo(llLstCode, slMapZone)
                        'ilBreakNo = mFindBreakNo(llLstCode, slFindZone)
                        blLstFound = mFindBreakNo(llLstCode, slFindZone, ilBreakNo, llGsfCode)
                        '11/18/14: Add Event code
                        'If ilBreakNo > 0 Then
                        If blLstFound Then
                            If tmAstInfo(llLoop).iRegionType > 0 Then
                                slISCI = Trim$(tmAstInfo(llLoop).sRISCI) 'sRISCI
                            Else
                                slISCI = Trim$(tmAstInfo(llLoop).sISCI)
                            End If
                            '11/18/14: Add Event Code
                            'AstInfo_rst.AddNew Array("ShttCode", "VefCode", "BreakNo", "adfCode", "AstCode", "AttCode", "PledgeStatus", "ISCI", "Processed"), Array(tmAstInfo(llLoop).iShttCode, tmAstInfo(llLoop).iVefCode, ilBreakNo, tmAstInfo(llLoop).iAdfCode, tmAstInfo(llLoop).lCode, tmAstInfo(llLoop).lAttCode, tmAstInfo(llLoop).iPledgeStatus, slISCI, False)
                            '7458 added feed date
'                            AstInfo_rst.AddNew Array("ShttCode", "VefCode", "BreakNo", "gsfCode", "adfCode", "AstCode", "AttCode", "PledgeStatus", "ISCI", "Processed"), Array(tmAstInfo(llLoop).iShttCode, tmAstInfo(llLoop).iVefCode, ilBreakNo, llGsfCode, tmAstInfo(llLoop).iAdfCode, tmAstInfo(llLoop).lCode, tmAstInfo(llLoop).lAttCode, tmAstInfo(llLoop).iPledgeStatus, slISCI, False)
                            AstInfo_rst.AddNew Array("ShttCode", "VefCode", "BreakNo", "gsfCode", "adfCode", "AstCode", "AttCode", "PledgeStatus", "ISCI", "Processed", "FeedDate"), Array(tmAstInfo(llLoop).iShttCode, tmAstInfo(llLoop).iVefCode, ilBreakNo, llGsfCode, tmAstInfo(llLoop).iAdfCode, tmAstInfo(llLoop).lCode, tmAstInfo(llLoop).lAttCode, tmAstInfo(llLoop).iPledgeStatus, slISCI, False, slFeedDate)
                            If Not myEnt.Add(slFeedDate, llGsfCode, Asts) Then
                                gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
                            End If
                        Else
                            'Error message
                        End If
                    End If
                Next llLoop
                cptt_rst.MoveNext
            Loop
        End If
    '    vpf_rst.MoveNext
    'Loop
    mCreateAST = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mCreateAst"
    Resume Next
End Function

Private Sub mBuildBreakArray(ilVefCode As Integer, slETLstFrom As String, slCTLstFrom As String, slMTLstFrom As String, slPTLstFrom As String)
    Dim llDate As Long
    Dim llStartLstDate As Long
    Dim llEndLstDate As Long
    Dim llODate As Long
    Dim llBreakNo As Long
    Dim ilPositionNo As Integer
    Dim llLogTime As Long
    Dim llLstLogDate As Long
    Dim llLstLogTime As Long
    Dim ilGsf As Integer
    Dim slZone As String
    Dim ilVefZone As Integer
    Dim ilZone As Integer
    Dim llLocalAdj As Long
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
    Dim llVefIndex As Integer
    Dim blIncludeLst As Boolean
    Dim slLstZone As String
    Dim slMapZone As String
    Dim llLstCode As Long

    On Error GoTo ErrHand
    'ReDim tmBreakInfo(0 To 0) As BREAKINFO
    mCloseBreakInfo
    Set BreakInfo_rst = mInitBreakInfo()
    slZone = "E"
    slETLstFrom = "E"
    slCTLstFrom = "E"
    slMTLstFrom = "E"
    slPTLstFrom = "E"
    llLocalAdj = 0
    llDate = gDateValue(gAdjYear(smDate))
    llStartLstDate = llDate
    llEndLstDate = llDate
    ilVefZone = gBinarySearchVef(CLng(ilVefCode))
    If ilVefZone <> -1 Then
        For ilZone = LBound(tgVehicleInfo(ilVefZone).sZone) To UBound(tgVehicleInfo(ilVefZone).sZone) Step 1
            If UCase$(Left$(Trim$(tgVehicleInfo(ilVefZone).sZone(ilZone)), 1)) = "E" Then
                slZone = UCase$(Left$(tgVehicleInfo(ilVefZone).sZone(tgVehicleInfo(ilVefZone).iBaseZone(ilZone)), 1))
                If (tgVehicleInfo(ilVefZone).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVefZone).sFed(ilZone)) <> "") Then
                    llLocalAdj = CLng(3600) * tgVehicleInfo(ilVefZone).iLocalAdj(ilZone)
                    If llLocalAdj > 0 Then
                        llStartLstDate = llStartLstDate - 1
                    Else
                        llEndLstDate = llEndLstDate + 1
                    End If
                End If
                Exit For
            End If
        Next ilZone
    
        For ilZone = LBound(tgVehicleInfo(ilVefZone).sZone) To UBound(tgVehicleInfo(ilVefZone).sZone) Step 1
            If (Trim$(tgVehicleInfo(ilVefZone).sFed(ilZone)) <> "") And (tgVehicleInfo(ilVefZone).iBaseZone(ilZone) <> -1) Then
                If (tgVehicleInfo(ilVefZone).sFed(ilZone) <> "*") Then
                    Select Case Left(tgVehicleInfo(ilVefZone).sZone(ilZone), 1)
                        Case "C"
                            slCTLstFrom = tgVehicleInfo(ilVefZone).sZone(tgVehicleInfo(ilVefZone).iBaseZone(ilZone))
                        Case "M"
                            slMTLstFrom = tgVehicleInfo(ilVefZone).sZone(tgVehicleInfo(ilVefZone).iBaseZone(ilZone))
                        Case "P"
                            slPTLstFrom = tgVehicleInfo(ilVefZone).sZone(tgVehicleInfo(ilVefZone).iBaseZone(ilZone))
                    End Select
                Else
                    Select Case Left(tgVehicleInfo(ilVefZone).sZone(ilZone), 1)
                        Case "C"
                            slCTLstFrom = "C"
                        Case "M"
                            slMTLstFrom = "M"
                        Case "P"
                            slPTLstFrom = "P"
                    End Select
                End If
            End If
        Next ilZone
    
    End If
    'Back date up to handle different zones
    If llStartLstDate = llDate Then
        llStartLstDate = llStartLstDate - 1
    End If
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
    SQLQuery = SQLQuery + " WHERE (vffVefCode = " & ilVefCode & ")"
    Set vff_rst = gSQLSelectCall(SQLQuery)
    If Not vff_rst.EOF Then
        ReDim imGameNo(0 To 1) As Integer
        ReDim lmGsfCode(0 To 1) As Long
        imGameNo(0) = 0
        lmGsfCode(0) = 0
        llVefIndex = gBinarySearchVef(CLng(ilVefCode))
        If llVefIndex <> -1 Then
            If tgVehicleInfo(llVefIndex).sVehType = "G" Then
                ReDim imGameNo(0 To 0) As Integer
                ReDim lmGsfCode(0 To 0) As Long
                SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & ilVefCode & " AND gsfAirDate = '" & Format$(smDate, sgSQLDateForm) & "'" & ")"
                Set gsf_rst = gSQLSelectCall(SQLQuery)
                Do While Not gsf_rst.EOF
                    imGameNo(UBound(imGameNo)) = gsf_rst!gsfGameNo
                    lmGsfCode(UBound(lmGsfCode)) = gsf_rst!gsfCode
                    ReDim Preserve imGameNo(0 To UBound(imGameNo) + 1) As Integer
                    ReDim Preserve lmGsfCode(0 To UBound(lmGsfCode) + 1) As Long
                    gsf_rst.MoveNext
                Loop
                gsf_rst.Close
            End If
        End If
        For ilGsf = 0 To UBound(lmGsfCode) - 1 Step 1
            llODate = -1
            llLogTime = -1
            SQLQuery = "SELECT * FROM lst "
            SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & ilVefCode
            If lmGsfCode(ilGsf) > 0 Then
                SQLQuery = SQLQuery + " AND lstGsfCode = " & lmGsfCode(ilGsf)
            End If
            SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
            '3/9/16: Fix the filter
            'SQLQuery = SQLQuery + " AND lstStatus < 20" 'Bypass MG/Bonus
            SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
            SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(llStartLstDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llEndLstDate, sgSQLDateForm) & "')" & ")"
            SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
            Set lst_rst = gSQLSelectCall(SQLQuery)
            If Not lst_rst.EOF Then
                'Build LSTMap
                mCloseLSTMap
                Set LSTMap_rst = mInitLSTMap()
                Do While Not lst_rst.EOF
                    '11/18/14: Add Event code
                    'LSTMap_rst.AddNew Array("LstCode", "Zone", "LogDate", "LogTime", "BreakNo", "Position", "GsfCode"), Array(lst_rst!lstCode, Left$(UCase(Trim$(lst_rst!lstZone)), 1), gDateValue(lst_rst!lstLogDate), gTimeToLong(lst_rst!lstLogTime, False), lst_rst!lstBreakNo, lst_rst!lstPositionNo)
                    LSTMap_rst.AddNew Array("LstCode", "Zone", "LogDate", "LogTime", "BreakNo", "Position", "GsfCode"), Array(lst_rst!lstCode, Left$(UCase(Trim$(lst_rst!lstZone)), 1), gDateValue(lst_rst!lstLogDate), gTimeToLong(lst_rst!lstLogTime, False), lst_rst!lstBreakNo, lst_rst!lstPositionNo, lmGsfCode(ilGsf))
                    lst_rst.MoveNext
                Loop
                
                lst_rst.MoveFirst
                Do While Not lst_rst.EOF
                    blSpotOk = True
                    ilAnf = gBinarySearchAnf(lst_rst!lstAnfCode)
                    If ilAnf <> -1 Then
                        If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                            blSpotOk = False
                        End If
                    End If
                    If (blSpotOk) And ((Left$(UCase(Trim$(lst_rst!lstZone)), 1) = slZone) Or (Trim$(lst_rst!lstZone) = "")) Then
                        llLstLogDate = gDateValue(gAdjYear(Format$(lst_rst!lstLogDate, sgShowDateForm)))
                        llLstLogTime = gTimeToLong(Format$(lst_rst!lstLogTime, sgShowTimeWSecForm), False)
                        llLstLogTime = llLstLogTime + llLocalAdj
                        If llLstLogTime < 0 Then
                            llLstLogTime = llLstLogTime + 86400
                            llLstLogDate = llLstLogDate - 1
                        ElseIf llLstLogTime > 86400 Then
                            llLstLogTime = llLstLogTime - 86400
                            llLstLogDate = llLstLogDate + 1
                        End If
                        'If (llLstLogDate = llDate) Then
                            If llODate <> llLstLogDate Then
                                llODate = llLstLogDate  'llDate
                                llBreakNo = 0
                                ilPositionNo = 0
                                llLogTime = -1
                            End If
                            If llLogTime <> llLstLogTime Then
                                llLogTime = llLstLogTime
                                llBreakNo = llBreakNo + 1
                                ilPositionNo = 0
                            End If
                            If lst_rst!lstsplitnetwork = "P" Then
                                ilPositionNo = ilPositionNo + 1
                            Else
                                ilPositionNo = ilPositionNo + 1
                            End If
                            'tmBreakInfo(UBound(tmBreakInfo)).lLstCode = lst_rst!lstCode
                            'tmBreakInfo(UBound(tmBreakInfo)).iBreak = llBreakNo
                            'tmBreakInfo(UBound(tmBreakInfo)).iPosition = ilPositionNo
                            'ReDim Preserve tmBreakInfo(0 To UBound(tmBreakInfo) + 1) As BREAKINFO
                            For ilZone = 0 To 3 Step 1
                                blIncludeLst = False
                                Select Case ilZone
                                    Case 0  'Eastern
                                        slMapZone = slETLstFrom
                                        slLstZone = "E"
                                        If (llLstLogDate = llDate) Then
                                            blIncludeLst = True
                                        End If
                                    Case 1  'Central
                                        slMapZone = slCTLstFrom
                                        slLstZone = "C"
                                        If (llLstLogDate = llDate) And (llLstLogTime < gTimeToLong("11PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                        If (llLstLogDate = llDate - 1) And (llLstLogTime >= gTimeToLong("11PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                    Case 2  'Moutain
                                        slMapZone = slMTLstFrom
                                        slLstZone = "M"
                                        If (llLstLogDate = llDate) And (llLstLogTime < gTimeToLong("10PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                        If (llLstLogDate = llDate - 1) And (llLstLogTime >= gTimeToLong("10PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                    Case 3  'Pacific
                                        slMapZone = slPTLstFrom
                                        slLstZone = "P"
                                        If (llLstLogDate = llDate) And (llLstLogTime < gTimeToLong("9PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                        If (llLstLogDate = llDate - 1) And (llLstLogTime >= gTimeToLong("9PM", False)) Then
                                            blIncludeLst = True
                                        End If
                                End Select
                                If blIncludeLst Then
                                    llLstCode = mFindLSTMap(lst_rst!lstCode, slMapZone)
                                    If llLstCode <> -1 Then
                                        '11/18/14: Add Event Code
                                        'BreakInfo_rst.AddNew Array("LstCode", "Zone", "BreakNo", "Position", "GsfCode"), Array(llLstCode, slLstZone, llBreakNo, ilPositionNo)
                                        BreakInfo_rst.AddNew Array("LstCode", "Zone", "BreakNo", "Position", "GsfCode"), Array(llLstCode, slLstZone, llBreakNo, ilPositionNo, lmGsfCode(ilGsf))
                                    End If
                                End If
                            Next ilZone
                        'End If
                    End If
                    lst_rst.MoveNext
                Loop
            End If
        Next ilGsf
    End If
    '11/18/14: Add Event Code
    'BreakInfo_rst.Sort = "LstCode,Zone,BreakNo"
    BreakInfo_rst.Sort = "LstCode,Zone,GsfCode,BreakNo"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mBuildBreakArray"
    Resume Next
End Sub

Private Function mInitBreakInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "LstCode", adInteger
        .Append "Zone", adChar, 1
        .Append "BreakNo", adInteger
        .Append "Position", adInteger
        .Append "GsfCode", adInteger
    End With
    rst.Open
    rst!lstCode.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitBreakInfo = rst
End Function

Private Sub mCloseBreakInfo()
    On Error Resume Next
    If Not BreakInfo_rst Is Nothing Then
        If (BreakInfo_rst.State And adStateOpen) <> 0 Then
            BreakInfo_rst.Close
        End If
        Set BreakInfo_rst = Nothing
    End If

End Sub

Private Function mFindBreakNo(llLstCode As Long, slZone As String, ilBreakNo As Integer, llGsfCode As Long) As Integer
    On Error GoTo ErrHandle
    mFindBreakNo = False
    BreakInfo_rst.Filter = "LstCode = " & llLstCode & " And Zone = '" & UCase(Left(slZone, 1)) & "'"
    If Not BreakInfo_rst.EOF Then
        ilBreakNo = BreakInfo_rst!breakno
        llGsfCode = BreakInfo_rst!gsfCode
        mFindBreakNo = True
    End If
    Exit Function
ErrHandle:
End Function

Private Function mInitAstInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ShttCode", adInteger
        .Append "VefCode", adInteger
        .Append "BreakNo", adInteger
        '11/18/14: Add Event Code
        .Append "gsfCode", adInteger
        .Append "AdfCode", adInteger
        .Append "AstCode", adInteger
        .Append "AttCode", adInteger
        .Append "PledgeStatus", adInteger
        .Append "ISCI", adChar, 20
        .Append "Processed", adBoolean
        '7458
        .Append "FeedDate", adChar, 10
    End With
    rst.Open
    rst!shttCode.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitAstInfo = rst
End Function

Private Sub mCloseAstInfo()
    On Error Resume Next
    If Not AstInfo_rst Is Nothing Then
        If (AstInfo_rst.State And adStateOpen) <> 0 Then
            AstInfo_rst.Close
        End If
        Set AstInfo_rst = Nothing
    End If

End Sub
'11/18/14: Add Event code
'Private Function mFindAstCode(ilShttCode As Integer, ilVefCode As Integer, ilBreakNo As Integer, ilAdfCode As Integer, llAttCode As Long, ilPledgeStatus As Integer, slISCI As String) As Long
Private Function mFindAstCode(ilShttCode As Integer, ilVefCode As Integer, ilBreakNo As Integer, llGsfCode As Long, slSportISCI As String, ilAdfCode As Integer, llAttCode As Long, ilPledgeStatus As Integer, slISCI As String, slFeedDate As String) As Long
    mFindAstCode = -1
    slFeedDate = ""
    '11/18/14: Add Event code
    'AstInfo_rst.Filter = "shttCode = " & ilShttCode & " And vefCode = " & ilVefCode & " And BreakNo = " & ilBreakNo & " And AdfCode = " & ilAdfCode
    If llGsfCode <= 0 Then
        AstInfo_rst.Filter = "shttCode = " & ilShttCode & " And vefCode = " & ilVefCode & " And BreakNo = " & ilBreakNo & " And AdfCode = " & ilAdfCode
        Do While Not AstInfo_rst.EOF
            If Not AstInfo_rst!Processed Then
                
                llAttCode = AstInfo_rst!attCode
                ilPledgeStatus = AstInfo_rst!PledgeStatus
                slISCI = AstInfo_rst!ISCI
                '7458
                slFeedDate = AstInfo_rst!FeedDate
                mFindAstCode = AstInfo_rst!astCode
                AstInfo_rst!Processed = True
                Exit Function
            End If
            AstInfo_rst.MoveNext
        Loop
    Else
        'Treat the ast spots as a pool of spots.
        'Serach for Matching ISCI first
        AstInfo_rst.Filter = "shttCode = " & ilShttCode & " And vefCode = " & ilVefCode & " And gsfCode = " & llGsfCode & " And AdfCode = " & ilAdfCode & " And ISCI = '" & slSportISCI & "'"
        Do While Not AstInfo_rst.EOF
            If Not AstInfo_rst!Processed Then
                llAttCode = AstInfo_rst!attCode
                ilPledgeStatus = AstInfo_rst!PledgeStatus
                slISCI = AstInfo_rst!ISCI
                 '7458
                slFeedDate = AstInfo_rst!FeedDate
               mFindAstCode = AstInfo_rst!astCode
                AstInfo_rst!Processed = True
                Exit Function
            End If
            AstInfo_rst.MoveNext
        Loop
        AstInfo_rst.Filter = "shttCode = " & ilShttCode & " And vefCode = " & ilVefCode & " And gsfCode = " & llGsfCode & " And AdfCode = " & ilAdfCode
        Do While Not AstInfo_rst.EOF
            If Not AstInfo_rst!Processed Then
                llAttCode = AstInfo_rst!attCode
                ilPledgeStatus = AstInfo_rst!PledgeStatus
                slISCI = AstInfo_rst!ISCI
                '7458
                slFeedDate = AstInfo_rst!FeedDate
                mFindAstCode = AstInfo_rst!astCode
                AstInfo_rst!Processed = True
                Exit Function
            End If
            AstInfo_rst.MoveNext
        Loop
    End If
End Function

Private Function mFindStationIndex(slSerialNo As String, slPortNo As String) As Integer
    Dim llSerialNo As Long
    Dim slPort As String
    Dim ilLoop As Integer
    
    mFindStationIndex = -1
    llSerialNo = Val(slSerialNo)
    For ilLoop = 0 To UBound(tmWegenerImport) - 1 Step 1
        If llSerialNo = Val(tmWegenerImport(ilLoop).sSerialNo1) Then
            slPort = ""
            Select Case slPortNo
                Case "1"
                    slPort = "A"
                Case "2"
                    slPort = "B"
                Case "3"
                    slPort = "C"
                Case "4"
                    slPort = "D"
            End Select
            If tmWegenerImport(ilLoop).sPort = slPort Then
                mFindStationIndex = ilLoop
                Exit Function
            End If
        End If
    Next ilLoop
End Function

Private Function mFindVefCode(slExportID As String) As Integer
    Dim ilVff As Integer
    mFindVefCode = -1
    For ilVff = LBound(tgVffInfo) To UBound(tgVffInfo) - 1 Step 1
        If UCase(Trim$(tgVffInfo(ilVff).sWegenerExportID)) = slExportID Then
            mFindVefCode = tgVffInfo(ilVff).iVefCode
            Exit Function
        End If
    Next ilVff

End Function

Private Function mInitISCIInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "AdfCode", adInteger
        .Append "ISCI", adChar, 20
    End With
    rst.Open
    rst!ISCI.Properties("optimize") = True
    Set mInitISCIInfo = rst
End Function

Private Sub mCloseISCIInfo()
    On Error Resume Next
    If Not ISCIInfo_rst Is Nothing Then
        If (ISCIInfo_rst.State And adStateOpen) <> 0 Then
            ISCIInfo_rst.Close
        End If
        Set ISCIInfo_rst = Nothing
    End If

End Sub

Private Function mBuildISCIInfo() As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand
    mBuildISCIInfo = False
    mCloseISCIInfo
    Set ISCIInfo_rst = mInitISCIInfo()
    SQLQuery = "SELECT Distinct cifAdfCode, cpfISCI "
    SQLQuery = SQLQuery + " FROM Cif_Copy_Inventory LEFT OUTER JOIN Cpf_Copy_Prodct_ISCI On cifCpfCode = cpfCode"
    SQLQuery = SQLQuery + " WHERE (cifPurged <> " & "'H'"
    SQLQuery = SQLQuery + " AND cifAdfCode <> 0 )"
    Set cif_rst = gSQLSelectCall(SQLQuery)
    Do While Not cif_rst.EOF
        If imTerminate Then
            Exit Function
        End If
        If Not IsNull(cif_rst!cpfISCI) Then
            ISCIInfo_rst.AddNew Array("AdfCode", "ISCI"), Array(cif_rst!cifAdfCode, cif_rst!cpfISCI)
        End If
        cif_rst.MoveNext
    Loop
    mBuildISCIInfo = True
    Exit Function
ErrHand:
    ChDir smCurDir
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mBuildISCIInfo"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
End Function

Private Function mFindAdfCode(slISCI As String) As Integer
    Dim slISCI20 As String
    mFindAdfCode = -1
    slISCI20 = slISCI
    Do While Len(slISCI20) < 20
        slISCI20 = slISCI20 & " "
    Loop
    ISCIInfo_rst.Filter = "ISCI=" & "'" & slISCI20 & "'"
    If Not ISCIInfo_rst.EOF Then
        mFindAdfCode = ISCIInfo_rst!adfCode
    End If
End Function

Private Function mInitPlayCmmdInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ShttCode", adInteger
        .Append "VefCode", adInteger
        .Append "BreakNo", adInteger
        .Append "Day", adChar, 2
        .Append "ID", adInteger
        .Append "Port", adChar, 1
    End With
    rst.Open
    rst!shttCode.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitPlayCmmdInfo = rst
End Function

Private Sub mClosePlayCmmdInfo()
    On Error Resume Next
    If Not PlayCmmdInfo_rst Is Nothing Then
        If (PlayCmmdInfo_rst.State And adStateOpen) <> 0 Then
            PlayCmmdInfo_rst.Close
        End If
        Set PlayCmmdInfo_rst = Nothing
    End If

End Sub

Private Function mFindPlayCmmd(ilShttCode As Integer, llID As Long, slPort As String, ilOutVefCode As Integer, ilOutBreakNo As Integer, slOutDay As String) As Integer
    mFindPlayCmmd = False
    PlayCmmdInfo_rst.Filter = "ShttCode = " & ilShttCode & " And ID = " & llID & " And Port = " & slPort    ' & ilVefCode
    If Not PlayCmmdInfo_rst.EOF Then
        ilOutVefCode = PlayCmmdInfo_rst!vefCode
        ilOutBreakNo = PlayCmmdInfo_rst!breakno
        slOutDay = PlayCmmdInfo_rst!Day
        mFindPlayCmmd = True
    End If
End Function

Private Sub mBuildPlayCmmdInfo(ilShttCode As Integer, ilVefCode As Integer, slBreakNo As String, slDay As String, llID As Long, slPort As String)
    Dim ilRet As Integer
    Dim ilOutVefCode As Integer
    Dim ilOutBreakNo As Integer
    Dim slOutDay As String
    
    ilRet = mFindPlayCmmd(ilShttCode, llID, slPort, ilOutVefCode, ilOutBreakNo, slOutDay)
    If ilRet = False Then
        PlayCmmdInfo_rst.AddNew Array("shttCode", "vefCode", "BreakNo", "Day", "ID", "Port"), Array(ilShttCode, ilVefCode, Val(slBreakNo), slDay, llID, slPort)
    Else
        'Update PlayCmmdInfo
        PlayCmmdInfo_rst!vefCode = ilVefCode
        PlayCmmdInfo_rst!breakno = Val(slBreakNo)
        PlayCmmdInfo_rst!Day = slDay
    End If
End Sub

Private Function mCheckFile() As Integer
    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slLine As String
    Dim slLocation As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    mCheckFile = False
    ilFound = False
    On Error GoTo ErrHand
    
    slLocation = Trim$(txtFile.Text)
    If fs.FILEEXISTS(slLocation) Then
        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
    Else
        If Not igCompelAutoImport Then
            Beep
            gMsgBox "Unable to open the Import file: " & slLocation, vbCritical
        Else
            gLogMsgWODT "W", hmResult, "Unable to open the Import file: " & slLocation & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
        
    Do While tlTxtStream.AtEndOfStream <> True
        slLine = tlTxtStream.ReadLine
        slLine = UCase(slLine)
        ilPos = InStr(1, UCase(slLine), "COMMAND TO PLAY PLAYLIST", vbTextCompare)
        If ilPos > 0 Then
            gParseCDFields slLine, False, smFields()
            'smDate = Format(smFields(1), "m/d/yy")
            smDate = Format(smFields(0), "m/d/yy")
            ilFound = True
            mCheckFile = True
            Exit Do
        End If
        ilPos = InStr(1, UCase(slLine), "INPUT_PL_USER_TEXT", vbTextCompare)
        If ilPos > 0 Then
            gParseCDFields slLine, False, smFields()
            'smDate = Format(smFields(1), "m/d/yy")
            smDate = Format(smFields(0), "m/d/yy")
            ilFound = True
            mCheckFile = True
            Exit Do
        End If
    Loop
    tlTxtStream.Close
    If Not ilFound And Not igCompelAutoImport Then
        Beep
        gMsgBox "No 'Command to Play Playlist' records found in the Import file", vbCritical
    Else
        gLogMsgWODT "W", hmResult, "No 'Command to Play Playlist' records found in the Import file" & Format(Now, "mm-dd-yy")
    End If
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmImportWegener-mCheckFile: "
        gLogMsgWODT "W", hmResult, "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & Format(Now, "mm-dd-yy")
        If Not igCompelAutoImport Then
            gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        Else
            gLogMsgWODT "W", hmResult, gMsg & Err.Description & "; Error #" & Err.Number & " " & Format(Now, "mm-dd-yy")
        End If
    End If
End Function

Private Function mGetZoneAdj(ilVefCode As Integer, slZone As String) As Integer
    Dim ilLocalAdj As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilVef As Integer
    Dim ilZone As Integer
    
    ilLocalAdj = 0
    ilVef = gBinarySearchVef(CLng(ilVefCode))
    If Len(Trim$(slZone)) <> 0 Then
        'Get zone
        If ilVef <> -1 Then
            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = Trim$(slZone) Then
                    If (tgVehicleInfo(ilVef).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVef).sFed(ilZone)) <> "") And (tgVehicleInfo(ilVef).iBaseZone(ilZone) <> -1) Then
                        ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                    End If
                    Exit For
                End If
            Next ilZone
        End If
    End If
    mGetZoneAdj = ilLocalAdj
End Function

Private Sub mSaveAttCode(llAttCode As Long)
    Dim llLoop As Long

    For llLoop = 0 To UBound(lmAttCode) - 1 Step 1
        If lmAttCode(llLoop) = llAttCode Then
            Exit Sub
        End If
    Next llLoop
    lmAttCode(UBound(lmAttCode)) = llAttCode
    ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
End Sub

Private Function mInitPledgeError() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ShttCode", adInteger
        .Append "VefCode", adInteger
    End With
    rst.Open
    rst!shttCode.Properties("optimize") = True
    Set mInitPledgeError = rst
End Function

Private Sub mClosePledgeError()
    On Error Resume Next
    If Not PledgeError_rst Is Nothing Then
        If (PledgeError_rst.State And adStateOpen) <> 0 Then
            PledgeError_rst.Close
        End If
        Set PledgeError_rst = Nothing
    End If

End Sub

Private Function mAddPledgeError(ilShttCode As Integer, ilVefCode As Integer) As Integer
    mAddPledgeError = False
    PledgeError_rst.Filter = "ShttCode = " & ilShttCode & " And VefCode = " & ilVefCode
    If PledgeError_rst.EOF Then
        PledgeError_rst.AddNew Array("ShttCode", "VefCode"), Array(ilShttCode, ilVefCode)
        mAddPledgeError = True
    End If
End Function

Private Function mAttExist(ilVefCode As Integer, ilShttCode As Integer) As Long
    Dim ilRet As Integer
    '7458 changed to long!
    On Error GoTo ErrHand
    mAttExist = False
    SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
    SQLQuery = SQLQuery & " AND " & "attShfCode = " & ilShttCode
    SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(smDate), sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(smDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(smDate), sgSQLDateForm) & "')" & ")"
    Set att_rst = gSQLSelectCall(SQLQuery)
    If Not att_rst.EOF Then
        '7458
'        mAttExist = True
        mAttExist = att_rst!attCode
    End If
Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mBuildISCIInfo"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
End Function

Private Function mInitAgreementError() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ShttCode", adInteger
        .Append "VefCode", adInteger
    End With
    rst.Open
    rst!shttCode.Properties("optimize") = True
    Set mInitAgreementError = rst
End Function

Private Sub mCloseAgreementError()
    On Error Resume Next
    If Not AgreementError_rst Is Nothing Then
        If (AgreementError_rst.State And adStateOpen) <> 0 Then
            AgreementError_rst.Close
        End If
        Set AgreementError_rst = Nothing
    End If

End Sub

Private Function mAddAgreementError(ilShttCode As Integer, ilVefCode As Integer) As Integer
    mAddAgreementError = False
    AgreementError_rst.Filter = "ShttCode = " & ilShttCode & " And VefCode = " & ilVefCode
    If AgreementError_rst.EOF Then
        AgreementError_rst.AddNew Array("ShttCode", "VefCode"), Array(ilShttCode, ilVefCode)
        mAddAgreementError = True
    End If
End Function

Private Function mInitDayError() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ShttCode", adInteger
        .Append "VefCode", adInteger
    End With
    rst.Open
    rst!shttCode.Properties("optimize") = True
    Set mInitDayError = rst
End Function

Private Sub mCloseDayError()
    On Error Resume Next
    If Not DayError_rst Is Nothing Then
        If (DayError_rst.State And adStateOpen) <> 0 Then
            DayError_rst.Close
        End If
        Set DayError_rst = Nothing
    End If

End Sub

Private Function mAddDayError(ilShttCode As Integer, ilVefCode As Integer) As Integer
    mAddDayError = False
    DayError_rst.Filter = "ShttCode = " & ilShttCode & " And VefCode = " & ilVefCode
    If DayError_rst.EOF Then
        DayError_rst.AddNew Array("ShttCode", "VefCode"), Array(ilShttCode, ilVefCode)
        mAddDayError = True
    End If
End Function

Private Function mInitMissingStation() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "Serial", adChar, 10
        .Append "Port", adChar, 1
        .Append "CallLetters", adChar, 40
        .Append "Shown", adBoolean
    End With
    rst.Open
    rst!Serial.Properties("optimize") = True
    Set mInitMissingStation = rst
End Function

Private Sub mCloseMissingStation()
    On Error Resume Next
    If Not MissingStation_rst Is Nothing Then
        If (MissingStation_rst.State And adStateOpen) <> 0 Then
            MissingStation_rst.Close
        End If
        Set MissingStation_rst = Nothing
    End If

End Sub

Private Function mFindMissingStation(slSerial As String, slPort As String) As String
    mFindMissingStation = ""
    If slPort = "1" Then
        slPort = "A"
    ElseIf slPort = "2" Then
        slPort = "B"
    ElseIf slPort = "3" Then
        slPort = "C"
    ElseIf slPort = "4" Then
        slPort = "D"
    End If
    MissingStation_rst.Filter = "Serial = '" & slSerial & "' And Port = '" & slPort & "'"     ' & ilVefCode
    If Not MissingStation_rst.EOF Then
        If MissingStation_rst!Shown = False Then
            mFindMissingStation = Trim$(MissingStation_rst!CALLLETTERS)
            MissingStation_rst!Shown = True
        End If
    Else
        MissingStation_rst.AddNew Array("Serial", "Port", "CallLetters", "Shown"), Array(slSerial, slPort, "Missing- Serial: " & slSerial & " Port: " & slPort, True)
        mFindMissingStation = "Missing- Serial: " & slSerial & " Port: " & slPort
    End If
End Function

Private Function mInitLSTMap() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "LstCode", adInteger
        .Append "Zone", adChar, 1
        .Append "LogDate", adInteger
        .Append "LogTime", adInteger
        .Append "BreakNo", adInteger
        .Append "Position", adInteger
        '11/18/14: Add Event code
        .Append "GsfCode", adInteger
    End With
    rst.Open
    rst!lstCode.Properties("optimize") = True
    Set mInitLSTMap = rst
End Function

Private Sub mCloseLSTMap()
    On Error Resume Next
    If Not LSTMap_rst Is Nothing Then
        If (LSTMap_rst.State And adStateOpen) <> 0 Then
            LSTMap_rst.Close
        End If
        Set LSTMap_rst = Nothing
    End If

End Sub

Private Function mFindLSTMap(llLstCode As Long, slMapZone As String) As Long
    If slMapZone = "E" Then
        mFindLSTMap = llLstCode
        Exit Function
    End If
    mFindLSTMap = -1
    LSTMap_rst.Filter = "LstCode = " & llLstCode
    If Not LSTMap_rst.EOF Then
        '11/18/14: Add Event Code
        'LSTMap_rst.Filter = "Zone = '" & slMapZone & "' And LogDate = " & LSTMap_rst!LOGDATE & " And LogTime = " & LSTMap_rst!logtime & " And BreakNo = " & LSTMap_rst!BreakNo & " And Position = " & LSTMap_rst!Position
        LSTMap_rst.Filter = "Zone = '" & Left$(slMapZone, 1) & "' And LogDate = " & LSTMap_rst!LOGDATE & " And LogTime = " & LSTMap_rst!logtime & " And BreakNo = " & LSTMap_rst!breakno & " And Position = " & LSTMap_rst!Position & " And GsfCode = " & LSTMap_rst!gsfCode
        If Not LSTMap_rst.EOF Then
            mFindLSTMap = LSTMap_rst!lstCode
            Exit Function
        End If
    End If
End Function

Private Function mInitImportSpotInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "VefCode", adInteger
        .Append "ShttCode", adInteger
        .Append "AdfCode", adInteger
        .Append "BreakNo", adInteger
        '11/18/14: Add Event code
        .Append "GsfCode", adInteger
        .Append "AirDate", adInteger
        .Append "AirTime", adInteger
        .Append "ISCI", adChar, 20
        .Append "Row", adInteger
    End With
    rst.Open
    rst!vefCode.Properties("optimize") = True
    'rst.Sort = "BreakNo"
    Set mInitImportSpotInfo = rst
End Function

Private Sub mCloseImportSpotInfo()
    On Error Resume Next
    If Not ImportSpotInfo_rst Is Nothing Then
        If (ImportSpotInfo_rst.State And adStateOpen) <> 0 Then
            ImportSpotInfo_rst.Close
        End If
        Set ImportSpotInfo_rst = Nothing
    End If

End Sub

Private Function mProcessImportedSpots(llUpdateCount As Long, llNoAstCount As Long, llISCIErrorCount As Long, blPlayError As Boolean, blPledgeError As Boolean, blAgreementError As Boolean) As Integer
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim ilBreakNo As Integer
    Dim ilAdfCode As Integer
    Dim llAttCode As Long
    Dim ilPledgeStatus As Integer
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llAstCode As Long
    Dim llRow As Long
    Dim ilRet As Integer
    Dim slCallLetters As String
    Dim slVehicleName As String
    Dim llVef As Long
    Dim ilShtt As Integer
    Dim slImportISCI As String
    Dim slAstISCI As String
    Dim llAdf As Long
    '11/18/14: Add Event code
    Dim llGsfCode As Long
    '7458
    Dim slFeedDate As String
    Dim llAttButNoAst As Long
    
    ImportSpotInfo_rst.Sort = "VefCode,shttCode"
    ilVefCode = -1
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .ThirdParty = Vendors.Wegener_Compel
        .TypeEnt = Importposted3rdparty
        .User = igUstCode
        .ErrorLog = cmPathForgLogMsg
        .CreateCopyForWeb = True
        .fileName = Mid(txtFile.Text, InStrRev(txtFile.Text, "\") + 1)
    End With
    Do While Not ImportSpotInfo_rst.EOF
        DoEvents
        If (ImportSpotInfo_rst!vefCode <> ilVefCode) Then
            mBuildBreakArray ImportSpotInfo_rst!vefCode, smETLstFrom, smCTLstFrom, smMTLstFrom, smPTLstFrom
        End If
        If (ImportSpotInfo_rst!vefCode <> ilVefCode) Or (ImportSpotInfo_rst!shttCode <> ilShttCode) Then
            '7458
            If ilVefCode > 0 And ilShttCode > 0 Then
                'previously created agreement needs to be sent to ents
                If Not myEnt.CreateEnts() Then
                    gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
                End If
            End If
            ilVefCode = ImportSpotInfo_rst!vefCode
            llVef = gBinarySearchVef(CLng(ilVefCode))
            If llVef <> -1 Then
                slVehicleName = Trim$(tgVehicleInfo(llVef).sVehicle)
            Else
                slVehicleName = "Vehicle Missing: " & ilVefCode
            End If
            ilShttCode = ImportSpotInfo_rst!shttCode
            ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
            If ilShtt <> -1 Then
                slCallLetters = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
            Else
                slCallLetters = "Station Missing: " & ilShttCode
            End If
            '7458
            With myEnt
                .Vehicle = ilVefCode
                .Station = ilShttCode
                .ProcessStart
            End With
            ilRet = mCreateAST(ilVefCode, ilShttCode)
            '7458
            With myEnt
                If Not AstInfo_rst.EOF Then
                    .Agreement = AstInfo_rst!attCode
                Else
                    .Agreement = 0
                End If
            End With
        End If
        ilBreakNo = ImportSpotInfo_rst!breakno
        '11/18/14: Add Event Code
        llGsfCode = ImportSpotInfo_rst!gsfCode
        ilAdfCode = ImportSpotInfo_rst!adfCode
        llAdf = gBinarySearchAdf(CLng(ilAdfCode))
        slAirDate = Format(ImportSpotInfo_rst!airDate, "m/d/yy")
        slAirTime = gLongToTime(ImportSpotInfo_rst!airTime)
        slImportISCI = ImportSpotInfo_rst!ISCI
        llRow = ImportSpotInfo_rst!Row
        '11/18/14: Add Event code
        'llAstCode = mFindAstCode(ilShttCode, ilVefCode, ilBreakNo, ilAdfCode, llAttCode, ilPledgeStatus, slAstISCI)
        '7458 need feeddate
'        llAstCode = mFindAstCode(ilShttCode, ilVefCode, ilBreakNo, llGsfCode, slImportISCI, ilAdfCode, llAttCode, ilPledgeStatus, slAstISCI)
        llAstCode = mFindAstCode(ilShttCode, ilVefCode, ilBreakNo, llGsfCode, slImportISCI, ilAdfCode, llAttCode, ilPledgeStatus, slAstISCI, slFeedDate)
'        '7458
'        myEnt.Agreement = llAttCode

        '************************** Start Compel Auto Import *****************************
        If UBound(tmCompelExportInfo) = 0 Then
            lmMaxCmpRecs = 10000
            lmExptCmpCnt = 0
            ReDim tmCompelExportInfo(0 To lmMaxCmpRecs) As COMPELEXPORTINFO
        Else
            If lmExptCmpCnt = lmMaxCmpRecs Then
                lmMaxCmpRecs = lmMaxCmpRecs + 5000
                ReDim Preserve tmCompelExportInfo(0 To lmMaxCmpRecs) As COMPELEXPORTINFO
            End If
        End If
        
        If llAstCode <> -1 Then
            slCompelStatus = "Posted"
        Else
            slCompelStatus = "Not Posted"
        End If

        tmCompelExportInfo(lmExptCmpCnt).sVefName = gGetVehNameByVefCode(ilVefCode)
        tmCompelExportInfo(lmExptCmpCnt).sCallLetters = Trim$(gGetCallLettersByShttCode(ilShttCode))
        If llGsfCode > 0 Then
            tmCompelExportInfo(lmExptCmpCnt).sGame = mGetGameTeamNames(llGsfCode)
        Else
            tmCompelExportInfo(lmExptCmpCnt).sGame = ""
        End If
        tmCompelExportInfo(lmExptCmpCnt).sAiredDate = slAirDate
        tmCompelExportInfo(lmExptCmpCnt).sAiredTime = slAirTime
        tmCompelExportInfo(lmExptCmpCnt).ilBreakNo = ilBreakNo
        ilAdfCode = gBinarySearchAdf(CLng(ilAdfCode))
        If ilAdfCode > 0 Then
            tmCompelExportInfo(lmExptCmpCnt).sAdvertiser = Trim$(tgAdvtInfo(ilAdfCode).sAdvtName)
        Else
            tmCompelExportInfo(lmExptCmpCnt).sAdvertiser = "Undefined"
        End If
        tmCompelExportInfo(lmExptCmpCnt).sISCI = Trim(slImportISCI)
        tmCompelExportInfo(lmExptCmpCnt).sStatus = Trim(slCompelStatus)
        tmCompelExportInfo(lmExptCmpCnt).lRow = llRow
        lmExptCmpCnt = lmExptCmpCnt + 1
        '************************** End Compel Auto Import *******************************

        If llAstCode <> -1 Then
            If StrComp(UCase(Trim$(slImportISCI)), UCase(Trim$(slAstISCI)), vbBinaryCompare) <> 0 Then
                llISCIErrorCount = llISCIErrorCount + 1
                If llAdf <> -1 Then
                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": ISCI not matching- " & Trim$(tgAdvtInfo(llAdf).sAdvtName) & " Exported " & Trim$(slAstISCI) & ", Imported " & Trim$(slImportISCI) & " (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & ")"
                Else
                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": ISCI not matching- Exported " & Trim$(slAstISCI) & ", Imported " & Trim$(slImportISCI) & " (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & ")"
                End If
            End If
            llUpdateCount = llUpdateCount + 1
            'Update ast
            SQLQuery = "UPDATE ast SET astCPStatus = 1, "
            SQLQuery = SQLQuery & "astAirDate = '" & Format$(slAirDate, sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "astAirTime = '" & Format$(slAirTime, sgSQLTimeForm) & "'"
            SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mProcessImportedSpots"
                Exit Function
            End If
            If UBound(tmExportWebSpot) = 0 Then
                lmMaxRecs = 10000
                lmExportWebCount = 0
                ReDim tmExportWebSpot(0 To lmMaxRecs) As EXPORTASTINFO
            Else
                If lmExportWebCount = lmMaxRecs Then
                    lmMaxRecs = lmMaxRecs + 5000
                    ReDim Preserve tmExportWebSpot(0 To lmMaxRecs)
                End If
            End If
            tmExportWebSpot(lmExportWebCount).lAstCode = llAstCode
            tmExportWebSpot(lmExportWebCount).lAttCode = llAttCode
            tmExportWebSpot(lmExportWebCount).iShttCode = ilShttCode
            tmExportWebSpot(lmExportWebCount).iVefCode = ilVefCode
            tmExportWebSpot(lmExportWebCount).sAiredDate = Format$(slAirDate, sgSQLDateForm)
            tmExportWebSpot(lmExportWebCount).sAiredTime = Format$(slAirTime, sgSQLTimeForm)
            tmExportWebSpot(lmExportWebCount).sISCI = slAstISCI
            tmExportWebSpot(lmExportWebCount).gsfCode = llGsfCode
            tmExportWebSpot(lmExportWebCount).sType = "C"
            lmExportWebCount = lmExportWebCount + 1
            If tgStatusTypes(gGetAirStatus(ilPledgeStatus)).iPledged = 2 Then
                If mAddPledgeError(ilShttCode, ilVefCode) Then
                    If llAdf <> -1 Then
                        gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Spot Aired but Pledge set to Not Carried (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & " " & Trim$(tgAdvtInfo(llAdf).sAdvtName) & ")"
                    Else
                        gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Spot Aired but Pledge set to Not Carried (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & ")"
                    End If
                End If
                If Not blPledgeError Then
                    SetResults "Spot Aired but Pledge set to Not Carried", 0
                    DoEvents
                    blPledgeError = True
                End If
            End If
            mSaveAttCode llAttCode
            '7458
            If Len(slFeedDate) > 0 Then
                If Not myEnt.Add(slFeedDate, llGsfCode, Ingested) Then
                    gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
                End If
            End If
        Else
            'Output error message
            llNoAstCount = llNoAstCount + 1
            '7458 changed to return att
            llAttButNoAst = mAttExist(ilVefCode, ilShttCode)
'            If mAttExist(ilVefCode, ilShttCode) Then
            If llAttButNoAst > 0 Then
                myEnt.Agreement = llAttButNoAst
                If llAdf <> -1 Then
                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Unable to Find Matching Spot (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & " " & Trim$(tgAdvtInfo(llAdf).sAdvtName) & ")"
                Else
                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Unable to Find Matching Spot (Row " & llRow & ": " & slAirDate & " " & slAirTime & " Break " & ilBreakNo & ")"
                End If
                If Not blPlayError Then
                    SetResults "Unable to Find Matching Affiliate Spot for Import Spot", 0
                    DoEvents
                    blPlayError = True
                End If
            Else
                If mAddAgreementError(ilShttCode, ilVefCode) Then
                    gLogMsgWODT "W", hmResult, "  " & slCallLetters & " on " & slVehicleName & ": Agreement missing (Row " & llRow & ")"
                End If
                If Not blAgreementError Then
                    SetResults "Agreement Missing", 0
                    DoEvents
                    blAgreementError = True
                End If
            End If
            '7458 have to add airdate, not feed date. may have no agreement #
            If Not myEnt.Add(slAirDate, llGsfCode, SentOrReceived) Then
                gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
            End If
        End If
        ImportSpotInfo_rst.MoveNext
    Loop
    '7458
    If Not myEnt.CreateEnts() Then
        gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
    End If
    ReDim Preserve tmCompelExportInfo(0 To lmExptCmpCnt)
    ilRet = mWriteCompelFile
    mProcessImportedSpots = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mProcessImportedSpots"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Exit Function
    mProcessImportedSpots = False
End Function

Private Sub mWebExportSpots()
    
    Dim llIdx As Long   'Integer
    Dim ilMissedReasonCode As Integer
    Dim llAstCode As Long
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim llWebExpCount As Long
    Dim llAttCode As Long
    Dim slISCI As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slExportType As String
    Dim slSpotLine As String
    Dim slSpotHeader As String
    Dim slNowTime As String
    Dim slNowDate As String
    Dim lmMaxEsfCode As Long
    Dim llAddRecs As Long
    Dim llDelRecs As Long
    Dim llTtlRecs As Long
    Dim llFileSize As Long
    Dim ilFileSent As Integer
    Dim ilRet As Integer
    Dim ilStartNewFiles As Integer
    Dim ilRetries As Integer
    '7458
    Dim ilFtpOk As Integer
    
    slSpotHeader = "astCode, attCode, ISCI, ActualDate, ActualTime, MissedReason, Source"
    Print #hmWebToDetail, slSpotHeader
    llWebExpCount = 0
    For llIdx = 0 To lmExportWebCount - 1 Step 1
        llWebExpCount = llWebExpCount + 1
        llAstCode = tmExportWebSpot(llIdx).lAstCode
        llAttCode = tmExportWebSpot(llIdx).lAttCode
        ilVefCode = tmExportWebSpot(llIdx).iVefCode
        ilShttCode = tmExportWebSpot(llIdx).iShttCode
        slISCI = Trim$(tmExportWebSpot(llIdx).sISCI)
        slAirDate = tmExportWebSpot(llIdx).sAiredDate
        slAirTime = tmExportWebSpot(llIdx).sAiredTime
        slExportType = tmExportWebSpot(llIdx).sType
        'Get Missed reason reference
        SQLQuery = "SELECT altMnfMissed FROM alt where altAstcode = " & llAstCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            ilMissedReasonCode = rst!altMnfMissed
        Else
            ilMissedReasonCode = 0
        End If
        slSpotLine = llAstCode & ", " & llAttCode & ", " & """" & slISCI & """," & slAirDate & ", " & slAirTime & ", " & ilMissedReasonCode & ", " & """" & slExportType & """"
        Print #hmWebToDetail, slSpotLine
        
    Next llIdx
        
    llWebExpCount = llWebExpCount
    mGetMissedReasons
    Close hmWebToDetail
    Close hmWebToHeader
    'Wait on any unfinished FTP jobs
    lgSTime3 = timeGetTime
    
    '10/3  For Testing Only
    'imFtpInProgress = False
    While imFtpInProgress
        ilRet = mCheckFTPStatus()
        Sleep (1000)
        DoEvents
    Wend
    lgETime3 = timeGetTime
    lgTtlTime3 = lgTtlTime3 + (lgETime3 - lgSTime3)
    ReDim mFtpArray(0 To 0)
    mFtpArray(0) = Trim$(smWebSpots)
    SetResults "    ", 0
    gLogMsgWODT "W", hmResult, "    "
    SetResults "*** Starting Export to Web Process ***", 0
    gLogMsgWODT "W", hmResult, "*** Starting Export to Web Process ***"
    SetResults "FTPing - " & Trim$(smWebSpots), 0
    gLogMsgWODT "W", hmResult, "FTPing - " & Trim$(smWebSpots)
    DoEvents
    gLogMsgWODT "W", hmResult, "Sending " & Trim$(mFtpArray(0)) & " " & Format(Now, "mm-dd-yy")
    gLogMsgWODT "W", hmResult, "Sending Files to web site. " & " " & Format(Now, "mm-dd-yy")
    'Dan added for testing
    '10/3  For Testing Only
    'bmFakeWeb = True
    If Not bmFakeWeb Then
        imFtpInProgress = True
        ilRet = csiFTPFileToServer(Trim$(mFtpArray(0)))
        ReDim Preserve mFtpArray(UBound(mFtpArray) + 1)
        While imFtpInProgress
            '7458
            'ilRet = mCheckFTPStatus()
            ilFtpOk = mCheckFTPStatus()
            Sleep (1000)
            DoEvents
        Wend
        lgSTime6 = timeGetTime
        ilRetries = 0
        imWaiting = True
        mProcessWebQueue
        While (imSomeThingToDo = True Or imWaiting = True) And Not imTerminate
            If imWaiting Then
                DoEvents
                ilRet = mCheckStatus()
                Sleep (1000)
            End If
        Wend
    Else
        ilFtpOk = -1
    End If
    '7458

    'New
    lmExportWebCount = 0
    ReDim tmExportWebSpot(0 To 0) As EXPORTASTINFO
    
    
    If bmFTPIsOn Then
        If ilFtpOk Then
            If Not myEnt.UpdateIncompleteByFilename(Successful, , smWebSpots) Then
                gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
            End If
        Else
            If Not myEnt.UpdateIncompleteByFilename(EntError, , smWebSpots) Then
                gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
            End If
        End If
    Else
        If Not myEnt.UpdateIncompleteByFilename(NotSent, , smWebSpots) Then
            gLogMsgWODT "W", hmResult, myEnt.ErrorMessage
        End If
    End If
    If ilRetries = 5 Then
        gLogMsgWODT "W", hmResult, "Error: Retries were exceeded in frmWebExportSchdSpot - mInitiateExport" & " " & Format(Now, "mm-dd-yy")
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & " " & Format(Now, "mm-dd-yy") & ".txt", "Export Wegener-mWebExportSpots"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
End Sub

Private Function mCheckStatus() As Integer
    Dim ilRet As Integer
    Dim slFile As String
    Dim slTemp As String
    Dim ilPos As Integer
    
    On Error GoTo ErrHand
    If (igDemoMode) Then
        imWaiting = False
        Exit Function
    End If
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    On Error GoTo ErrHand2:
    slTemp = smWebExports & Trim$(smWebSpots)
    '8886
   ' slFile = Dir(slTemp)
    'If Len(slFile) = 0 Then
     If gFileExist(slTemp) = FILEEXISTSNOT Then
        Exit Function
    End If
    On Error GoTo ErrHand
    ilRet = mExCheckWebWorkStatus(Trim$(smWebSpots), "Wegener Updates")
    If ilRet = True Then
        gLogMsgWODT "W", hmResult, Trim$(smWebSpots) & " Import Successful." & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "W", hmResult, Trim$(smWebSpots) & " Import Successful." & " " & Format(Now, "mm-dd-yy")
        SetResults "     -- " & Trim$("Wegener Updates") & " Imp. Successful.", 0
        DoEvents
        ilPos = 0
        ilPos = InStr(slTemp, "WebSpots")
        imWaiting = False
        imSomeThingToDo = False
        Call gEndWebSession("WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt")
    Else
        imWaiting = True
    End If
    Exit Function
ErrHand2:
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Exit Function
End Function

Private Function mWebOpenFiles(sExpType As String) As Integer

    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slTemp As String
    
    mWebOpenFiles = False
    imExporting = True
    ilRet = 0
    On Error GoTo cmdExportErr:
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & sExpType & ".txt"
    slTemp = slTemp & "_" & sgUserName & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & "_C" & ".txt"
    smWebSpots = "WebUpdateSpots_" & slTemp
    smWebToFileDetail = smWebExports & smWebSpots
    smWebHeader = "WebHeaders_" & slTemp
    smWebToFileHeader = smWebExports & smWebHeader
    smWebCopyRot = "CpyRotCom_" & slTemp
    smWebToCpyRot = smWebExports & smWebCopyRot
    smWebMultiUse = "MultiUse_" & slTemp
    smWebToMultiUse = smWebExports & smWebMultiUse
    'D.S. Check bit map to see if using games.  If not no sense in exporting of showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
    smWebEventInfo = "EventInfo_" & slTemp
    smWebToEventInfo = smWebExports & smWebEventInfo
    End If
    'slDateTime = FileDateTime(smWebToFileDetail)
    ilRet = gFileExist(smWebToFileDetail)
    If ilRet = 0 Then
        Screen.MousePointer = vbDefault
        slDateTime = gFileDateTime(smWebToFileDetail)
        If Not igCompelAutoImport Then
            ilRet = gMsgBox("Export Previously Created " & slDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        Else
            gLogMsgWODT "W", hmResult, "Export Previously Created " & slDateTime & " Continue with Export by Replacing File"
        End If
        If ilRet = vbCancel Then
            gLogMsgWODT "W", hmResult, "** Terminated Because Export File Existed **" & " " & Format(Now, "mm-dd-yy")
            Close #hmWebToDetail
            imExporting = False
            Exit Function
        End If
        Screen.MousePointer = vbHourglass
        Kill smWebToFileDetail
        Kill smWebToFileHeader
        Kill smWebToCpyRot
        Kill smWebToMultiUse
    End If
    On Error GoTo 0
    'ilRet = 0
    On Error GoTo cmdExportErr:
    'hmWebToDetail = FreeFile
    'Open smWebToFileDetail For Output Lock Write As hmWebToDetail
    ilRet = gFileOpen(smWebToFileDetail, "Output Lock Write", hmWebToDetail)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "** Terminated - " & smWebToFileDetail & " failed to open. **" & " " & Format(Now, "mm-dd-yy")
        Close #hmWebToDetail
        imExporting = False
        Screen.MousePointer = vbDefault
        If Not igCompelAutoImport Then
            gMsgBox "Open Error #" & Str$(Err.Numner) & smWebToFileDetail, vbOKOnly, "Open Error"
        Else
            gLogMsgWODT "W", hmResult, "Open Error #" & Str$(Err.Numner) & smWebToFileDetail & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
    'hmWebToHeader = FreeFile
    'Open smWebToFileHeader For Output Lock Write As hmWebToHeader
    ilRet = gFileOpen(smWebToFileHeader, "Output Lock Write", hmWebToHeader)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "** Terminated - " & smWebToFileHeader & " failed to open. **" & " " & Format(Now, "mm-dd-yy")
        Close #hmWebToDetail
        Close #hmWebToHeader
        imExporting = False
        Screen.MousePointer = vbDefault
        If Not igCompelAutoImport Then
            gMsgBox "Open Error #" & Str$(Err.Number) & smWebToFileHeader, vbOKOnly, "Open Error"
        Else
            gLogMsgWODT "W", hmResult, "Open Error #" & Str$(Err.Number) & smWebToFileHeader & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
    'hmWebToCpyRot = FreeFile
    'Open smWebToCpyRot For Output Lock Write As hmWebToCpyRot
    ilRet = gFileOpen(smWebToCpyRot, "Output Lock Write", hmWebToCpyRot)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "** Terminated - " & smWebToCpyRot & " failed to open. **" & " " & Format(Now, "mm-dd-yy")
        Close #hmWebToDetail
        Close #hmWebToHeader
        Close #hmWebToCpyRot
        imExporting = False
        Screen.MousePointer = vbDefault
        If Not igCompelAutoImport Then
            gMsgBox "Open Error #" & Str$(Err.Numner) & smWebToCpyRot, vbOKOnly, "Open Error"
        Else
            gLogMsgWODT "W", hmResult, "Open Error #" & Str$(Err.Numner) & smWebToCpyRot & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
    'hmWebToMultiUse = FreeFile
    'Open smWebToMultiUse For Output Lock Write As hmWebToMultiUse
    ilRet = gFileOpen(smWebToMultiUse, "Output Lock Write", hmWebToMultiUse)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "** Terminated - " & smWebToMultiUse & " failed to open. **" & " " & Format(Now, "mm-dd-yy")
        Close #hmWebToDetail
        Close #hmWebToHeader
        Close #hmWebToCpyRot
        Close #hmWebToMultiUse
        imExporting = False
        Screen.MousePointer = vbDefault
        If Not igCompelAutoImport Then
            gMsgBox "Open Error #" & Str$(Err.Numner) & smWebToMultiUse, vbOKOnly, "Open Error"
        Else
            gLogMsgWODT "W", hmResult, "Open Error #" & Str$(Err.Numner) & smWebToMultiUse & " " & Format(Now, "mm-dd-yy")
        End If
        Exit Function
    End If
    'D.S. Check bit map to see if using games.  If not no sense in exporting or showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        'smWebToEventInfo = FreeFile
        'Open smWebToEventInfo For Output Lock Write As hmWebToEventInfo
        ilRet = gFileOpen(smWebToEventInfo, "Output Lock Write", hmWebToEventInfo)
        If ilRet <> 0 Then
            gLogMsgWODT "W", hmResult, "** Terminated - " & smWebToEventInfo & " failed to open. **" & " " & Format(Now, "mm-dd-yy")
            Close #hmWebToDetail
            Close #hmWebToHeader
            Close #hmWebToCpyRot
            Close #hmWebToMultiUse
            Close #hmWebToMultiUse
            imExporting = False
            Screen.MousePointer = vbDefault
            If Not igCompelAutoImport Then
               gMsgBox "Open Error #" & Str$(Err.Numner) & smWebToEventInfo, vbOKOnly, "Open Error"
            Else
                gLogMsgWODT "W", hmResult, "Open Error #" & Str$(Err.Numner) & smWebToEventInfo & " " & Format(Now, "mm-dd-yy")
            End If
            Exit Function
        End If
    End If
    DoEvents
    Print #hmWebToHeader, gBuildWebHeaderDetail()
    Print #hmWebToCpyRot, "Code, Comment"
    'D.S. Check bit map to see if using games.  If not no sense in exporting of showing it
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) = USINGSPORTS) Then
        Print #hmWebToMultiUse, "Code, GameDate, GameStartTime, VisitTeamName, VisitTeamAbbr, HomeTeamName, HomeTeamAbbr, LanguageCode, FeedSource, EventCarried, AttCode"
    End If
    gLogMsgWODT "W", hmResult, "** Storing Output into " & smWebToFileDetail & " And " & smWebToFileHeader & "**" & " " & Format(Now, "mm-dd-yy")
    mWebOpenFiles = True
    Exit Function
cmdExportErr:
'    ilRet = Err
'    Resume Next
End Function

Private Function mGetMissedReasons() As Integer

    Dim ilRet As Integer
    Dim rst As ADODB.Recordset
    Dim slResults As String
    Dim slStr As String

    On Error GoTo Err_Handler
    mGetMissedReasons = False
    lmTtlMultiUse = 0
    Print #hmWebToMultiUse, "[MissedReasons]"
    Print #hmWebToMultiUse, "Code, Reason"
    SQLQuery = "select mnfCode, mnfName from MNF_Multi_Names where mnftype = 'M' and (mnfCodeStn = 'A' or mnfCodeStn = 'B')"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        slStr = Trim$(rst!mnfCode) & "," & """" & Trim$(rst!mnfName) & """"
        Print #hmWebToMultiUse, slStr
        lmTtlMultiUse = lmTtlMultiUse + 1
        lmFileMultiUseCount = lmFileMultiUseCount + 1
        rst.MoveNext
    Wend
    ilRet = ilRet
    mGetMissedReasons = True
    rst.Close
    Close hmWebToMultiUse
    Exit Function
Err_Handler:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & " " & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener -mCheckFTPStatus"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    'debug
    'Resume Next
    Exit Function
End Function

Public Function mCheckFTPStatus() As Boolean

    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    On Error GoTo ErrHand
    If (igDemoMode) Then
        imFtpInProgress = False
        Exit Function
    End If
    mCheckFTPStatus = False
    imFtpInProgress = True
    ilRet = csiFTPGetStatus(tmCsiFtpStatus)
    '1 = Busy, 0 = Not Busy
    If tmCsiFtpStatus.iState = 1 Then
        Exit Function
    Else
        If tmCsiFtpStatus.iStatus <> 0 Then
            ' Errors occured.
            ilRet = csiFTPGetError(tmCsiFtpErrorInfo)
            If igCompelAutoImport Then
                gLogMsgWODT "W", hmResult, "Error: FAILED to FTP " & " & " & tmCsiFtpErrorInfo.sInfo & " " & Format(Now, "mm-dd-yy")
            Else
                MsgBox "FTP Failed. " & tmCsiFtpErrorInfo.sInfo
            End If
            gLogMsgWODT "W", hmResult, "Error: " & "FAILED to FTP " & tmCsiFtpErrorInfo.sFileThatFailed & " " & Format(Now, "mm-dd-yy")
            gLogMsgWODT "W", hmResult, "Error: " & "FAILED to FTP " & tmCsiFtpErrorInfo.sFileThatFailed & " " & Format(Now, "mm-dd-yy")
            SetResults "FTP Failed. ", 0
            gLogMsgWODT "W", hmResult, "FTP Failed. "
            DoEvents
            Exit Function
        Else
            For ilLoop = 0 To UBound(mFtpArray) - 1 Step 1
                ilRet = gTestFTPFileExists(mFtpArray(ilLoop))
'                If igCompelAutoImport Then
                    ilRet = 1
'                End If
                If ilRet = 1 Then
                    SetResults "Success, FTP - " & mFtpArray(ilLoop), 0
                    gLogMsgWODT "W", hmResult, "Success, FTP - " & mFtpArray(ilLoop)
                    DoEvents
                    gLogMsgWODT "W", hmResult, "   Success, FTP - " & mFtpArray(ilLoop)
                    imFtpInProgress = False
                    mCheckFTPStatus = True
                End If
            Next ilLoop
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener -mCheckFTPStatus"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    'debug
    'Resume Next
    Exit Function
End Function

Private Sub mProcessWebQueue()

    Dim ilRet As Integer
    Dim ilRetry As Integer
    Dim ilLoop As Integer
    Dim ilAllWentOK As Integer
    Dim slTemp As String
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    
    On Error GoTo ErrHand
    imSomeThingToDo = True
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    'smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & ".txt"
    Call mWaitForWebLock
    ilAllWentOK = True
    SetResults "- Imp. " & smWebSpots, 0

    gLogMsgWODT "W", hmResult, "- Imp. " & smWebSpots
    If Not gExecExtStoredProc(smWebSpots, "ImportUpdatedSpots.exe", False, False) Then
        SetResults "FAIL: Unable to instruct Web site to run ImportUpdatedSpots.exe", RGB(255, 0, 0)
        gLogMsgWODT "W", hmResult, "FAIL: Unable to instruct Web site to run ImportUpdatedSpots.exe" & " " & Format(Now, "mm-dd-yy")
        cmdCancel.Caption = "&Done"
        Screen.MousePointer = vbDefault
        ilAllWentOK = False
    Else
        gLogMsgWODT "W", hmResult, "Importing " & smWebSpots & " " & Format(Now, "mm-dd-yy")
        imWaiting = True
    End If
    DoEvents
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener - mProcessWebQueue"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    'debug
    'Resume Next
    Exit Sub
End Sub

'***************************************************************************************
' JD 08-22-2007
' This function was added to handle a special case occurring in the function
' mCheckWebWorkStatus. We believe a network error is causing the error handler
' to fire. Adding retry code to the function mCheckWebWorkStatus itself did not
' seem feasable because we did not know where the error was actually occuring and
' simplying calling a resume next could cause even more trouble.
'
'***************************************************************************************
Private Function mExCheckWebWorkStatus(sFileName As String, sTypeExpected As String) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String

    On Error GoTo Err_Handler:
    mExCheckWebWorkStatus = -1
    If (igDemoMode) Then
        mExCheckWebWorkStatus = 0
        Exit Function
    End If
    For ilLoop = 1 To 10
        ilRet = mCheckWebWorkStatus(sFileName, sTypeExpected)
        mExCheckWebWorkStatus = ilRet
        If ilRet <> -2 Then ' Retry only when this status is returned.
            Exit Function
        End If
        gLogMsgWODT "W", hmResult, "mExCheckWebWorkStatus is retrying due to an error in mCheckWebWorkStatus" & " " & Format(Now, "mm-dd-yy")
        DoEvents
        Sleep 2000  ' Delay for two seconds when retrying.
    Next
    If ilRet = -2 Then
        ilRet = -1  ' Keep the original error of -1 so all callers can process the error normally.
        gMsg = "A timeout has occured in frmWebExportSchdSpot - mExCheckWebWorkStatus"
        gLogMsgWODT "W", hmResult, gMsg & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "W", hmResult, " " & " " & Format(Now, "mm-dd-yy")
    End If
    Exit Function
Err_Handler:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener -mCheckFTPStatus"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    'debug
    'Resume Next
    Exit Function
End Function

Private Sub mWaitForWebLock()
    On Error GoTo ErrHandler
    Dim ilLoop As Integer
    Dim ilTotalMinutes As Integer
    Dim ilNotSaidWebServerWasBusy As Boolean
    Dim slLastMessage As String
    Dim slThisMessage As String
    Dim ilRow As Integer
    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String
    Dim ilRet As Integer

    ilNotSaidWebServerWasBusy = False
    slLastMessage = "Nothing"
    While 1
        DoEvents
        ilTotalMinutes = gStartWebSession("WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt")
        If ilTotalMinutes = 0 Then
            'Start the Export Process
            gLogMsgWODT "W", hmResult, "Web Session Started Successfully" & Format(Now, "mm-dd-yy")
            Exit Sub
        End If
        If Not ilNotSaidWebServerWasBusy Then
            ilNotSaidWebServerWasBusy = True
            SetResults "The Server is Busy. Standby...", 0
            gLogMsgWODT "W", hmResult, "The Server is Busy. Standby..."
            DoEvents
        End If
        If ilTotalMinutes > 1 Then
            slThisMessage = "Max wait time is " & Trim(Str(ilTotalMinutes)) & " Minutes."
        Else
            slThisMessage = "Max wait time is " & Trim(Str(ilTotalMinutes)) & " Minute."
        End If
        If slThisMessage <> slLastMessage Then
            ilRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slLastMessage)
            If lbcMsg.ListCount And ilRow >= 0 Then
                lbcMsg.RemoveItem ilRow
            End If
            SetResults slThisMessage, 0
            gLogMsgWODT "W", hmResult, slThisMessage
            slLastMessage = slThisMessage
        End If
        ' Wait here for 15 seconds. This loop allows the cancel button to be pressed as well.
        For ilLoop = 0 To 60
            If imTerminate Then
                Exit Sub
            End If
            DoEvents
            Sleep (250)   ' Wait 1/4 of a second
        Next
    Wend
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportWegener-mWaitForWebLock"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Exit Sub
End Sub

Private Function mCheckWebWorkStatus(sFileName As String, sTypeExpected As String) As Integer
 
    'D.S. 6/22/05
    
    'input - sFilemane is the unique file name that is the key into the web
    'server database to check it's status
    'Web Server Status - 0 = Done, 1 = Working and 2 = Error
    'Loop while the web server is busy processing spots and emails
    'Check the server every 10 seconds Report status
    
    Dim llWaitTime As Long
    Dim ilModResult As Integer
    Dim imStatus As Integer
    Dim slResult As String
    Dim llNumRows As Long
    Dim ilTimedOut As Integer
    Dim ilRet As Integer
    Dim ilWaitValue As Integer
    'Debug information
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    'Number of Seconds to Sleep
    Const clNumSecsToSleep As Long = 2
    Const clSleepValue As Long = clNumSecsToSleep * cmOneSecond
    'Assuming clNumSecsToSleep is 10 then a mod value of 6 would
    'be 6 loops at 10 seconds each or 1 minute
    Const clModValue As Integer = 6
    
    On Error GoTo ErrHand
    mCheckWebWorkStatus = False
    If Not gHasWebAccess() Then
        Exit Function
    End If
    ilWaitValue = 2
    llWaitTime = 0
    imStatus = 1
    ilRet = False
    Do While imStatus = 1 And llWaitTime < ilWaitValue And ilRet = False
        DoEvents
        If imTerminate Then
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imExporting = False
            SetResults "Export was canceled.", 0
            gLogMsgWODT "W", hmResult, "** User Terminated **" & " " & Format(Now, "mm-dd-yy")
            gLogMsgWODT "W", hmResult, "** User Terminated **" & " " & Format(Now, "mm-dd-yy")
            Exit Function
        End If
        SQLQuery = "Select Count(*) from WorkStatus Where FileName = " & "'" & sFileName & "'"
        llNumRows = gExecWebSQLWithRowsEffected(SQLQuery)
        llWaitTime = llWaitTime + 1
        ilModResult = llWaitTime Mod clModValue
        If llNumRows = -1 Then
            'An error was returned
            imStatus = 2
            smStatus = "2"
        End If
        If llNumRows > 0 Then
            SQLQuery = "Select FileName, Status, Msg1, Msg2, DTStamp from WorkStatus Where FileName = " & "'" & sFileName & "'"
            'Get the status information from the web server database and write it to a file
            Call gRemoteExecSql(SQLQuery, smWebWorkStatus, "WebExports", True, True, 30)
            DoEvents
            smStatus = "1"
            ilRet = mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports", sTypeExpected)
            llWaitTime = llWaitTime + 1
            ilModResult = llWaitTime Mod clModValue
            imStatus = CInt(smStatus)
            'Handle Web Error Condition
            If imStatus = 2 Then
                gLogMsgWODT "W", hmResult, "   " & "The Web Server Returned an ERROR. See Below. " & " " & Format(Now, "mm-dd-yy")
                gLogMsgWODT "W", hmResult, "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
                Call gEndWebSession("WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt")
                mCheckWebWorkStatus = False
                Exit Function
            End If
            If ilModResult = 0 And imStatus = 1 Then
                DoEvents
                SetResults "   " & smMsg1, 0
                SetResults "   " & smMsg2, 0
                gLogMsgWODT "W", hmResult, "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
                gLogMsgWODT "W", hmResult, "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
                DoEvents
            End If
        End If
        If imStatus = 1 Then
            Sleep clSleepValue
        End If
    Loop
    If llWaitTime >= 900 Then
        'We timed out
        gLogMsgWODT "W", hmResult, "   " & "A timeout occured while waiting on the web server for a response." & " " & Format(Now, "mm-dd-yy")
        SetResults "A timeout waiting on a web server response.", 0
        gLogMsgWODT "W", hmResult, "A timeout waiting on a web server response."
        DoEvents
        Call gEndWebSession("WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt")
        mCheckWebWorkStatus = False
        Exit Function
    End If
    'Show the final message with the totals of spots imported an emails sent
    imStatus = 0
    On Error Resume Next
    imStatus = CInt(smStatus)
    On Error GoTo ErrHand
    'Handle Web Error Condition
    If imStatus = 2 Then
        gLogMsgWODT "W", hmResult, "   " & "The Web Server Returned an ERROR. See Below. " & " " & Format(Now, "mm-dd-yy")
        gLogMsgWODT "W", hmResult, "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
        Call gEndWebSession("WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt")
        mCheckWebWorkStatus = False
        Exit Function
    End If
    If ilRet Then
    gLogMsgWODT "W", hmResult, "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
    gLogMsgWODT "W", hmResult, "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm") & " " & Format(Now, "mm-dd-yy")
    mCheckWebWorkStatus = True
    End If
Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    mCheckWebWorkStatus = -2
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "A general error has occured in frmWebExportSchdSpot - mCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsgWODT "W", hmResult, gMsg & " " & Format(Now, "mm-dd-yy")
    Exit Function
End Function

Private Function mProcessWebWorkStatusResults(sFileName As String, sIniValue As String, sTypeExpected) As Boolean

    'D.S. 6/22/05
    'Open the file with the web server status information and set the variables

    Dim slLocation As String
    Dim hlFrom As Integer
    Dim ilRet  As Integer
    Dim ilLen As Integer
    Dim ilPos As Integer
    Dim slTemp As String
    Dim llCount As Long
    
    On Error GoTo ErrHand
    mProcessWebWorkStatusResults = False
    Call gLoadOption(sgWebServerSection, sIniValue, slLocation)
    slLocation = gSetPathEndSlash(slLocation, True)
    slLocation = slLocation & sFileName
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable to open the file." & " " & Format(Now, "mm-dd-yy")
        smStatus = "1"
        Exit Function
    End If
    'Skip past the header record
    ilRet = 0
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Close #hlFrom
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable read/input statement." & " " & Format(Now, "mm-dd-yy")
        smStatus = "1"
        Exit Function
    End If
    On Error GoTo ErrHand
    slTemp = smMsg2
    ilLen = Len(slTemp)
    ilPos = InStr(slTemp, ":") Or InStr(slTemp, ".")
    If ilPos > 0 Then
        llCount = Val(Mid$(slTemp, ilPos + 1, ilLen))
        'Start logging here
        gLogMsgWODT "W", hmResult, "   WorkStatus: " & smFileName & ", " & smStatus & ", " & smMsg1 & ", " & smMsg2 & ", " & smDTStamp & " " & Format(Now, "mm-dd-yy")
        If InStr(slTemp, "Total Spots Imported") Then
            If Trim$(sTypeExpected) <> "Wegener Updates" Then
                mProcessWebWorkStatusResults = False
                gLogMsgWODT "W", hmResult, "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Event Info Imported:" & " " & Format(Now, "mm-dd-yy")
                Exit Function
            End If
            If llCount <> lmExportWebCount Then
                gLogMsgWODT "W", hmResult, "Error Counts Not Matching: Local = " & lmExportWebCount & " Web = " & llCount & " " & Format(Now, "mm-dd-yy")
            End If
            lmTtlEventSpots = lmTtlEventSpots + lmExportWebCount
            lmWebTtlEventSpots = lmWebTtlEventSpots + llCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Records Processed:") And sTypeExpected = "WebSpots" Then
            If Trim$(sTypeExpected) <> "WebSpots" Then
                mProcessWebWorkStatusResults = False
                gLogMsgWODT "W", hmResult, "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Records Processed:" & " " & Format(Now, "mm-dd-yy")
                Exit Function
            End If
            If llCount <> lmExportWebCount Then
                gLogMsgWODT "W", hmResult, "Error Counts Not Matching: Local = " & lmExportWebCount & " Web = " & llCount & " " & Format(Now, "mm-dd-yy")
            End If
            lmWebTtlSpots = lmWebTtlSpots + llCount
            lmTotalAddSpotCount = lmTotalAddSpotCount + lmExportWebCount
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Records Processed:") And sTypeExpected = "TotalSpots" Then
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "Total Emails Sent:") Then
            If Trim$(sTypeExpected) <> "WebEmails" Then
                mProcessWebWorkStatusResults = False
                gLogMsgWODT "W", hmResult, "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: Total Emails Sent:" & " " & Format(Now, "mm-dd-yy")
                Exit Function
            End If
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
        If InStr(slTemp, "ReIndex Complete.") Then
            If Trim$(sTypeExpected) <> "ReIndex" Then
                mProcessWebWorkStatusResults = False
                gLogMsgWODT "W", hmResult, "Error: mProcessWebWorkStatusResults, Expecting: " & sTypeExpected & " Recv: ReIndex Complete." & " " & Format(Now, "mm-dd-yy")
                Exit Function
            End If
            mProcessWebWorkStatusResults = True
            Exit Function
        End If
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener -mProcessWebWorkStatusResults"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Exit Function
End Function

Private Function mInitFTP() As Boolean

    Dim slTemp As String
    Dim ilRet As Integer
    Dim slSection As String
    
    On Error GoTo ErrHand
    mInitFTP = False
    If igTestSystem <> True Then
        slSection = "Locations"
    Else
        slSection = "TestLocations"
    End If
    'Support for CSI_Utils FTP functions
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tmCsiFtpInfo.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tmCsiFtpInfo.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tmCsiFtpInfo.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tmCsiFtpInfo.sPWD)
    Call gLoadOption(sgWebServerSection, "WebExports", tmCsiFtpInfo.sSendFolder)
    Call gLoadOption(sgWebServerSection, "WebImports", tmCsiFtpInfo.sRecvFolder)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tmCsiFtpInfo.sServerDstFolder)
    Call gLoadOption(sgWebServerSection, "FTPExportDir", tmCsiFtpInfo.sServerSrcFolder)
    Call gLoadOption("slSection", "DBPath", tmCsiFtpInfo.sLogPathName)
    tmCsiFtpInfo.sLogPathName = Trim$(tmCsiFtpInfo.sLogPathName) & "\" & "Messages\FTPLog.txt"
    ilRet = csiFTPInit(tmCsiFtpInfo)
    sgStartupDirectory = CurDir$
    sgIniPathFileName = sgStartupDirectory & "\Affiliat.Ini"
    Call gLoadOption(sgWebServerSection, "FTPPort", slTemp)
    tgCsiFtpFileListing.nPort = CInt(slTemp)
    Call gLoadOption(sgWebServerSection, "FTPAddress", tgCsiFtpFileListing.sIPAddress)
    Call gLoadOption(sgWebServerSection, "FTPUID", tgCsiFtpFileListing.sUID)
    Call gLoadOption(sgWebServerSection, "FTPPWD", tgCsiFtpFileListing.sPWD)
    Call gLoadOption("slSection", "DBPath", tgCsiFtpFileListing.sLogPathName)
    Call gLoadOption(sgWebServerSection, "FTPImportDir", tgCsiFtpFileListing.sPathFileMask)
    mInitFTP = True
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mInitFTP"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Exit Function
End Function

Private Function CurDrive() As String
  
  Dim slTemp As String
  
  slTemp = CurDir
  smCurDrive = Left$(slTemp, InStr(slTemp, ":"))
  CurDrive = smCurDrive
End Function

Private Sub SetResults(Msg As String, FGC As Long)

    Dim ilLen As Integer
    Dim ilMaxLen As Integer
    Dim slTemp As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    
    If lbcMsg.ListCount > 0 Then
        For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
            slTemp = lbcMsg.List(ilLoop)
            'create horz. scrool bar if the text is wider than the list box
            ilLen = Me.TextWidth(slTemp)
            If Me.ScaleMode = vbTwips Then
                ilLen = ilLen / Screen.TwipsPerPixelX  ' if twips change to pixels
            End If
            If ilLen > ilMaxLen Then
                ilMaxLen = ilLen
            End If
        Next ilLoop
        SendMessageByNum lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, ilMaxLen + 250, 0
    End If
    
End Sub


Private Function mFindEventCmmd(slSportInfo As String, slAirDate As String, ilVefCode As Integer, llGsfCode As Long, slDay As String) As Integer
    Dim slExportID As String
    Dim ilPos As Integer
    Dim slEventNo As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mFindEventCmmd = False
    ilPos = InStr(1, slSportInfo, ":", vbBinaryCompare)
    If ilPos <= 0 Then
        Exit Function
    End If
    slExportID = Left(slSportInfo, ilPos - 1)
    ilVefCode = mFindVefCode(slExportID)
    If ilVefCode = -1 Then
        Exit Function
    End If
    slEventNo = Mid(slSportInfo, ilPos + 1)
    SQLQuery = "SELECT * FROM gsf_Game_Schd WHERE gsfVefCode = " & ilVefCode & " AND gsfAirDate >= '" & Format(DateAdd("d", -1, slAirDate), sgSQLDateForm) & "' AND  gsfAirDate <= '" & Format(DateAdd("d", 1, slAirDate), sgSQLDateForm) & "'" & " AND gsfGameNo = " & slEventNo
    Set gsf_rst = gSQLSelectCall(SQLQuery)
    If gsf_rst.EOF Then
        Exit Function
    End If
    llGsfCode = gsf_rst!gsfCode
    slDay = Left(Format(gsf_rst!gsfAirDate, "ddd"), 2)
    mFindEventCmmd = True
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mFindEvent"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    mFindEventCmmd = False
    Exit Function
End Function

Private Sub mnuFakeWeb_Click()
    If mnuFakeWeb.Checked Then
        mnuFakeWeb.Checked = False
        bmFakeWeb = False
    Else
        mnuFakeWeb.Checked = True
        bmFakeWeb = True
    End If
End Sub

Private Sub mnuHaltWeb_Click()
    If mnuHaltWeb.Checked Then
        mnuHaltWeb.Checked = False
        bmHaltWeb = False
    Else
        mnuHaltWeb.Checked = True
        bmHaltWeb = True
    End If

End Sub

Private Function mFindEndPlay(slPlaylistName As String, slExportID As String, ilVefCode As Integer, ilBreakNo As Integer, slDay As String)
    mFindEndPlay = False
    If Len(slPlaylistName) < 6 Then
        Exit Function
    End If
    slExportID = ""
    If Mid(slPlaylistName, Len(slPlaylistName) - 2, 2) = "BK" Then
        ilBreakNo = Val(right(slPlaylistName, 1))
        slDay = Mid(slPlaylistName, Len(slPlaylistName) - 4, 2)
        slExportID = Left(slPlaylistName, Len(slPlaylistName) - 5)
    ElseIf Mid(slPlaylistName, Len(slPlaylistName) - 3, 2) = "BK" Then
        ilBreakNo = Val(right(slPlaylistName, 2))
        slDay = Mid(slPlaylistName, Len(slPlaylistName) - 5, 2)
        slExportID = Left(slPlaylistName, Len(slPlaylistName) - 6)
    End If
    If slExportID <> "" Then
        ilVefCode = mFindVefCode(slExportID)
        If ilVefCode <> -1 Then
            mFindEndPlay = True
        End If
    End If
End Function

Private Function mFindFile(sFolderName As String) As Integer
    
    'D.S. 9/17/19
    Dim fso As New FileSystemObject
    Dim fil As file
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mFindFile = 0
    ReDim smCompelFileNames(0 To 0)
    For Each fil In fso.GetFolder(sFolderName).Files
        smCompelFileNames(UBound(smCompelFileNames)) = fil.Name
        ReDim Preserve smCompelFileNames(0 To UBound(smCompelFileNames) + 1)
    Next
    mFindFile = UBound(smCompelFileNames)
    Set fso = Nothing
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mFindFile"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Set fso = Nothing
    Exit Function
End Function

Private Function mMoveFileAndRename(sFileName As String) As Boolean
    
    'D.S. 9/17/19
    Dim fso As New FileSystemObject
    Dim filesys
    Dim slFileBaseName As String
    Dim slFileName As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mMoveFileAndRename = False
    slFileName = smCompelImpPath & "\" & sFileName
    fso.MoveFile slFileName, smCompelSavePath & "\"
    slFileName = smCompelSavePath & "\" & sFileName
    slFileBaseName = fso.GetBaseName(sFileName)
    fso.MoveFile slFileName, smCompelSavePath & "\" & slFileBaseName & ".sav"
    mMoveFileAndRename = True
    Set fso = Nothing
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mMoveFileAndRename"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Set fso = Nothing
    Exit Function
End Function

Private Function mDeleteFilesByDate(sFolderName As String, sDaysToRetain As String) As Boolean
    
    'D.S. 9/17/19
    Dim fso As New FileSystemObject
    Dim fil As file
    Dim slFileDate As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mDeleteFilesByDate = False
    For Each fil In fso.GetFolder(sFolderName).Files
        slFileDate = fil.DateLastModified
        If DateDiff("d", slFileDate, Now) > CInt(sDaysToRetain) Then
            fso.DeleteFile fil, True
        End If
    Next
    Set fso = Nothing
    mDeleteFilesByDate = True
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mDeleteFilesByDate"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Set fso = Nothing
    Exit Function
End Function

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "CAI"
End Sub

Private Function mWriteCompelFile() As Boolean
    Dim llLoop As Long
    Dim ilRet As Integer
    Dim slVefName As String
    Dim slStnCall As String
    Dim slGame As String
    Dim slAirDt As String
    Dim slAirTm As String
    Dim slAdvNm As String
    Dim slISCI As String
    Dim slStatus As String
    Dim ilBrkNo As Integer
    Dim llRow As Long
    Dim hlToFile As Integer
    Dim slFileName As String
    Dim slStr As String
    
    On Error GoTo ErrHand
    mWriteCompelFile = False
    hlToFile = FreeFile
    slFileName = smCompelSavePath & "\" & slCompelBaseName & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & "_SpotsOnly" & ".csv"
    ilRet = gFileOpen(slFileName, "Output Lock Write", hlToFile)
    If ilRet <> 0 Then
        gLogMsgWODT "W", hmResult, "Unable to open file " & slFileName & vbCrLf & "Compel Auto Import Open File Failed" & " " & Format(Now, "mm-dd-yy")
        Exit Function
    End If
    
    Print #hlToFile, "VefName, CallLetters, Game, AirDate, AirTime, AdvName, ISCI, Status, BreakNo, Row"
    
    For llLoop = 0 To lmExptCmpCnt - 1
        slVefName = tmCompelExportInfo(llLoop).sVefName
        slStnCall = tmCompelExportInfo(llLoop).sCallLetters
        slGame = tmCompelExportInfo(llLoop).sGame
        slAirDt = tmCompelExportInfo(llLoop).sAiredDate
        slAirTm = tmCompelExportInfo(llLoop).sAiredTime
        slAdvNm = tmCompelExportInfo(llLoop).sAdvertiser
        slISCI = tmCompelExportInfo(llLoop).sISCI
        slStatus = tmCompelExportInfo(llLoop).sStatus
        ilBrkNo = tmCompelExportInfo(llLoop).ilBreakNo
        llRow = tmCompelExportInfo(llLoop).lRow
        slStr = slVefName & "," & slStnCall & "," & slGame & "," & slAirDt & "," & slAirTm & "," & slAdvNm & "," & slISCI & "," & slStatus & "," & ilBrkNo & "," & llRow
        Print #hlToFile, slStr
    Next llLoop
    mWriteCompelFile = True
    lmExptCmpCnt = 0
    ReDim tmCompelExportInfo(0 To 0) As COMPELEXPORTINFO
    Close #hlToFile
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mWriteCompelFile"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    Close #hlToFile
    Exit Function
End Function

Private Function mGetGameTeamNames(lGameCode As Long) As String
    Dim gsf_rst As ADODB.Recordset
    Dim mnf_rst As ADODB.Recordset
    Dim slHomeTeam As String
    Dim slVisitTeam As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    mGetGameTeamNames = ""
    SQLQuery = "SELECT gsfHomeMnfCode, gsfVisitMnfCode FROM GSF_Game_Schd WHERE gsfCode = " & lGameCode
    Set gsf_rst = gSQLSelectCall(SQLQuery)
    SQLQuery = "select mnfName from MNF_Multi_Names where mnfCode = " & gsf_rst!gsfHomeMnfCode
    Set mnf_rst = gSQLSelectCall(SQLQuery)
    slHomeTeam = mnf_rst!mnfName
    SQLQuery = "select mnfName from MNF_Multi_Names where mnfCode = " & gsf_rst!gsfVisitMnfCode
    Set mnf_rst = gSQLSelectCall(SQLQuery)
    slVisitTeam = mnf_rst!mnfName
    gsf_rst.Close
    mnf_rst.Close
    mGetGameTeamNames = Trim(slHomeTeam) & " vs. " & Trim(slVisitTeam)
    Exit Function
ErrHand:
    gHandleError "WegenerImportResult_" & Format(Now, "mm-dd-yy") & ".txt", "Import Wegener-mGetGameTeamNames"
    ilRet = gAlertAdd("U", "C", 0, Format(Now, "ddddd"))
    gsf_rst.Close
    mnf_rst.Close
    Exit Function
End Function

